import os
import time
import subprocess
from tkinter import Tk, filedialog
import openpyxl
from pywinauto import Application
import pyautogui
from PIL import Image
import pytesseract
import pandas as pd
import cv2
import numpy as np

def localizar_e_clicar_imagem(nome_imagem, cliques=1, arrasto_x=0, arrasto_y=0):
    diretorio_projeto = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    caminho_real_imagem = os.path.join(diretorio_projeto, nome_imagem)
    
    print_tela = pyautogui.screenshot()
    imagem_tela = cv2.cvtColor(np.array(print_tela), cv2.COLOR_RGB2BGR)
    imagem_alvo = cv2.imread(caminho_real_imagem)
    
    if imagem_alvo is None:
        return None
        
    resultado = cv2.matchTemplate(imagem_tela, imagem_alvo, cv2.TM_CCOEFF_NORMED)
    _, max_val, _, max_loc = cv2.minMaxLoc(resultado)
    
    if max_val >= 0.8:
        h, w = imagem_alvo.shape[:2]
        centro_x = max_loc[0] + (w // 2) + arrasto_x
        centro_y = max_loc[1] + (h // 2) + arrasto_y
        
        pyautogui.moveTo(centro_x, centro_y, duration=0.4)
        pyautogui.click(clicks=cliques, interval=0.2)
        return (centro_x, centro_y)
    return None

def extrair_dados_do_nome(nome_arquivo):
    """Extrai cartão, posto e data baseando-se no padrão do nome do arquivo."""
    nome_sem_extensao, _ = os.path.splitext(nome_arquivo)
    partes = nome_sem_extensao.split(" ")
    cartao = "Goodcard" if "Goodcard" in nome_sem_extensao else "Não identificado"
    posto = "Não identificado"
    if "Centro Automotivo" in nome_sem_extensao:
        posto = "Centro Automotivo"
    elif "Arquipelago" in nome_sem_extensao:
        posto = "Posto Archipelago"
    elif "Mareli" in nome_sem_extensao:
        posto = "Posto Mareli"
    data_lancamento = partes[-1] if partes else "Não identificada"
    return cartao, posto, data_lancamento

def processar_planilha_seguro(caminho_arquivo):
    print("[PLANILHA] Verificando e estruturando colunas da planilha selecionada...")
    if caminho_arquivo.endswith('.xls'):
        pasta_destino = os.path.dirname(caminho_arquivo)
        nome_base, _ = os.path.splitext(caminho_arquivo)
        caminho_trabalho = nome_base + ".xlsx"
        caminhos_libreoffice = [
            r"C:\Program Files\LibreOffice\program\scalc.exe",
            r"C:\Program Files (x86)\LibreOffice\program\scalc.exe"
        ]
        scalc_path = next((c for c in caminhos_libreoffice if os.path.exists(c)), None)
        if scalc_path:
            comando = f'"{scalc_path}" --headless --convert-to xlsx "{caminho_arquivo}" --outdir "{pasta_destino}"'
            subprocess.run(comando, shell=True, check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            try: os.remove(caminho_arquivo) 
            except: pass
        else:
            return False
    else:
        caminho_trabalho = caminho_arquivo
    try:
        wb = openpyxl.load_workbook(caminho_trabalho)
        ws = wb.active
        linha_cnpj = None
        for row in range(1, 20):
            val_a = str(ws.cell(row=row, column=1).value).strip()
            val_b = str(ws.cell(row=row, column=2).value).strip()
            if val_a == "CNPJ" or val_b == "CNPJ":
                linha_cnpj = row
                break
        if not linha_cnpj:
            linha_cnpj = 6
        ws.insert_cols(1)
        ws.cell(row=linha_cnpj, column=1).value = "Conferido"
        pasta = os.path.dirname(caminho_trabalho)
        nome_completo = os.path.basename(caminho_trabalho)
        nome_sem_ext, _ = os.path.splitext(nome_completo)
        if nome_sem_ext.endswith("_conferido"):
            nome_sem_ext = nome_sem_ext.replace("_conferido", "")
        caminho_final = os.path.join(pasta, f"{nome_sem_ext}_conferido.xlsx")
        wb.save(caminho_trabalho)
        wb.close()
        if os.path.exists(caminho_trabalho):
            if os.path.exists(caminho_final):
                os.remove(caminho_final)
            os.rename(caminho_trabalho, caminho_final)
        print(f"[SUCESSO] Planilha editada e salva como: {os.path.basename(caminho_final)}")
        return True
    except Exception as e:
        print(f"[ERRO] Falha crítica ao manipular as células da planilha: {e}")
        return False

def selecionar_e_processar_projeto():
    print("Aguardando seleção do arquivo pelo usuário...")
    root = Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    caminho_arquivo = filedialog.askopenfilename(
        title="Selecione o arquivo de Fechamento de Caixa",
        filetypes=[("Arquivos de Fechamento", "*.xls *.xlsx")]
    )
    if not caminho_arquivo:
        return None
    caminho_arquivo = os.path.abspath(caminho_arquivo)
    nome_arquivo = os.path.basename(caminho_arquivo)
    cartao, posto, data_lancamento = extrair_dados_do_nome(nome_arquivo)
    print(f"\n[DADOS CAPTURADOS DO NOME]")
    print(f"💳 Cartão: {cartao}")
    print(f"⛽ Posto: {posto}")
    print(f"📅 Data: {data_lancamento}\n")
    processar_planilha_seguro(caminho_arquivo)
    return {"cartao": cartao, "posto": posto, "data": data_lancamento}

def abrir_e_logar_webposto():
    dados_caixa = selecionar_e_processar_projeto()
    if not dados_caixa:
        return
    print("\n[SISTEMA] Iniciando a automação do sistema webPosto...")
    caminho_programa = r"C:\Quality\web\QualityPosto.exe"
    try:
        print(f"[SISTEMA] Executando o arquivo do programa em: {caminho_programa}")
        app = Application(backend="win32").start(caminho_programa)
        print("[SISTEMA] Aguardando 5 segundos para o carregamento da tela de login...")
        time.sleep(5)
        
        print("[SISTEMA] Capturando a janela de credenciais e forçando o primeiro plano...")
        janela_login = app.top_window()
        janela_login.set_focus()
        janela_login.click_input()
        
        print("[ROBÔ] Digitando o usuário: GESSICAMARELI")
        janela_login.type_keys("^a{BACKSPACE}GESSICAMARELI", set_foreground=True)
        janela_login.type_keys("{TAB}")
        
        print("[ROBÔ] Digitando a senha cadastrada...")
        janela_login.type_keys("12345")
        
        print("[ROBÔ] Enviando comando ENTER para validar o acesso...")
        janela_login.type_keys("{ENTER}")
        print("[SUCESSO] Automação de login concluída!")
        
        print("[SISTEMA] Conectando ao motor óptico Tesseract do projeto...")
        diretorio_projeto = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        pytesseract.pytesseract.tesseract_cmd = os.path.join(diretorio_projeto, "Tesseract-OCR", "tesseract.exe")
        
        print("[SISTEMA] Aguardando 4 segundos para a abertura da listagem de filiais...")
        time.sleep(4)
        janela_principal = app.top_window()
        janela_principal.set_focus()
        
        posto_detectado = dados_caixa['posto']
        pulos_seta = 0
        if posto_detectado == "Centro Automotivo": pulos_seta = 0
        elif posto_detectado == "Posto Mareli": pulos_seta = 1
        elif posto_detectado == "Posto Archipelago": pulos_seta = 2
            
        print(f"[ROBÔ] Ativando foco no grid de postos da tela vermelha...")
        pyautogui.click(x=janela_principal.rectangle().left + 200, y=janela_principal.rectangle().top + 180)
        time.sleep(0.5)
        
        print(f"[ROBÔ] Pressionando a SETA PARA BAIXO {pulos_seta} vezes para selecionar o posto correto...")
        for _ in range(pulos_seta):
            pyautogui.press('down')
            time.sleep(0.3)
            
        print("[ROBÔ] Pressionando ENTER para acionar o botão vermelho 'Selecionar'...")
        pyautogui.press('enter')
        print("[SUCESSO] Posto confirmado na interface do sistema!")
        
        print("[SISTEMA] Aguardando 8 segundos para a carga completa da tela inicial...")
        time.sleep(8)
        
        janela_principal = app.top_window()
        janela_principal.set_focus()
        
        print("[POP-UP] Tratando possíveis pop-ups iniciais...")
        try:
            rect = janela_principal.rectangle()
            for _ in range(3):
                pyautogui.click(x=rect.left + 25, y=rect.bottom - 20)
                time.sleep(0.8)
        except: pass
        pyautogui.press('esc')
        time.sleep(1.5)
        
        # --- BUSCA POR IMAGEM REFINADA COM O SEU NOVO PRINT CINZA ---
        print("[VISÃO] Procurando a aba 'Financeiro' (fundo cinza normal) no menu superior...")
        if localizar_e_clicar_imagem('aba_financeiro_menu.png', cliques=1):
            print("[SUCESSO] Aba Financeiro clicada com precisão de imagem!")
            time.sleep(2) # Aguarda a barra expandir para o lado de baixo
        else:
            print("[ERRO] Não encontrou a imagem 'aba_financeiro_menu.png'. Certifique-se de salvá-la na pasta smartposto.")
            return
            
        print("[VISÃO] Procurando o botão 'Fechamento de Caixa' na barra de ferramentas...")
        if localizar_e_clicar_imagem('fechamento_de_caixa.png', cliques=1):
            print("[SUCESSO] Janela de Fechamento de Caixa disparada!")
        else:
            print("[SISTEMA] Usando redundância de texto para o Fechamento de Caixa...")
            foto_ribbon = pyautogui.screenshot()
            dados_ribbon = pytesseract.image_to_data(foto_ribbon, output_type=pytesseract.Output.DICT)
            x_fc, y_fc = None, None
            for i, texto in enumerate(dados_ribbon['text']):
                texto_limpo = texto.strip().lower()
                if "fechamento" in texto_limpo or "caixa" in texto_limpo:
                    y_atual = dados_ribbon['top'][i] + (dados_ribbon['height'][i] // 2)
                    if y_atual < 200: 
                        x_fc = dados_ribbon['left'][i] + (dados_ribbon['width'][i] // 2)
                        y_fc = y_atual
                        break
            if x_fc and y_fc:
                pyautogui.moveTo(x_fc, y_fc, duration=0.5)
                pyautogui.click()
            else:
                print("[ERRO] Não foi possível localizar o botão do Fechamento de Caixa.")
                return

        print("[SISTEMA] Aguardando 4 segundos para o carregamento da janela interna...")
        time.sleep(4) 
        data_para_preencher = dados_caixa['data'].replace("-", "").replace("/", "")
        print(f"[SISTEMA] String de data formatada para envio: {data_para_preencher}")
        
        # --- PASSO 1: PREENCHER AS DATAS DE FORMA SEGURA ---
        ponto_inicio = localizar_e_clicar_imagem('campo_data.png', cliques=2)
        if ponto_inicio:
            x_ini, y_ini = ponto_inicio
            pyautogui.hotkey('ctrl', 'a')
            pyautogui.press('backspace')
            pyautogui.write(data_para_preencher, interval=0.05)
            pyautogui.press('esc')
            time.sleep(0.5)
            
            pyautogui.moveTo(x_ini, y_ini + 35, duration=0.4)
            pyautogui.click(clicks=2, interval=0.2)
            time.sleep(0.3)
            pyautogui.hotkey('ctrl', 'a')
            pyautogui.press('backspace')
            pyautogui.write(data_para_preencher, interval=0.05)
            pyautogui.press('esc')
            time.sleep(0.5)
        else:
            print("[ERRO] Não localizou o campo de data.")
            return

        # --- PASSO 2: CLICAR EM CONSULTAR ---
        # Salva as coordenadas do botão Consultar para usar de barreira vertical na tabela
        ponto_consultar = localizar_e_clicar_imagem('botao_consultar.png')
        if not ponto_consultar:
            print("[ERRO] Botão Consultar não localizado.")
            return
        y_btn = ponto_consultar[1]
        time.sleep(4)

        # --- PASSO 3: CLIQUE DUPLO NO 1º TURNO ---
        if not localizar_e_clicar_imagem('texto_1turno.png', cliques=2):
            print("[ERRO] Linha do 1º Turno não localizada.")
            return

        # --- PASSO 4: SELECIONAR A ABA CARTÃO ---
        time.sleep(5)
        if not localizar_e_clicar_imagem('aba_cartao.png'):
            print("[ERRO] Aba Cartão não localizada.")
            return
        time.sleep(2)

        # --- LOOP DE BUSCA EXCLUSIVA DA PRIMEIRA OCORRÊNCIA ---
        def executar_filtro_goodcard(tentativa=1):
            print(f"\n[FILTRO] Analisando a lista de administradoras (Página {tentativa})...")
            
            diretorio_projeto = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
            caminho_admin = os.path.join(diretorio_projeto, 'texto_administradora.png')
            print_tela = pyautogui.screenshot()
            img_tela = cv2.cvtColor(np.array(print_tela), cv2.COLOR_RGB2BGR)
            img_alvo = cv2.imread(caminho_admin)
            
            if img_alvo is not None:
                res = cv2.matchTemplate(img_tela, img_alvo, cv2.TM_CCOEFF_NORMED)
                _, max_val, _, max_loc = cv2.minMaxLoc(res)
                if max_val >= 0.8:
                    h, w = img_alvo.shape[:2]
                    pyautogui.moveTo(max_loc[0] + w + 15, max_loc[1] + (h // 2), duration=0.4)
                    time.sleep(0.6)
            
            if localizar_e_clicar_imagem('filtro_administradora.png', cliques=1):
                time.sleep(1.5)
                
                caminho_goodcard = os.path.join(diretorio_projeto, 'opcao_goodcard.png')
                print_lista = pyautogui.screenshot()
                img_lista = cv2.cvtColor(np.array(print_lista), cv2.COLOR_RGB2BGR)
                img_good = cv2.imread(caminho_goodcard)
                
                encontrou_goodcard = False
                if img_good is not None:
                    res_good = cv2.matchTemplate(img_lista, img_good, cv2.TM_CCOEFF_NORMED)
                    _, max_val_good, _, max_loc_good = cv2.minMaxLoc(res_good)
                    if max_val_good >= 0.8:
                        h_g, w_g = img_good.shape[:2]
                        pyautogui.moveTo(max_loc_good[0] + (w_g // 2), max_loc_good[1] + (h_g // 2), duration=0.4)
                        pyautogui.click()
                        print(f"[🎯 SUCESSO] GOODCARD encontrado e marcado na Página {tentativa}! O robô parou de avançar.")
                        encontrou_goodcard = True
                
                if not encontrou_goodcard:
                    print(f"[AVISO] GOODCARD não está na Página {tentativa}. Avançando...")
                    pyautogui.press('esc')
                    time.sleep(1)
                    
                    if localizar_e_clicar_imagem('botao_avancar_caixa.png', cliques=1):
                        print("[SISTEMA] Próxima página carregada. Reiniciando checagem...")
                        time.sleep(4)
                        executar_filtro_goodcard(tentativa = tentativa + 1)
                    else:
                        print("[ERRO] Botão Avançar não localizado.")
            else:
                print("[ERRO] Não foi possível abrir o funil nesta página.")

        # Dispara a busca controlada
        executar_filtro_goodcard()
                
    except Exception as e:
        print(f"[ERRO] Falha geral de execução: {e}")

if __name__ == "__main__":
    abrir_e_logar_webposto()
