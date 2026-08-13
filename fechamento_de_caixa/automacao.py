import os
import time
import subprocess
from tkinter import Tk, filedialog
import pandas as pd
from pywinauto import Application

def extrair_dados_do_nome(nome_arquivo):
    """Extrai cartão, posto e data baseando-se no padrão do nome do arquivo."""
    nome_sem_extensao, _ = os.path.splitext(nome_arquivo)
    partes = nome_sem_extensao.split(" ")
    
    cartao = "Goodcard" if "Goodcard" in nome_sem_extensao else "Não identificado"
    
    posto = "Não identificado"
    if "Centro Automotivo" in nome_sem_extensao:
        posto = "Centro Automotivo"
    elif "Arquipelago" in nome_sem_extensao:
        posto = "Posto Arquipelago"
    elif "Mareli" in nome_sem_extensao:
        posto = "Posto Mareli"
        
    data_lancamento = partes[-1] if partes else "Não identificada"
    return cartao, posto, data_lancamento

def abrir_via_libreoffice(caminho_arquivo, caminho_novo):
    """Caso o pandas falhe, usa o LibreOffice Calc via terminal para salvar como .xlsx limpo."""
    print("Tentando converter o arquivo usando o LibreOffice Calc...")
    
    # Caminhos mais comuns onde o LibreOffice fica instalado no Windows
    caminhos_libreoffice = [
        r"C:\Program Files\LibreOffice\program\scalc.exe",
        r"C:\Program Files (x86)\LibreOffice\program\scalc.exe"
    ]
    
    scalc_path = None
    for caminho in caminhos_libreoffice:
        if os.path.exists(caminho):
            scalc_path = caminho
            break
            
    if not scalc_path:
        print("❌ O LibreOffice Calc não foi encontrado nas pastas padrão do Windows.")
        return False
        
    try:
        pasta_destino = os.path.dirname(caminho_arquivo)
        # Comando do LibreOffice para converter arquivos em background de forma invisível
        comando = f'"{scalc_path}" --headless --convert-to xlsx "{caminho_arquivo}" --outdir "{pasta_destino}"'
        
        # Executa a conversão do sistema
        subprocess.run(comando, shell=True, check=True)
        
        # O LibreOffice cria um arquivo com o mesmo nome, só mudando para .xlsx
        nome_base, _ = os.path.splitext(caminho_arquivo)
        arquivo_convertido = nome_base + ".xlsx"
        
        if os.path.exists(arquivo_convertido):
            # Agora abre o arquivo novo e coloca a coluna Conferido
            df = pd.read_excel(arquivo_convertido)
            df.columns = [str(col).strip() for col in df.columns]
            df['Conferido'] = ""
            
            # Reorganiza para colocar 'Conferido' na primeira coluna
            outras = [c for c in df.columns if c != 'Conferido']
            df = df[['Conferido'] + outras]
            
            # Salva por cima com o nome correto terminado em _conferido.xlsx
            df.to_excel(caminho_novo, index=False)
            
            # Apaga o arquivo temporário convertido sem o sufixo
            os.remove(arquivo_convertido)
            return True
            
    except Exception as e:
        print(f"❌ Falha ao processar pelo LibreOffice: {e}")
    return False

import openpyxl

import openpyxl

def processar_planilha_seguro(caminho_arquivo):
    print("Modificando o arquivo original e inserindo uma nova coluna...")
    
    # 1. Converte .xls para .xlsx usando o LibreOffice se necessário
    if caminho_arquivo.endswith('.xls'):
        print("⚠️ Convertendo formato antigo .xls para aceitar modificações nativas...")
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
            print("❌ LibreOffice não encontrado para converter o arquivo.")
            return False
    else:
        caminho_trabalho = caminho_arquivo

    try:
        # 2. Abre a planilha original preservando o layout
        wb = openpyxl.load_workbook(caminho_trabalho)
        ws = wb.active
        
        # 3. Procura a linha onde está a palavra "CNPJ" para saber onde alinhar
        linha_cnpj = None
        for row in range(1, 20):
            # Procura na coluna A ou B original antes de mexer na estrutura
            val_a = str(ws.cell(row=row, column=1).value).strip()
            val_b = str(ws.cell(row=row, column=2).value).strip()
            if val_a == "CNPJ" or val_b == "CNPJ":
                linha_cnpj = row
                break
                
        if not linha_cnpj:
            linha_cnpj = 6 # Caso não encontre por algum motivo, assume a linha 6 do padrão
            
        # 4. CRIA A NOVA COLUNA À ESQUERDA (Coluna A vira uma nova coluna limpa)
        # O comando insert_cols(1) cria uma coluna vazia na posição 1 (Coluna A) e empurra o resto para a direita
        ws.insert_cols(1)
        
        # 5. Escreve "Conferido" na nova Coluna A, exatamente na linha do CNPJ
        ws.cell(row=linha_cnpj, column=1).value = "Conferido"
        print(f"✨ Nova coluna criada! 'Conferido' adicionado na célula A{linha_cnpj}.")
            
        # 6. Salva e altera o nome do arquivo incluindo o '_conferido'
        pasta = os.path.dirname(caminho_trabalho)
        nome_completo = os.path.basename(caminho_trabalho)
        nome_sem_ext, _ = os.path.splitext(nome_completo)
        
        if nome_sem_ext.endswith("_conferido"):
            nome_sem_ext = nome_sem_ext.replace("_conferido", "")
            
        caminho_final = os.path.join(pasta, f"{nome_sem_ext}_conferido.xlsx")
        
        wb.save(caminho_trabalho)
        wb.close()
        
        # Altera o nome do arquivo fisicamente no Windows
        if os.path.exists(caminho_trabalho):
            if os.path.exists(caminho_final):
                os.remove(caminho_final) # Remove se já existir um teste antigo com o mesmo nome
            os.rename(caminho_trabalho, caminho_final)
            
        print(f"✅ Arquivo modificado com sucesso: {os.path.basename(caminho_final)}")
        return True
        
    except Exception as e:
        print(f"❌ Erro ao processar arquivo: {e}")
        return False

def selecionar_e_processar_projeto():
    print("Aguardando seleção do arquivo...")
    
    root = Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    
    caminho_arquivo = filedialog.askopenfilename(
        title="Selecione o arquivo de Fechamento de Caixa",
        filetypes=[("Arquivos de Fechamento", "*.xls *.xlsx")]
    )
    
    if not caminho_arquivo:
        print("Nenhum arquivo foi selecionado. Processo cancelado.")
        return None
    
    caminho_arquivo = os.path.abspath(caminho_arquivo)
    nome_arquivo = os.path.basename(caminho_arquivo)
    print(f"Arquivo selecionado: {nome_arquivo}")
    
    cartao, posto, data_lancamento = extrair_dados_do_nome(nome_arquivo)
    
    print(f"\n[DADOS CAPTURADOS DO NOME]")
    print(f"💳 Cartão: {cartao}")
    print(f"⛽ Posto: {posto}")
    print(f"📅 Data: {data_lancamento}\n")
    
    processar_planilha_seguro(caminho_arquivo)
    return {"cartao": cartao, "posto": posto, "data": data_lancamento}

def abrir_e_logar_webposto():
    # 1. Executa a seleção e modificação da planilha
    dados_caixa = selecionar_e_processar_projeto()
    if not dados_caixa:
        return
        
    print("\nIniciando a automação do webPosto...")
    caminho_programa = r"C:\Quality\web\QualityPosto.exe"
    
    try:
        # 2. Abre o programa e faz o login
        app = Application(backend="win32").start(caminho_programa)
        print("Programa aberto. Aguardando a tela de login carregar...")
        time.sleep(5)
        
        janela_login = app.top_window()
        print("Forçando a janela a aparecer no primeiro plano...")
        janela_login.set_focus()
        
        print("Digitando o usuário...")
        janela_login.click_input()
        janela_login.type_keys("^a{BACKSPACE}GESSICAMARELI", set_foreground=True)
        janela_login.type_keys("{TAB}")
        
        print("Digitando a senha...")
        janela_login.type_keys("12345")
        
        print("Pressionando Enter para entrar...")
        janela_login.type_keys("{ENTER}")
        
        print("Automação de login concluída!")
        
        # 3. Aguardar a tela de seleção de posto carregar
        print("Aguardando a tela de seleção de postos carregar...")
        time.sleep(4)
        
        janela_principal = app.top_window()
        janela_principal.set_focus()
        
        # 4. LÓGICA DE SELEÇÃO DO POSTO
        posto_detectado = dados_caixa['posto']
        print(f"Identificado no arquivo: '{posto_detectado}'.")
        
        # Mapeia quantas vezes o robô precisa apertar a 'seta para baixo' para chegar no posto certo
        # Olhando a imagem: Centro = 0 (já começa nele), Mareli = 1, Arquipelago = 2
        pulos_seta = 0
        if posto_detectado == "Centro Automotivo":
            pulos_seta = 0
        elif posto_detectado == "Posto Mareli":
            pulos_seta = 1
        elif posto_detectado == "Posto Arquipelago":
            pulos_seta = 2
            
        print(f"Navegando na lista de filiais...")
        import pyautogui
        
        # Dá um clique físico no meio da lista para garantir que o teclado está ativo ali dentro
        pyautogui.click(x=janela_principal.rectangle().left + 200, y=janela_principal.rectangle().top + 180)
        time.sleep(0.5)
        
        # Desce na lista o número de vezes necessário
        for _ in range(pulos_seta):
            pyautogui.press('down')
            time.sleep(0.3)

        # 5. CLIQUE NO BOTÃO VERMELHO "SELECIONAR"
        print("Confirmando a seleção no botão vermelho 'Selecionar'...")
        pyautogui.press('enter')
        print("✅ Comando de entrada enviado com sucesso!")
        
        # 6. TRATAMENTO DO POP-UP "NOVIDADES WEBPOSTO"
        print("Aguardando carregamento da tela principal...")
        time.sleep(5)
        
        janela_principal = app.top_window()
        janela_principal.set_focus()
        
        # Sequência de fechamento dos pop-ups (Check verde + ESC)
        try:
            rect = janela_principal.rectangle()
            posicao_x_check = rect.left + 25
            posicao_y_check = rect.bottom - 20
            
            for i in range(3):
                pyautogui.click(x=posicao_x_check, y=posicao_y_check)
                time.sleep(0.8)
        except:
            pass
            
        pyautogui.press('esc')
        time.sleep(1) # Aguarda 1 segundo para garantir que a tela de trás reaja
        
        # 7. CLICAR EM FINANCEIRO E FECHAMENTO DE CAIXA
        print("Abrindo o menu Financeiro via clique posicional...")
        
        janela_principal = app.top_window()
        janela_principal.set_focus()
        time.sleep(0.5)
        
        rect_main = janela_principal.rectangle()
        
        # 1º Clique: Abre a aba Financeiro
        posicao_x_financeiro = rect_main.left + 265
        posicao_y_financeiro = rect_main.top + 45
        pyautogui.click(x=posicao_x_financeiro, y=posicao_y_financeiro)
        time.sleep(1.5)
        
        print("Clicando no botão 'Fechamento de Caixa'...")
        # 2º Clique: Abre a janela interna do Fechamento
        posicao_x_fechamento = rect_main.left + 690
        posicao_y_fechamento = rect_main.top + 105
        pyautogui.click(x=posicao_x_fechamento, y=posicao_y_fechamento)
        
        # 8. PREENCHER AS DATAS E CONSULTAR (MÉTODO VISÃO ÓPTICA - INDEPENDENTE DE MONITOR)
        print("Aguardando a janela do Fechamento de Caixa carregar...")
        time.sleep(4) 
        
        data_para_preencher = dados_caixa['data'].replace("-", "").replace("/", "")
        print(f"Preparando digitação da data: {data_para_preencher}")
        
        import pyautogui
        from PIL import Image
        import pytesseract
        
        # Descobre onde o script do projeto está rodando no computador atual
        diretorio_projeto = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

        # Aponta para a pasta do Tesseract que está dentro do próprio projeto
        pytesseract.pytesseract.tesseract_cmd = os.path.join(diretorio_projeto, "Tesseract-OCR", "tesseract.exe")

        
        print("Buscando as palavras 'Data' na tela visualmente...")
        try:
            # 1. Tira uma foto instantânea da tela inteira
            foto_tela = pyautogui.screenshot()
            
            # 2. O Python lê a foto e descobre as coordenadas de todas as palavras da tela
            dados_leitura = pytesseract.image_to_data(foto_tela, output_type=pytesseract.Output.DICT)
            
            pontos_data = []
            
            # Varre todas as palavras encontradas na tela para achar onde está escrito "Data"
            for i, texto in enumerate(dados_leitura['text']):
                if "Data" in texto:
                    # Captura a posição física X e Y da palavra na sua tela
                    x = dados_leitura['left'][i] + (dados_leitura['width'][i] // 2)
                    y = dados_leitura['top'][i] + (dados_leitura['height'][i] // 2)
                    pontos_data.append((x, y))
            
            # Baseado no layout da sua imagem, as duas primeiras vezes que a palavra "Data" 
            # aparece no bloco cinza superior correspondem a 'Data Inicio' e 'Data Fim'
            if len(pontos_data) >= 2:
                print(f"✅ Textos de data localizados visualmente na tela!")
                
                # --- Preenchendo Data Inicio ---
                x_inicio, y_inicio = pontos_data[0]
                # SUA IDÉIA: Anda 90 pixels para a direita a partir do texto lido
                pyautogui.moveTo(x_inicio + 90, y_inicio, duration=0.4)
                pyautogui.click(clicks=2, interval=0.2) # Duplo clique que você descobriu
                time.sleep(0.3)
                pyautogui.hotkey('ctrl', 'a')
                pyautogui.press('backspace')
                pyautogui.write(data_para_preencher, interval=0.05)
                pyautogui.press('esc')
                time.sleep(0.5)
                
                # --- Preenchendo Data Fim ---
                x_fim, y_fim = pontos_data[1]
                pyautogui.moveTo(x_fim + 90, y_fim, duration=0.4)
                pyautogui.click(clicks=2, interval=0.2)
                time.sleep(0.3)
                pyautogui.hotkey('ctrl', 'a')
                pyautogui.press('backspace')
                pyautogui.write(data_para_preencher, interval=0.05)
                pyautogui.press('esc')
                time.sleep(0.5)
                
            else:
                print("❌ Não foi possível ler a palavra 'Data' na tela. Verifique se a janela está aberta.")
                
            # --- 3. LOCALIZAR O BOTÃO CONSULTAR PELO TEXTO ---
            print("Buscando o botão 'Consultar' visualmente...")
            for i, texto in enumerate(dados_leitura['text']):
                if "Consultar" in texto or "onsultar" in texto:
                    x_btn = dados_leitura['left'][i] + (dados_leitura['width'][i] // 2)
                    y_btn = dados_leitura['top'][i] + (dados_leitura['height'][i] // 2)
                    
                    pyautogui.moveTo(x_btn, y_btn, duration=0.4)
                    pyautogui.click()
                    print("✅ Botão Consultar acionado com sucesso visualmente!")
                    break
                    
        except Exception as e_visao:
            print(f"⚠️ Falha no motor de leitura óptica da tela: {e_visao}")
            print("Verifique se o Tesseract foi instalado corretamente no Passo 2.")
            
    except Exception as e:
        print(f"❌ Ocorreu um erro na integração com a tela do webPosto: {e}")

if __name__ == "__main__":
    abrir_e_logar_webposto()



