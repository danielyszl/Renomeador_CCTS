import pdfplumber
import pytesseract
import re
import os
from tkinter import Tk, filedialog
from datetime import datetime
import comtypes.client
import shutil

gerar_log = True

def mover_para_erro(pasta_origem, nome_arquivo):
    pasta_erro = os.path.join(pasta_origem, "Arquivos_Nao_Processados")
    if not os.path.exists(pasta_erro):
        os.makedirs(pasta_erro)
    
    caminho_origem = os.path.join(pasta_origem, nome_arquivo)
    caminho_destino = os.path.join(pasta_erro, nome_arquivo)
    
    try:
        shutil.move(caminho_origem, caminho_destino) #move o arquivo original para a pasta de erro
        return True
    except Exception as e:
        print(f"-> Erro ao mover arquivo para pasta de erro: {e}")
        return False

def doc_para_pdf(caminho_entrada):
    
    try:
        word = comtypes.client.CreateObject('Word.Application')
        word.Visible = False
        abs_in = os.path.abspath(caminho_entrada)
        caminho_pdf = os.path.splitext(abs_in)[0] + ".pdf"
        
        doc = word.Documents.Open(abs_in)
        doc.SaveAs(caminho_pdf, FileFormat=17) #17 formato pdf
        doc.Close()
        word.Quit()
        return caminho_pdf
    
    except Exception as e:
        print(f"-> Erro na conversão para PDF: {e}")
        return None
    
def texto_legivel(texto):
    if not texto:
        return False
    texto = texto.strip()
    if texto.count("(cid:") > 5:
        return False
    if not re.search(r'[A-Za-zÀ-ÿ]{3,}', texto):
        return False

    return True

def registrar_no_log(diretorio, mensagem):
    if gerar_log:
        log_path = os.path.join(diretorio, "historico_renomeio.txt")
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with open(log_path, "a", encoding="utf-8") as f:
            f.write(f"[{timestamp}] {mensagem}\n") 

def extrair_anos_vigencia(texto):
    texto_limpo = " ".join(texto.split())
    match_trecho = re.search(r"(vigência|data\s*base).{0,250}", texto_limpo, re.IGNORECASE)
    
    if match_trecho:
        trecho = match_trecho.group(0)
        padrao_duplo = re.findall(r"\b(20\d{2})\b\s*(?:a|até|/|-)\s*\b(20\d{2})\b", trecho, re.IGNORECASE)

        if padrao_duplo:
            return padrao_duplo[0][0], padrao_duplo[0][1]
        anos = re.findall(r"\b(20\d{2})\b", trecho)

        if len(anos) >= 2: #len diz quantos elementos tem na lista
            return anos[0], anos[1]
        elif len(anos) == 1: 
            return anos[0], anos[0]
        
    return "N/A", "N/A"

def processar_documentos():
    root = Tk()
    root.withdraw()
    pasta = filedialog.askdirectory(title="Selecione a pasta dos arquivos")
    root.destroy()

    if not pasta: return

    arquivos = [f for f in os.listdir(pasta) if f.lower().endswith(('.pdf', '.docx', '.doc'))]

    padrao_pronto = re.compile(r"^\d+.*___\d{4}.*\.pdf$")
    
    for nome_arquivo in arquivos:
     
     #with open("LOG_LEITURA.txt", "w", encoding="utf-8") as log:

        if nome_arquivo.startswith("~$"):
            print(f"-> Ignorando arquivo temporário do sistema: {nome_arquivo}") 
            continue
        if padrao_pronto.match(nome_arquivo): #ignora se já estiver no padrão
            print(f"-> Pulando (Já está no padrão): {nome_arquivo}")
            continue
        if "Arquivos_Nao_Processados" in nome_arquivo: #ignora a própria pasta de erros se ela estiver listada
            continue
        caminho_original = os.path.join(pasta, nome_arquivo)
        if "Arquivos_Nao_Processados" in caminho_original:
            continue
        if not os.path.isfile(caminho_original): #garante que é arquivo
            continue

        ext = nome_arquivo.lower()
        ano_in, ano_out = "N/A", "N/A"

        print(f"-> Processando: {nome_arquivo}")
        
        caminho_pdf_trabalho = caminho_original
        texto_extraido = ""
        
        nome_base = os.path.splitext(nome_arquivo)[0]
        match_filiais = re.match(r"^([\d]{1,3}(?:[\s,]+[\d]{1,3})*)", nome_base)

        if ext.endswith(('.doc', '.docx')):
            print("-> Convertendo para PDF...")
            caminho_pdf_trabalho = doc_para_pdf(caminho_original)
            if not caminho_pdf_trabalho: continue
   
        with pdfplumber.open(caminho_pdf_trabalho) as pdf:
            for pagina in pdf.pages[:3]:
                t = pagina.extract_text()
                if texto_legivel(t):
                    texto_extraido += t + "\n"
                else: 
                    img = pagina.to_image(resolution=200).original
                    texto_extraido += pytesseract.image_to_string(img, lang="por") + "\n"

        ano_in, ano_out = extrair_anos_vigencia(texto_extraido)
        
        if ano_in == "N/A" or not match_filiais:
            motivo = "Vigência não encontrada" if ano_in == "N/A" else "Filial não encontrada no nome"
            print(f"> [FALHA] {motivo}: {nome_arquivo}. Movendo original para pasta de erro.")

            if ext.endswith(('.doc', '.docx')) and os.path.exists(caminho_pdf_trabalho):
                if caminho_pdf_trabalho != caminho_original: #não renomeia e apaga o pdf temporário
                    os.remove(caminho_pdf_trabalho)

            mover_para_erro(pasta, nome_arquivo)
            continue

        filiais = match_filiais.group(1).strip().rstrip(',') if match_filiais else "Doc"

        if not match_filiais:
            print(f"-> [IGNORADO] Padrão de filiais não encontrado no nome: {nome_arquivo}")
            if ext.endswith(('.doc', '.docx')) and caminho_pdf_trabalho != caminho_original: 
                if os.path.exists(caminho_pdf_trabalho):
                    os.remove(caminho_pdf_trabalho) #limpa o temporário se necessário
            continue

        if ano_in == ano_out:
           vigencia_formatada = ano_in 
        else:
           vigencia_formatada = f"{ano_in}-{ano_out}"

        novo_nome = f"{filiais}___{vigencia_formatada}.pdf"
        novo_caminho = os.path.join(pasta, novo_nome)

        try:
            if ext.endswith('.pdf'):
                os.rename(caminho_original, novo_caminho)
            else:
                if os.path.exists(novo_caminho):
                    os.remove(caminho_pdf_trabalho) 
                    print(f"> [ERRO] PDF {novo_nome} já existe.")
                else:
                    os.rename(caminho_pdf_trabalho, novo_caminho)
                    if os.path.exists(novo_caminho):
                         os.remove(caminho_original) #apaga o arquivo original doc/docx
            
            print(f"-> SUCESSO: Salvo como {novo_nome}")
            registrar_no_log(pasta, f"Convertido/Renomeado: {nome_arquivo} -> {novo_nome}")

            #log.write(f"FILE: {nome_arquivo}\n{texto_extraido}\n{'-'*50}\n")

        except Exception as e:
            print(f"-> Erro ao finalizar arquivo: {e}")

if __name__ == "__main__":
    processar_documentos()