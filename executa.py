import pdfplumber
import re
import os
import shutil
import requests
import json
from datetime import datetime

# ================= CONFIGURAÇÃO =================
URL_WEBAPP = "https://script.google.com/macros/s/AKfycbwYOBCSEwak53AA_eXIzubyi5dHSrKL2wpAeKvogzR3MHaxIsPDOuFuRaxTL3WwLHNX/exec"

PASTA_ENTRADA = './pedidos'
PASTA_LIDOS = './pedidos/lidos'
# =================================================

def limpar_valor_monetario(texto):
    if not texto: return 0.0
    texto = texto.lower().replace('r$', '').strip().replace('.', '').replace(',', '.')
    try: return float(texto)
    except: return 0.0

def identificar_unidade(texto):
    if re.search(r'\bPAR\b', texto, re.IGNORECASE): return "PAR"
    if re.search(r'\bM\b|\bMTS\b|\bMETRO\b', texto, re.IGNORECASE): return "METRO"
    return "UNID"

def extrair_local_entrega(texto):
    texto_upper = texto.upper()
    if "NE-03" in texto_upper or "SEST" in texto_upper: return "Santo Estêvão (NE-03)"
    if "NE-08" in texto_upper or "ITABERABA" in texto_upper: return "Itaberaba (NE-08)"
    if "NE-09" in texto_upper or "VDC" in texto_upper: return "Vitória da Conquista (NE-09)"

    matches = re.findall(r'Cidade:\s*([A-Z\s]+)', texto)
    for c in matches:
        if "CRUZ DAS ALMAS" not in c.upper(): return c.strip().upper()
    return "Local Não Identificado"

def processar_pdf_dass(caminho_arquivo, nome_arquivo):
    try:
        with pdfplumber.open(caminho_arquivo) as pdf:
            texto_completo = ""
            for page in pdf.pages:
                texto_completo += page.extract_text() or ""

            if "DASS" not in texto_completo and "01287588" not in texto_completo:
                return None

            # --- 1. DATA DE RECEBIMENTO (Baseado na Data da Emissão) ---
            # Busca Data da emissão (ex: 16/12/2025)
            match_emissao = re.search(r'Data da emissão:\s*(\d{2}/\d{2}/\d{4})', texto_completo, re.IGNORECASE)
            
            if match_emissao:
                data_recebimento = match_emissao.group(1)
            else:
                # Fallback: Cabeçalho Hora/Data (Data de impressão)
                match_header = re.search(r'Hora.*?Data\s*(\d{2}/\d{2}/\d{4})', texto_completo, re.DOTALL)
                data_recebimento = match_header.group(1) if match_header else datetime.now().strftime("%d/%m/%Y")

            # --- 2. DATA DO PEDIDO (Entrega) ---
            # CORREÇÃO: Só busca a data DEPOIS de encontrar "Prev. Ent." para ignorar CEPs do cabeçalho
            idx_inicio_tabela = texto_completo.find("Prev. Ent.")
            if idx_inicio_tabela == -1:
                idx_inicio_tabela = 0 # Se não achar o cabeçalho, procura no texto todo (fallback)

            texto_para_busca = texto_completo[idx_inicio_tabela:]

            # Procura NCM (8 dígitos) seguido de Data
            match_entrega = re.search(r'\d{8}.*?(\d{2}/\d{2}/\d{4})', texto_para_busca, re.DOTALL)

            if match_entrega:
                data_pedido = match_entrega.group(1)
            else:
                # Se não achar data na tabela, usa a de emissão como fallback
                data_pedido = data_recebimento

            # --- 3. ORDEM DE COMPRA ---
            match_ordem = re.search(r'Ordem de compra\s+(\d+)', texto_completo, re.IGNORECASE)
            ordem_compra = match_ordem.group(1) if match_ordem else "N/D"

            # --- DADOS GERAIS ---
            match_marca = re.search(r'Marca:\s*([^\n]+)', texto_completo)
            marca_geral = match_marca.group(1).strip() if match_marca else "N/D"

            local_geral = extrair_local_entrega(texto_completo)
            unidade_geral = identificar_unidade(texto_completo)

            # --- TOTAIS ---
            valor_total_doc = 0.0
            qtd_total_doc = 0

            match_valor_tot = re.search(r'Total valor:\s*([\d\.,]+)', texto_completo)
            if match_valor_tot: 
                valor_total_doc = limpar_valor_monetario(match_valor_tot.group(1))

            match_qtd_tot = re.search(r'Total peças:\s*([\d\.,]+)', texto_completo)
            if match_qtd_tot: 
                qtd_total_doc = int(limpar_valor_monetario(match_qtd_tot.group(1)))

            valor_formatado = f"R$ {valor_total_doc:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

            # --- CRIAÇÃO DO OBJETO ---
            pedido_unico = {
                "dataPedido": data_pedido,           # Data da Entrega (Prev. Ent.)
                "dataRecebimento": data_recebimento, # Data da Emissão
                "arquivo": nome_arquivo,
                "cliente": "Grupo DASS",
                "marca": marca_geral,
                "local": local_geral,
                "qtd": qtd_total_doc,
                "unidade": unidade_geral,
                "valor": valor_formatado,
                "ordemCompra": ordem_compra          # Ordem de Compra
            }

            return [pedido_unico]

    except Exception as e:
        print(f"Erro ao abrir {nome_arquivo}: {e}")
        return []

def mover_arquivos_processados(lista_arquivos):
    if not os.path.exists(PASTA_LIDOS): os.makedirs(PASTA_LIDOS)
    print(f"\n📦 Movendo arquivos processados para: {PASTA_LIDOS}")
    for arquivo in set(lista_arquivos):
        try:
            caminho_origem = os.path.join(PASTA_ENTRADA, arquivo)
            caminho_destino = os.path.join(PASTA_LIDOS, arquivo)
            if os.path.exists(caminho_destino): os.remove(caminho_destino)
            shutil.move(caminho_origem, caminho_destino)
        except: pass

def main():
    if not os.path.exists(PASTA_ENTRADA):
        os.makedirs(PASTA_ENTRADA)
        print(f"Pasta criada. Coloque PDFs em '{PASTA_ENTRADA}'.")
        return

    todos_pedidos_para_envio = []
    arquivos_para_mover = []

    arquivos = [f for f in os.listdir(PASTA_ENTRADA) if f.lower().endswith('.pdf')]

    print(f"📂 Processando {len(arquivos)} arquivos...")
    print("-" * 95)
    print(f"{'EMISSÃO':<12} | {'ENTREGA':<12} | {'OC':<10} | {'MARCA':<15} | {'VALOR'}")
    print("-" * 95)

    for arq in arquivos:
        lista_pedidos = processar_pdf_dass(os.path.join(PASTA_ENTRADA, arq), arq)

        if lista_pedidos:
            for p in lista_pedidos:
                todos_pedidos_para_envio.append(p)
                print(f"✅ {p['dataRecebimento']:<12} | {p['dataPedido']:<12} | {p['ordemCompra']:<10} | {p['marca']:<15} | {p['valor']}")
            arquivos_para_mover.append(arq)
        else:
            print(f"⚠️ Ignorado: {arq}")

    if todos_pedidos_para_envio:
        print("-" * 75)
        print(f"📤 Enviando {len(todos_pedidos_para_envio)} pedidos para Google Sheets...")
        
        try:
            response = requests.post(URL_WEBAPP, json={"pedidos": todos_pedidos_para_envio}, timeout=30)

            print(f"\n📡 Status: {response.status_code}")
            
            if response.status_code == 200:
                print(f"☁️ SUCESSO! Google recebeu os dados.")
                mover_arquivos_processados(arquivos_para_mover)
            else:
                print(f"❌ Erro HTTP {response.status_code}: {response.text}")

        except Exception as e:
            print(f"\n❌ Erro de conexão: {e}")
    else:
        print("\n⚠️ Nenhum pedido válido encontrado.")

    input("\nPressione ENTER para fechar...")

if __name__ == "__main__":
    main()
