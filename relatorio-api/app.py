import os
import re
import json
import tempfile
import base64
import requests
from flask import Flask, request
from google.cloud import bigquery
import pandas as pd
from datetime import datetime

app = Flask(__name__)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Suporte a credenciais via variável de ambiente (para deploy no Render)
_creds_json = os.environ.get("GOOGLE_CREDENTIALS_JSON")
if _creds_json and not os.environ.get("GOOGLE_APPLICATION_CREDENTIALS"):
    _tmp = tempfile.NamedTemporaryFile(mode="w", suffix=".json", delete=False)
    _tmp.write(_creds_json)
    _tmp.flush()
    os.environ["GOOGLE_APPLICATION_CREDENTIALS"] = _tmp.name

client = bigquery.Client(project="gcp-maas-proj-manutencao")

# =========================================================
# CONFIGURAÇÕES DA EMPRESA E STATUS
# =========================================================
NOME_ARQUIVO_LOGO_FIXA = "logo_empresa_fixa.png" 

MAPA_STATUS = {
    "01": "LOCADO",
    "02": "RESERVA",
    "03": "SERVICOS",
    "04": "DISPONIVEL",
    "05": "NEGOCIADO",
    "06": "VENDA",
    "07": "VENDIDO",
    "08": "EM ADEQUAÇÃO",
    "10": "AGUARDANDO DEMANDA",
    "12": "DISTRATADO"
}

def carregar_logo_fixa_base64():
    caminho_logo = os.path.join(BASE_DIR, NOME_ARQUIVO_LOGO_FIXA)
    if os.path.exists(caminho_logo):
        try:
            with open(caminho_logo, "rb") as image_file:
                encoded_string = base64.b64encode(image_file.read()).decode('utf-8')
                ext = os.path.splitext(NOME_ARQUIVO_LOGO_FIXA)[1].lower().replace('.','')
                if not ext: ext = "png"
                return f"data:image/{ext};base64,{encoded_string}"
        except Exception as e:
            print(f"Erro ao ler logo fixa: {e}")
            return ""
    else:
        print(f"AVISO: Arquivo de logo fixa não encontrado em: {caminho_logo}")
        return ""

def normalizar_id(valor):
    v = str(valor).strip().upper()
    if v.endswith('.0'):
        v = v[:-2]
    if v == 'NAN':
        return ""
    if v.isdigit() and len(v) < 8:
        return v.zfill(8)
    return v

def ajustar_link_imagem(link):
    link = str(link).strip()
    if 'drive.google.com/file/d/' in link:
        try:
            file_id = link.split('/d/')[1].split('/')[0]
            return f"https://drive.google.com/uc?export=view&id={file_id}"
        except:
            return link
    elif 'drive.google.com/open?id=' in link:
        try:
            file_id = link.split('id=')[1].split('&')[0]
            return f"https://drive.google.com/uc?export=view&id={file_id}"
        except:
            return link
    return link

def carregar_logos():
    caminho_arquivo = os.path.join(BASE_DIR, "Clientes.xlsx")
    logos = {}
    mensagem_debug = ""
    amostra_ids = []
    
    try:
        if not os.path.exists(caminho_arquivo):
            return {}, f"ERRO: Arquivo Excel não encontrado em {caminho_arquivo}", []
            
        df = pd.read_excel(caminho_arquivo, sheet_name="Relatorio", dtype=str)
        
        if "ID CLIENTE" not in df.columns:
            colunas_encontradas = ", ".join(df.columns.tolist())
            return {}, f"ERRO: Coluna 'ID CLIENTE' não encontrada! Colunas lidas: {colunas_encontradas}", []
            
        mensagem_debug = f"SUCESSO: Planilha lida perfeitamente ({len(df)} linhas)."
        
        df_sorted = df.sort_values(by="CLIENTE", ascending=True)
        
        for _, row in df_sorted.iterrows():
            id_excel = normalizar_id(row.get("ID CLIENTE", ""))
            
            logo_link = str(row.get("Link", "")).strip()
            if logo_link.lower() == 'nan' or not logo_link:
                logo_link = ""
            else:
                logo_link = ajustar_link_imagem(logo_link)
                
            if id_excel:
                logos[id_excel] = {
                    "nome": str(row.get("CLIENTE", "")).strip(),
                    "logo": logo_link
                }
                if len(amostra_ids) < 5:
                    amostra_ids.append(f"'{id_excel}'")
                    
    except Exception as e:
        mensagem_debug = f"ERRO FATAL ao ler Excel: {str(e)}"
        
    return logos, mensagem_debug, amostra_ids

def limpar_nome_veiculo(nome_bruto, placa):
    nome = str(nome_bruto).strip()
    placa = str(placa).strip()
    if placa and placa in nome:
        nome = nome.replace(placa, "")
    nome = re.sub(r'^[\s\-/]+', '', nome) 
    nome = re.sub(r'[\s\-/]+$', '', nome)
    partes = [p.strip() for p in re.split(r'[-/]', nome) if p.strip()]
    if partes:
        return partes[0]
    return "Modelo Indefinido"
@app.route('/')
@app.route('/relatorio')
def relatorio():
    cliente_id_original = request.args.get('cliente', request.args.get('CLIENTE', '')).strip()
    cliente_id_norm = normalizar_id(cliente_id_original) if cliente_id_original else ""

    logos, msg_planilha, amostra_chaves = carregar_logos()
    logo_fixa_src = carregar_logo_fixa_base64()

    condicao = f"WHERE T9_CLIENTE = '{cliente_id_original}'" if cliente_id_original else ""
    query = f"""
        SELECT T9_CODBEM AS `PREFIXO`, T9_NOME, T9_STATUS, T9_PLACA, T9_CLIENTE, T9_POSCONT, T9_ANOMOD, T9_ANOFAB, T9_CHASSI, T9_XCONTRA, T9_XLOTE
        FROM `gcp-maas-proj-manutencao.silver.MAAS_ST9`
        {condicao}
        ORDER BY T9_XCONTRA, T9_STATUS, T9_NOME
    """
    resultados = list(client.query(query).result())

    info_cliente = logos.get(cliente_id_norm, {
        "nome": "Todos os Clientes",
        "logo": ""
    })

    logo_cliente_src = logo_fixa_src if not cliente_id_original else info_cliente.get('logo', '')
    status_download_cliente = "Não foi tentado (Sem Link)"
    
    if logo_cliente_src and str(logo_cliente_src).startswith('http'):
        status_download_cliente = "Tentando baixar..."
        try:
            headers = {'User-Agent': 'Mozilla/5.0'}
            response = requests.get(logo_cliente_src, headers=headers, timeout=5)
            if response.status_code == 200:
                content_type = response.headers.get('Content-Type', '')
                if 'text/html' not in content_type:
                    b64 = base64.b64encode(response.content).decode('utf-8')
                    logo_cliente_src = f"data:{content_type};base64,{b64}"
                    status_download_cliente = "SUCESSO!"
                else:
                    logo_cliente_src = ""
            else:
                logo_cliente_src = ""
        except Exception:
            logo_cliente_src = ""
            
    info_cliente['logo_render_cliente'] = logo_cliente_src

    veiculos_processados = []
    for row in resultados:
        v = dict(row)
        v['MODELO_LIMPO'] = limpar_nome_veiculo(v.get('T9_NOME', ''), v.get('T9_PLACA', ''))
        veiculos_processados.append(v)

    dados_agrupados = {}
    contagem_status = {}
    contagem_modelo = {}
    
    for v in veiculos_processados:
        contrato = v.get('T9_XCONTRA') or "Sem Contrato"
        status_codigo = str(v.get('T9_STATUS') or "Sem Status").strip()
        status_descricao = MAPA_STATUS.get(status_codigo, status_codigo)
        modelo = v.get('MODELO_LIMPO') or "Sem Modelo"
        
        contagem_status[status_descricao] = contagem_status.get(status_descricao, 0) + 1
        contagem_modelo[modelo] = contagem_modelo.get(modelo, 0) + 1
        
        if contrato not in dados_agrupados: dados_agrupados[contrato] = {}
        if status_descricao not in dados_agrupados[contrato]: dados_agrupados[contrato][status_descricao] = {}
        if modelo not in dados_agrupados[contrato][status_descricao]: dados_agrupados[contrato][status_descricao][modelo] = []
        
        dados_agrupados[contrato][status_descricao][modelo].append(v)

    contratos_lista = sorted([c for c in dados_agrupados.keys() if c != "Sem Contrato"])
    total_contratos = len(contratos_lista) 
    
    if "Sem Contrato" in dados_agrupados:
        contratos_lista.append("Sem Contrato")

    total_veiculos = len(veiculos_processados)
    contagem_status_ordenada = dict(sorted(contagem_status.items(), key=lambda item: item[1], reverse=True))
    contagem_modelo_ordenada = dict(sorted(contagem_modelo.items(), key=lambda item: item[1], reverse=True))

    linhas_tabela_resumo = ""
    for status, qtd in contagem_status_ordenada.items():
        linhas_tabela_resumo += f"<tr><td>{status}</td><td style='text-align: center; font-weight: bold; font-size: 13px;'>{qtd}</td></tr>"

    linhas_tabela_modelo = ""
    for mod, qtd in contagem_modelo_ordenada.items():
        linhas_tabela_modelo += f"<tr><td>{mod}</td><td style='text-align: center; font-weight: bold; font-size: 13px;'>{qtd}</td></tr>"

    data_geracao = datetime.now().strftime("%d/%m/%Y %H:%M")

    opcoes_dropdown = '<option value="">-- Selecione um Cliente --</option>'
    for id_cli, info in logos.items():
        selecionado = 'selected' if id_cli == cliente_id_norm else ''
        opcoes_dropdown += f'<option value="{id_cli}" {selecionado}>{info["nome"]}</option>'

    html = f"""
    <!DOCTYPE html>
    <html lang="pt-BR">
    <head>
        <meta charset="UTF-8">
        <title>Relatório - {info_cliente['nome']}</title>
        <style>
            body {{ font-family: 'Segoe UI', Arial, sans-serif; background-color: #e9ecef; margin: 0; padding: 20px; color: #333; }}
            
            /* --- BARRA DE CONTROLE --- */
            .barra-controle {{ background-color: #002939; padding: 15px 20px; border-radius: 5px; margin: 0 auto 20px auto; max-width: 210mm; display: flex; justify-content: space-between; align-items: center; box-shadow: 0 4px 6px rgba(0,0,0,0.1); color: white; }}
            .barra-controle label {{ font-weight: bold; margin-right: 10px; }}
            .barra-controle select {{ padding: 8px; border-radius: 4px; border: 1px solid #ccc; width: 300px; font-size: 14px; outline: none; }}
            .barra-controle button {{ background-color: #ffffff; color: #002939; border: none; padding: 8px 15px; border-radius: 4px; cursor: pointer; font-weight: bold; transition: 0.2s; }}
            .barra-controle button:hover {{ background-color: #e9ecef; }}
            
            .folha-a4 {{ background-color: #ffffff; max-width: 210mm; margin: 0 auto; padding: 30px 40px; box-shadow: 0 4px 8px rgba(0,0,0,0.1); border-radius: 4px; min-height: 297mm; }}
            
            .cabecalho {{ display: flex; justify-content: space-between; align-items: center; border-bottom: 3px solid #002f47; padding-bottom: 15px; margin-bottom: 20px; }}
            .cabecalho .logo-fixa {{ flex: 1; text-align: left; }}
            .cabecalho .logo-fixa img {{ max-height: 60px; max-width: 120px; object-fit: contain; }}
            .cabecalho .titulo-central {{ flex: 2; text-align: center; }}
            .cabecalho .titulo-central h1 {{ margin: 0; font-size: 20px; color: #002939; text-transform: uppercase; letter-spacing: 1px; }}
            .cabecalho .titulo-central p {{ margin: 5px 0 0 0; font-size: 16px; color: #002939; font-weight: bold; }}
            .cabecalho .logo-cliente {{ flex: 1; text-align: right; }}
            .cabecalho .logo-cliente img {{ max-height: 60px; max-width: 120px; object-fit: contain; }}
            .sem-logo {{ font-size: 10px; color: #002939; font-style: italic; }}

            /* --- ESTILOS DO RESUMO (COM ÍCONES) --- */
            .sessao-resumo {{ margin-top: 10px; margin-bottom: 40px; }}
            h3.titulo-resumo {{ background-color: #002939; color: #ffffff; padding: 10px 15px; font-size: 16px; margin-bottom: 20px; border-radius: 4px; text-align: center; letter-spacing: 1px; }}
            
            .cards-resumo {{ display: flex; gap: 20px; margin-bottom: 30px; }}
            .card {{ flex: 1; background-color: #f8f9fa; border: 1px solid #dee2e6; border-radius: 5px; padding: 20px; box-shadow: 0 2px 4px rgba(0,0,0,0.02); display: flex; align-items: center; justify-content: center; gap: 20px; }}
            .card-icone {{ width: 50px; height: 50px; background-color: #002939; color: white; border-radius: 50%; display: flex; align-items: center; justify-content: center; flex-shrink: 0; }}
            .card-icone svg {{ width: 26px; height: 26px; stroke-width: 2; }}
            .card-info {{ display: flex; flex-direction: column; text-align: left; }}
            .card-titulo {{ display: block; font-size: 13px; color: #555; text-transform: uppercase; font-weight: bold; margin-bottom: 5px; }}
            .card-valor {{ display: block; font-size: 32px; color: #002939; font-weight: bold; line-height: 1; }}
            
            h4.subtitulo-resumo {{ color: #002939; font-size: 15px; margin-bottom: 15px; border-bottom: 2px solid #dee2e6; padding-bottom: 5px; }}
            
            /* --- TABELAS --- */
            table {{ width: 100%; border-collapse: collapse; margin-bottom: 20px; font-size: 11px; margin-left: 10px; width: calc(100% - 10px); }}
            th, td {{ border: 1px solid #dee2e6; padding: 7px; text-align: center; }}
            th {{ background-color: #e9ecef; color: #333; font-weight: bold; text-transform: uppercase; }}
            tr:nth-child(even) {{ background-color: #f8f9fa; }}
            
            .tabela-resumo {{ width: 60%; margin: 0 auto; margin-left: auto; }}
            .tabela-resumo th {{ background-color: #002939; color: white; }}

            .sessao-contrato {{ margin-top: 25px; }}
            h3.titulo-contrato {{ background-color: #002939; color: #ffffff; padding: 10px 15px; font-size: 16px; margin-bottom: 15px; border-radius: 4px; }}
            .sessao-status {{ margin-left: 10px; margin-bottom: 20px; }}
            h4.titulo-status {{ background-color: #f1f3f5; padding: 8px 12px; border-left: 5px solid #002939; font-size: 15px; margin-bottom: 15px; color: #333; }}
            h5.titulo-modelo {{ font-size: 14px; color: #495057; margin-bottom: 8px; padding-left: 10px; border-bottom: 2px solid #dee2e6; padding-bottom: 5px; }}
            
            .rodape {{ text-align: center; margin-top: 40px; padding-top: 10px; border-top: 1px solid #dee2e6; font-size: 10px; color: #888; line-height: 1.5; }}
            .mensagem-vazia {{ text-align: center; color: #666; font-style: italic; margin-top: 50px; font-size: 18px; }}
            
            /* ========================================================= */
            /* --- ESTILOS DE IMPRESSÃO --- */
            /* ========================================================= */
            @page {{ size: A4; margin: 10mm 15mm; }}
            @media print {{
                body {{ background-color: transparent; padding: 0; margin: 0; }}
                .barra-controle {{ display: none !important; }}
                .folha-a4 {{ max-width: 100%; box-shadow: none; border: none; padding: 0; margin: 0; min-height: auto; }}
                * {{ -webkit-print-color-adjust: exact !important; print-color-adjust: exact !important; }}
                .cabecalho {{ margin-bottom: 10px; padding-bottom: 5px; }}
                .quebra-pagina-impressao {{ page-break-after: always; }}
                .sessao-contrato + .sessao-contrato {{ page-break-before: always; break-before: page; }}
                .sessao-contrato {{ margin-top: 10px; }}
                h3.titulo-contrato {{ margin-bottom: 10px; padding: 6px 12px; }}
                h4.titulo-status {{ margin-bottom: 8px; padding: 6px 10px; }}
                table {{ margin-bottom: 10px; page-break-inside: auto; }}
                tr {{ page-break-inside: avoid; page-break-after: auto; }}
                h3, h4, h5 {{ page-break-after: avoid; }}
                .rodape {{ page-break-inside: avoid; margin-top: 20px; }}
            }}
        </style>
    </head>
    <body>
    
        <div class="barra-controle">
            <div>
                <label for="seletor-cliente">🏢 Filtrar Cliente:</label>
                <select id="seletor-cliente" onchange="window.location.href='/relatorio?cliente=' + this.value;">
                    {opcoes_dropdown}
                </select>
            </div>
<button onclick="window.print()">🖨️ Imprimir Relatório</button>
        </div>

        <div class="folha-a4">
            
            <div class="cabecalho">
                <div class="logo-fixa">
                    """
    if logo_fixa_src:
        html += f'<img src="{logo_fixa_src}" alt="Logo Empresa">'
    else:
        html += f'<span class="sem-logo">[Arquivo {NOME_ARQUIVO_LOGO_FIXA} não encontrado]</span>'
        
    html += f"""
                </div>
                
                <div class="titulo-central">
                    <h1>Relatório de Veículos</h1>
                    <p>{info_cliente['nome']}</p>
                </div>
                
                <div class="logo-cliente">
                    """
    if info_cliente['logo_render_cliente']:
        html += f'<img src="{info_cliente['logo_render_cliente']}" alt="Logo Cliente">'
    else:
        html += '<span class="sem-logo">[Sem Logo Cliente]</span>'
        
    html += f"""
                </div>
            </div>
            """

    if not dados_agrupados:
        html += "<div class='mensagem-vazia'>Nenhum veículo encontrado para este cliente.</div>"
    else:
        html += f"""
            <div class="sessao-resumo">
                <h3 class="titulo-resumo">RESUMO GERAL</h3>
                
                <div class="cards-resumo">
                    <div class="card">
                        <div class="card-icone">
                            <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-linecap="round" stroke-linejoin="round">
                                <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path>
                                <polyline points="14 2 14 8 20 8"></polyline>
                                <line x1="16" y1="13" x2="8" y2="13"></line>
                                <line x1="16" y1="17" x2="8" y2="17"></line>
                                <polyline points="10 9 9 9 8 9"></polyline>
                            </svg>
                        </div>
                        <div class="card-info">
                            <span class="card-titulo">Contratos Ativos</span>
                            <span class="card-valor">{total_contratos}</span>
                        </div>
                    </div>
                    <div class="card">
                        <div class="card-icone">
                            <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-linecap="round" stroke-linejoin="round">
                                <rect x="1" y="3" width="15" height="13"></rect>
                                <polygon points="16 8 20 8 23 11 23 16 16 16 16 8"></polygon>
                                <circle cx="5.5" cy="18.5" r="2.5"></circle>
                                <circle cx="18.5" cy="18.5" r="2.5"></circle>
                            </svg>
                        </div>
                        <div class="card-info">
                            <span class="card-titulo">Total de Veículos</span>
                            <span class="card-valor">{total_veiculos}</span>
                        </div>
                    </div>
                </div>
                
                <h4 class="subtitulo-resumo">Detalhamento por Status</h4>
                <table class="tabela-resumo">
                    <thead>
                        <tr>
                            <th>Status Operacional</th>
                            <th style="text-align: center; width: 120px;">Quantidade</th>
                        </tr>
                    </thead>
                    <tbody>
                        {linhas_tabela_resumo}
                    </tbody>
                </table>
                
                <h4 class="subtitulo-resumo" style="margin-top: 30px;">Detalhamento por Modelo</h4>
                <table class="tabela-resumo">
                    <thead>
                        <tr>
                            <th>Modelo do Veículo</th>
                            <th style="text-align: center; width: 120px;">Quantidade</th>
                        </tr>
                    </thead>
                    <tbody>
                        {linhas_tabela_modelo}
                    </tbody>
                </table>
            </div>
            
            <div class="quebra-pagina-impressao"></div>
        """

        for contrato in contratos_lista:
            dict_status = dados_agrupados[contrato]
            
            html += f"<div class='sessao-contrato'>"
            html += f"<h3 class='titulo-contrato'>CONTRATO: {contrato}</h3>"
            for status_desc, dict_modelo in dict_status.items():
                
                total_status = sum(len(veiculos) for veiculos in dict_modelo.values())
                
                html += f"<div class='sessao-status'>"
                html += f"<h4 class='titulo-status'>{status_desc} <span style='font-size: 13px; font-weight: normal; color: #555;'>({total_status} veículos)</span></h4>"
                for modelo, veiculos in dict_modelo.items():
                    html += f"<h5 class='titulo-modelo'>Modelo: {modelo} <span style='font-weight: normal; color: #666;'>({len(veiculos)} veículos)</span></h5>"

                    veiculos = sorted(veiculos, key=lambda v: str(v.get('T9_XLOTE') or '').strip())

                    html += "<table><thead><tr><th>Lote</th><th>Prefixo</th><th>Placa</th><th>Chassi</th><th>Ano Mod.</th><th>Ano Fab.</th><th style='text-align: right;'>Hodômetro</th></tr></thead><tbody>"
                    i = 0
                    while i < len(veiculos):
                        v = veiculos[i]
                        lote_atual = v.get('T9_XLOTE') or '-'
                        rowspan = 1
                        while i + rowspan < len(veiculos) and (veiculos[i + rowspan].get('T9_XLOTE') or '-') == lote_atual:
                            rowspan += 1
                        def fmt_hod(val):
                            try:
                                return f"{int(float(val)):,}".replace(",", ".") if val and str(val).strip() not in ['-', '', 'None', 'nan'] else '-'
                            except (ValueError, TypeError):
                                return val or '-'
                        cod_bem = v.get('PREFIXO') or v.get('T9_CODBEM') or '-'
                        html += f"<tr><td rowspan='{rowspan}' style='vertical-align: top;'>{lote_atual}</td><td>{cod_bem}</td><td>{v.get('T9_PLACA') or '-'}</td><td>{v.get('T9_CHASSI') or '-'}</td><td>{v.get('T9_ANOMOD') or '-'}</td><td>{v.get('T9_ANOFAB') or '-'}</td><td style='text-align: right;'>{fmt_hod(v.get('T9_POSCONT'))}</td></tr>"
                        for j in range(1, rowspan):
                            v2 = veiculos[i + j]
                            cod_bem2 = v2.get('PREFIXO') or v2.get('T9_CODBEM') or '-'
                            html += f"<tr><td>{cod_bem2}</td><td>{v2.get('T9_PLACA') or '-'}</td><td>{v2.get('T9_CHASSI') or '-'}</td><td>{v2.get('T9_ANOMOD') or '-'}</td><td>{v2.get('T9_ANOFAB') or '-'}</td><td style='text-align: right;'>{fmt_hod(v2.get('T9_POSCONT'))}</td></tr>"
                        i += rowspan
                    html += "</tbody></table>"
                html += "</div>"
            html += "</div>"

    html += f"""
            <div class='rodape'>
                Documento gerado automaticamente pelo Sistema MAAS.<br>
                Gerado em: {data_geracao}
            </div>
        </div>
    </body>
    </html>
    """
    return html

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)