import streamlit as st
import pdfplumber
import pandas as pd
import re
import os
import io
import zipfile
from datetime import datetime
from fpdf import FPDF
from num2words import num2words
import unicodedata

# --- 1. DADOS FIXOS (Hardcoded) ---
SECRETARIAS = ["Secretaria de Assistência Social"]
SETORES = ["Fundo Municipal de Assistência Social (FMAS)"]
ASSINANTES = ["Felipe Cunha da Silva - Encarregado Operacional do FMAS"]

# --- 2. CONFIGURAÇÕES DA PÁGINA ---
st.set_page_config(page_title="Gerador de Pagamentos - FMAS", layout="wide")

# --- 3. FUNÇÕES AUXILIARES ---

def sanitize_filename(name):
    """Remove caracteres especiais para nomes de arquivos."""
    nfkd_form = unicodedata.normalize('NFKD', name)
    only_ascii = nfkd_form.encode('ASCII', 'ignore').decode('ASCII')
    return re.sub(r'[\\/*?:"<>|]', "", only_ascii)

def format_currency_ptbr(value):
    """Formata valor para R$ 0.000,00."""
    try:
        return f"R$ {value:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except:
        return "R$ 0,00"

def value_to_extenso(value):
    """Converte valor para extenso em português."""
    try:
        return num2words(value, lang='pt_BR', to='currency')
    except:
        return ""

def extract_data_from_pdf(pdf_file):
    """Extrai dados do PDF usando pdfplumber e Regex."""
    text = ""
    tables = []
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            text += page.extract_text() or ""
            page_tables = page.extract_tables()
            if page_tables:
                tables.extend(page_tables)

    # Identificação AF vs AS
    is_as = bool(re.search(r"3\.3\.90\.39\.\d{2}", text))
    tipo_doc = "Autorização de Serviço (AS)" if is_as else "Autorização de Fornecimento (AF)"
    tipo_sigla = "AS" if is_as else "AF"

    # Regex para extração
    af_as_match = re.search(r"(?:Autorização de (?:Fornecimento|Serviço)|AF|AS)\s*(?:Nº|nº|No)?\s*(\d+/\d{4})", text, re.IGNORECASE)
    af_as_num = af_as_match.group(1) if af_as_match else ""

    fornecedor_match = re.search(r"Fornecedor:\s*(.+)", text, re.IGNORECASE)
    fornecedor = fornecedor_match.group(1).strip() if fornecedor_match else "Não Identificado"

    # Solicitante
    solicitante_match = re.search(r"Solicitante:\s*(.+)", text, re.IGNORECASE)
    solicitante = solicitante_match.group(1).strip() if solicitante_match else ""

    cnpj_match = re.search(r"\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}", text)
    cnpj = cnpj_match.group(0) if cnpj_match else ""

    empenhos = re.findall(r"Empenho\s*(?:Nº|nº|No)?\s*(\d+/\d{4})", text, re.IGNORECASE)
    empenhos = list(dict.fromkeys(empenhos))

    data_match = re.search(r"Data:\s*(\d{2}/\d{2}/\d{4})", text, re.IGNORECASE)
    data_af = data_match.group(1) if data_match else datetime.now().strftime("%d/%m/%Y")

    valores = re.findall(r"R\$\s*([\d\.,]+)", text)
    valor_total = 0.0
    if valores:
        try:
            valor_str = valores[-1].replace(".", "").replace(",", ".")
            valor_total = float(valor_str)
        except:
            pass

    # Extração de Itens e Descrição Resumida
    itens_data = []
    descricao_resumida = ""
    for table in tables:
        for row in table:
            if row and len(row) >= 5:
                if "Item" in str(row[0]) or "Descrição" in str(row):
                    continue
                itens_data.append(row)
                if not descricao_resumida:
                    # Tenta pegar a descrição do primeiro item
                    for cell in row:
                        if cell and len(str(cell)) > 10:
                            descricao_resumida = str(cell)[:100]
                            break

    if not descricao_resumida:
        descricao_resumida = f"Referente a {tipo_sigla} {af_as_num}"

    return {
        "filename": pdf_file.name,
        "tipo_doc": tipo_doc,
        "tipo_sigla": tipo_sigla,
        "af_as_num": af_as_num,
        "fornecedor": fornecedor,
        "solicitante": solicitante,
        "cnpj": cnpj,
        "empenhos": empenhos if empenhos else [""],
        "data_af": data_af,
        "valor_total": valor_total,
        "itens": itens_data,
        "descricao": descricao_resumida
    }

# --- 4. GERAÇÃO DE PDF ---

class PDFGenerator(FPDF):
    def __init__(self, orientation='P', unit='mm', format='A4', bg_image=None):
        super().__init__(orientation, unit, format)
        self.bg_image = bg_image

    def header(self):
        if self.bg_image:
            # Se for paisagem, as dimensões são 297x210
            w = 297 if self.cur_orientation == 'L' else 210
            h = 210 if self.cur_orientation == 'L' else 297
            self.image(self.bg_image, x=0, y=0, w=w, h=h)

def generate_memorando(data, memo_num, assinante_info, secretaria, setor, bg_image=None):
    pdf = PDFGenerator(bg_image=bg_image)
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)
    
    pdf.set_font("Arial", "B", 12)
    setor_abreviado = "".join([w[0] for w in setor.split() if w[0].isupper()])
    pdf.cell(0, 10, f"MEMORANDO Nº. {memo_num} - {setor_abreviado}", ln=True, align='L')
    
    pdf.set_font("Arial", "", 10)
    pdf.ln(5)
    pdf.cell(0, 5, f"Da: {secretaria}", ln=True)
    pdf.cell(0, 5, f"Setor: {setor}", ln=True)
    pdf.cell(0, 5, f"Para: {setor}", ln=True)
    pdf.ln(5)
    
    pdf.set_font("Arial", "B", 10)
    pdf.cell(0, 5, f"Ref: Prestação de contas do Empenho de nº. {data['empenho']} em nome de {data['fornecedor']}.", ln=True)
    pdf.ln(5)
    
    pdf.set_font("Arial", "", 10)
    meses = ["janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"]
    hoje = datetime.now()
    data_extenso = f"Caraguatatuba, {hoje.day} de {meses[hoje.month-1]} de {hoje.year}"
    pdf.cell(0, 10, data_extenso, ln=True, align='L')
    pdf.ln(10)
    
    valor_formatado = format_currency_ptbr(data['valor_total'])
    valor_extenso = value_to_extenso(data['valor_total'])
    
    corpo = (f"Encaminho nota fiscal de nº. {data['nf']}, em nome de {data['fornecedor']} "
             f"no valor de {valor_formatado} ({valor_extenso}) referente à execução do "
             f"Empenho de nº. {data['empenho']} correspondente à {data['tipo_doc']} {data['af_as_num']} "
             f"datada de {data['data_af']}.")
    
    pdf.multi_cell(0, 6, corpo, align='J')
    pdf.ln(20)
    
    pdf.set_font("Arial", "I", 10)
    pdf.cell(0, 5, "Atenciosamente,", ln=False, align='L')
    
    pdf.set_font("Arial", "B", 10)
    nome_assinante = assinante_info.split(" - ")[0]
    cargo_assinante = assinante_info.split(" - ")[1] if " - " in assinante_info else ""
    
    pdf.cell(0, 5, nome_assinante, ln=True, align='R')
    pdf.set_font("Arial", "", 10)
    pdf.cell(0, 5, cargo_assinante, ln=True, align='R')
    
    return pdf.output()

def generate_requisicao(data, req_num, secretaria, bg_image=None):
    pdf = PDFGenerator(bg_image=bg_image)
    pdf.add_page()
    pdf.set_font("Arial", "B", 12)
    pdf.cell(0, 10, f"Requisição: {req_num} - Ano {datetime.now().year}", ln=True, align='C')
    pdf.set_font("Arial", "", 10)
    pdf.cell(0, 5, "Almoxarifado: 01 CENTRAL", ln=True)
    pdf.cell(0, 5, f"Centro de Custo: {secretaria}", ln=True)
    pdf.cell(0, 5, "Veículo: GERAL", ln=True)
    pdf.ln(10)
    
    pdf.set_font("Arial", "B", 8)
    col_widths = [10, 15, 15, 20, 100, 30]
    headers = ["Item", "Qtd", "Unid", "Código", "Descrição", "V. Unit"]
    for i, h in enumerate(headers):
        pdf.cell(col_widths[i], 7, h, border=1, align='C')
    pdf.ln()
    
    pdf.set_font("Arial", "", 7)
    for item in data['itens']:
        row = [str(x)[:40] if x else "" for x in item]
        while len(row) < 6: row.append("")
        for i, val in enumerate(row[:6]):
            pdf.cell(col_widths[i], 6, val, border=1)
        pdf.ln()
    return pdf.output()

def generate_protocolo(rows, destinatario, bg_image=None):
    """Gera Protocolo de Entrega em Paisagem."""
    pdf = PDFGenerator(orientation='L', bg_image=bg_image)
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)
    
    hoje = datetime.now().strftime("%d/%m/%Y")
    pdf.set_font("Arial", "B", 14)
    pdf.cell(0, 10, "PROTOCOLO DE ENTREGA", ln=True, align='C')
    pdf.ln(5)
    
    pdf.set_font("Arial", "", 11)
    texto_topo = f"Relação de Notas Fiscais entregue ao {destinatario} em {hoje}"
    if destinatario == "Contabilidade":
        texto_topo = f"Relação de Notas Fiscais entregue à Contabilidade em {hoje}"
    
    pdf.cell(0, 10, texto_topo, ln=True, align='L')
    pdf.ln(5)
    
    # Tabela
    pdf.set_font("Arial", "B", 9)
    # Largura total Paisagem A4: 297mm - margens (20mm) = ~277mm
    # Colunas: Memo/Req, Empenho, Descrição, Solicitante, NF, Valor
    col_widths = [35, 35, 80, 60, 30, 35]
    col_memo_req = "Nº Memorando" if destinatario == "Contabilidade" else "Nº Requisição"
    headers = [col_memo_req, "Nº Empenho", "Descrição Resumida", "Solicitante", "Nº NF", "Valor NF"]
    
    for i, h in enumerate(headers):
        pdf.cell(col_widths[i], 8, h, border=1, align='C')
    pdf.ln()
    
    pdf.set_font("Arial", "", 8)
    for row in rows:
        # Pega o número sequencial (Memo ou Req)
        num_seq = str(row.get('Memo_Num', '')) if destinatario == "Contabilidade" else str(row.get('Req_Num', ''))
        
        data_row = [
            num_seq,
            str(row['Empenho']),
            str(row['Descrição'])[:50],
            str(row['Solicitante'])[:40],
            str(row['Número NF']),
            format_currency_ptbr(row['Valor'])
        ]
        
        # Calcula altura necessária para a linha (baseado na descrição)
        h_row = 7
        for i, val in enumerate(data_row):
            pdf.cell(col_widths[i], h_row, val, border=1)
        pdf.ln()
        
    pdf.ln(15)
    
    # Rodapé de Assinatura
    pdf.set_font("Arial", "", 10)
    pdf.cell(0, 10, "Recebido por: ____________________________________________________________________", ln=True)
    pdf.cell(0, 10, "Assinatura: ______________________________________________________________________", ln=True)
    pdf.cell(0, 10, "Data de Recebimento: ____/____/______", ln=True)
    
    return pdf.output()

# --- 5. INTERFACE STREAMLIT ---

st.title("📄 Gerador Automático de Pagamentos - FMAS")
st.markdown("Sistema de processamento de AF/AS e geração de documentos em lote.")

# BLOCO 1: Configurações de Emissão
with st.container():
    st.subheader("⚙️ Configurações de Emissão")
    col1, col2, col3 = st.columns(3)
    
    with col1:
        memo_inicial = st.number_input("Memorando Nº Inicial", min_value=1, value=1, step=1)
        req_inicial = st.number_input("Requisição Nº Inicial", min_value=1, value=1, step=1)
        
    with col2:
        secretaria_sel = st.selectbox("Secretaria", SECRETARIAS)
        setor_sel = st.selectbox("Setor", SETORES)
        
    with col3:
        assinante_sel = st.selectbox("Usuário Assinante", ASSINANTES)
        bg_file = st.file_uploader("🖼️ Papel Timbrado / Marca d'água (Opcional)", type=["png", "jpg", "jpeg"])

# BLOCO 2: Upload e Processamento
with st.container():
    st.subheader("📤 Upload e Processamento")
    uploaded_files = st.file_uploader("Selecione os arquivos PDF (AF/AS)", type="pdf", accept_multiple_files=True)
    
    if st.button("🚀 Processar Documentos", use_container_width=True):
        if not uploaded_files:
            st.warning("Por favor, faça o upload de pelo menos um arquivo PDF.")
        else:
            processed_data = []
            for pdf_file in uploaded_files:
                data = extract_data_from_pdf(pdf_file)
                for emp in data['empenhos']:
                    processed_data.append({
                        "Selecionar": True,
                        "Arquivo PDF": data['filename'],
                        "Tipo": data['tipo_sigla'],
                        "Número AF/AS": data['af_as_num'],
                        "Número NF": "",
                        "Empenho": emp,
                        "Fornecedor": data['fornecedor'],
                        "Solicitante": data['solicitante'],
                        "Descrição": data['descricao'],
                        "Valor": data['valor_total'],
                        "Data AF": data['data_af'],
                        "Itens": data['itens'],
                        "Tipo Doc": data['tipo_doc']
                    })
            st.session_state['df_processado'] = pd.DataFrame(processed_data)
            st.success(f"{len(uploaded_files)} arquivos processados com sucesso!")

# BLOCO 3: Dashboard de Validação em Lote
if 'df_processado' in st.session_state:
    with st.container():
        st.subheader("📋 Dashboard de Validação")
        
        def get_status(row):
            if not row['Número NF']: return "🟠 Aguardando NF"
            if not row['Número AF/AS'] or not row['Empenho']: return "🔴 Divergência"
            return "🟢 Pronto para Gerar"

        df = st.session_state['df_processado']
        df['Status'] = df.apply(get_status, axis=1)
        
        edited_df = st.data_editor(
            df,
            column_config={
                "Selecionar": st.column_config.CheckboxColumn(default=True),
                "Número NF": st.column_config.TextColumn("Número NF", help="Formato ####", max_chars=10),
                "Valor": st.column_config.NumberColumn("Valor Total", format="R$ %.2f"),
                "Status": st.column_config.TextColumn("Status", disabled=True),
                "Solicitante": st.column_config.TextColumn("Solicitante"),
                "Descrição": st.column_config.TextColumn("Descrição Resumida"),
                "Itens": None, "Tipo Doc": None, "Data AF": None
            },
            disabled=["Arquivo PDF", "Tipo", "Fornecedor", "Status", "Valor"],
            hide_index=True, use_container_width=True, key="data_editor"
        )
        st.session_state['df_processado'] = edited_df

        # Botões de Exportação
        st.divider()
        selected_rows = edited_df[edited_df['Selecionar'] == True]
        
        if not selected_rows.empty:
            col_exp1, col_exp2 = st.columns(2)
            
            with col_exp1:
                if st.button(f"📦 Gerar Documentos e ZIP ({len(selected_rows)})", use_container_width=True):
                    zip_buffer = io.BytesIO()
                    bg_bytes = bg_file.getvalue() if bg_file else None
                    
                    # Temporário para protocolos
                    rows_for_protocol = []
                    
                    with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                        current_memo = memo_inicial
                        current_req = req_inicial
                        
                        for index, row in selected_rows.iterrows():
                            if not row['Número NF']: continue
                                
                            doc_data = {
                                "nf": row['Número NF'], "fornecedor": row['Fornecedor'],
                                "valor_total": row['Valor'], "empenho": row['Empenho'],
                                "tipo_doc": row['Tipo Doc'], "af_as_num": row['Número AF/AS'],
                                "data_af": row['Data AF'], "itens": row['Itens']
                            }
                            
                            row_info = row.to_dict()
                            row_info['Memo_Num'] = current_memo
                            
                            # Memorando
                            memo_pdf = generate_memorando(doc_data, current_memo, assinante_sel, secretaria_sel, setor_sel, bg_image=bg_file)
                            zip_file.writestr(f"MEMO_{current_memo}_NF_{row['Número NF']}.pdf", memo_pdf)
                            current_memo += 1
                            
                            # Requisição
                            if row['Tipo'] == "AF":
                                row_info['Req_Num'] = current_req
                                req_pdf = generate_requisicao(doc_data, current_req, secretaria_sel, bg_image=bg_file)
                                zip_file.writestr(f"REQ_{current_req}_NF_{row['Número NF']}.pdf", req_pdf)
                                current_req += 1
                            
                            rows_for_protocol.append(row_info)
                    
                    st.session_state['rows_for_protocol'] = rows_for_protocol
                    st.download_button(
                        label="📥 Baixar Arquivo ZIP",
                        data=zip_buffer.getvalue(),
                        file_name=f"Documentos_FMAS_{datetime.now().strftime('%Y%m%d')}.zip",
                        mime="application/zip", use_container_width=True
                    )

            # BLOCO: Protocolos de Entrega
            if 'rows_for_protocol' in st.session_state:
                st.subheader("📑 Protocolos de Entrega")
                col_p1, col_p2 = st.columns(2)
                
                rows_p = st.session_state['rows_for_protocol']
                rows_as = [r for r in rows_p if r['Tipo'] == "AS"]
                rows_af = [r for r in rows_p if r['Tipo'] == "AF"]
                
                with col_p1:
                    if rows_as:
                        proto_as = generate_protocolo(rows_as, "Contabilidade", bg_image=bg_file)
                        st.download_button("🧾 Protocolo Contabilidade (AS)", data=proto_as, 
                                         file_name="Protocolo_Contabilidade.pdf", mime="application/pdf", use_container_width=True)
                    else: st.info("Sem AS selecionadas para protocolo.")
                
                with col_p2:
                    if rows_af:
                        proto_af = generate_protocolo(rows_af, "Almoxarifado", bg_image=bg_file)
                        st.download_button("📦 Protocolo Almoxarifado (AF)", data=proto_af, 
                                         file_name="Protocolo_Almoxarifado.pdf", mime="application/pdf", use_container_width=True)
                    else: st.info("Sem AF selecionadas para protocolo.")
        else:
            st.info("Selecione ao menos uma linha na tabela para gerar os documentos.")
        else:
            st.info("Selecione ao menos uma linha na tabela para gerar os documentos.")

# --- RODAPÉ ---
st.divider()
st.caption("Desenvolvido para Fundo Municipal de Assistência Social (FMAS) - Caraguatatuba/SP")
