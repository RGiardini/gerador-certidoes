import streamlit as st
import os
import hashlib
import zipfile
from io import BytesIO
import datetime
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, Cm
from supabase import create_client, Client

# ==========================================
# 1. CONFIGURAÇÃO DA PÁGINA E BANCO DE DADOS
# ==========================================
st.set_page_config(page_title="Sistema de Certidões", layout="centered")

st.markdown("""
    <style>
    .block-container { padding-top: 1rem; padding-bottom: 1rem; }
    h1 { font-size: 22px; text-align: center; margin-bottom: 0; padding-bottom: 0;}
    .stCheckbox { margin-top: -5px; margin-bottom: -5px; }
    </style>
""", unsafe_allow_html=True)

# Conexão com o Supabase
@st.cache_resource
def iniciar_conexao():
    url = st.secrets["SUPABASE_URL"]
    key = st.secrets["SUPABASE_KEY"]
    return create_client(url, key)

supabase: Client = iniciar_conexao()

def gerar_hash_senha(senha):
    return hashlib.sha256(senha.encode()).hexdigest()

# ==========================================
# 2. CONTROLE DE SESSÃO E LOGIN
# ==========================================
if "usuario_logado" not in st.session_state:
    st.session_state["usuario_logado"] = None

if st.session_state["usuario_logado"] is None:
    st.title("⚖️ Sistema de Certidões - TJMG")
    
    aba_login, aba_cadastro = st.tabs(["Entrar", "Criar Nova Conta"])
    
    with aba_login:
        st.write("Acesse sua conta para gerar certidões.")
        usuario_login = st.text_input("Usuário:", key="log_usr").lower().strip()
        senha_login = st.text_input("Senha:", type="password", key="log_pwd")
        
        if st.button("Entrar", type="primary", use_container_width=True):
            if usuario_login and senha_login:
                resposta = supabase.table("banco_usuarios").select("*").eq("usuario", usuario_login).execute()
                
                if len(resposta.data) > 0:
                    dados_bd = resposta.data[0]
                    senha_criptografada = gerar_hash_senha(senha_login)
                    if dados_bd["senha"] == senha_criptografada:
                        st.session_state["usuario_logado"] = usuario_login
                        st.rerun()
                    else:
                        st.error("Senha incorreta!")
                else:
                    st.error("Usuário não encontrado. Vá na aba 'Criar Nova Conta'.")
            else:
                st.warning("Preencha usuário e senha.")
                
    with aba_cadastro:
        st.write("Primeiro acesso? Crie seu usuário e senha abaixo.")
        novo_usuario = st.text_input("Novo Usuário (sem espaços):", key="cad_usr").lower().strip()
        nova_senha = st.text_input("Crie uma Senha:", type="password", key="cad_pwd")
        
        if st.button("Criar Conta", use_container_width=True):
            if novo_usuario and nova_senha:
                checar = supabase.table("banco_usuarios").select("*").eq("usuario", novo_usuario).execute()
                if len(checar.data) > 0:
                    st.error("⚠️ Este nome de usuário já está em uso. Escolha outro.")
                else:
                    supabase.table("banco_usuarios").insert({
                        "usuario": novo_usuario,
                        "senha": gerar_hash_senha(nova_senha),
                        "nome": "",
                        "cargo": "Oficial de Justiça Avaliador",
                        "matricula": ""
                    }).execute()
                    st.success("✅ Conta criada com sucesso! Vá na aba 'Entrar' para acessar o sistema.")
            else:
                st.error("Preencha o usuário e a senha para criar a conta.")
                
    st.stop()

# ==========================================
# 3. DADOS DO USUÁRIO E MENU LATERAL
# ==========================================
usuario_atual = st.session_state["usuario_logado"]
resposta_usuario = supabase.table("banco_usuarios").select("*").eq("usuario", usuario_atual).execute()
dados_usuario = resposta_usuario.data[0]

with st.sidebar:
    st.write(f"👤 Olá, **{usuario_atual.title()}**!")
    st.divider()
    menu = st.radio("Navegação:", ["📝 Gerar Certidão", "📂 Minhas Certidões", "⚙️ Meu Perfil"])
    st.divider()
    if st.button("Sair (Logout)"):
        st.session_state["usuario_logado"] = None
        st.rerun()

# ==========================================
# 4. TELA: MEU PERFIL (NUVEM)
# ==========================================
if menu == "⚙️ Meu Perfil":
    st.title("⚙️ Configurar Meu Perfil")
    st.write("Estes dados serão inseridos no final das suas certidões (Fonte tamanho 8).")
    
    novo_nome = st.text_input("Nome Completo:", value=dados_usuario.get("nome", ""))
    novo_cargo = st.text_input("Cargo:", value=dados_usuario.get("cargo", ""))
    nova_matricula = st.text_input("Matrícula (ex: PJPI: 12345):", value=dados_usuario.get("matricula", ""))
    
    st.write("**Sua Assinatura (Fundo branco ou transparente):**")
    arquivo_assinatura = st.file_uploader("Envie a foto da sua assinatura", type=["png", "jpg", "jpeg"])
    
    if st.button("💾 Salvar Perfil", type="primary"):
        supabase.table("banco_usuarios").update({
            "nome": novo_nome,
            "cargo": novo_cargo,
            "matricula": nova_matricula
        }).eq("usuario", usuario_atual).execute()
        
        if arquivo_assinatura is not None:
            try:
                supabase.storage.from_("assinaturas_usuarios").remove([f"{usuario_atual}.png"])
            except:
                pass
            supabase.storage.from_("assinaturas_usuarios").upload(
                file=arquivo_assinatura.getvalue(),
                path=f"{usuario_atual}.png",
                file_options={"content-type": arquivo_assinatura.type}
            )
                
        st.success("Perfil atualizado e salvo na nuvem com sucesso!")

# ==========================================
# 5. TELA: MINHAS CERTIDÕES
# ==========================================
elif menu == "📂 Minhas Certidões":
    st.title("📂 Minhas Certidões Salvas")
    st.write("Baixe ou exclua seus arquivos salvos na nuvem.")
    
    try:
        arquivos_nuvem = supabase.storage.from_("certidoes_usuarios").list(usuario_atual)
    except:
        arquivos_nuvem = []
    
    arquivos = [arq for arq in arquivos_nuvem if arq["name"] != ".emptyFolder" and arq["name"] != ""]
    
    if not arquivos:
        st.info("Nenhuma certidão salva ainda.")
    else:
        arquivos.sort(key=lambda x: x["created_at"], reverse=True)
        
        c_sel, c_nome, c_data = st.columns([1, 4, 3])
        c_sel.write("**Selecionar**")
        c_nome.write("**Nome do Arquivo**")
        c_data.write("**Data de Criação**")
        st.divider()
        
        arquivos_selecionados = []
        
        for item in arquivos:
            c1, c2, c3 = st.columns([1, 4, 3])
            try:
                data_obj = datetime.datetime.fromisoformat(item["created_at"].replace("Z", "+00:00"))
                data_br = data_obj.strftime("%d/%m/%Y às %H:%M")
            except:
                data_br = "Data desconhecida"

            with c1:
                if st.checkbox("", key=f"chk_{item['name']}"):
                    arquivos_selecionados.append(item['name'])
            with c2:
                st.write(item['name'])
            with c3:
                st.write(data_br)
                
        st.divider()
        
        if arquivos_selecionados:
            st.write(f"**{len(arquivos_selecionados)} arquivo(s) selecionado(s)**")
            c_btn1, c_btn2 = st.columns(2)
            
            with c_btn1:
                if st.button("📥 Preparar Download (ZIP)", type="primary", use_container_width=True):
                    with st.spinner("Baixando da nuvem..."):
                        zip_buffer = BytesIO()
                        with zipfile.ZipFile(zip_buffer, "w") as zip_file:
                            for arq in arquivos_selecionados:
                                arquivo_bytes = supabase.storage.from_("certidoes_usuarios").download(f"{usuario_atual}/{arq}")
                                zip_file.writestr(arq, arquivo_bytes)
                                
                        st.download_button(
                            label="✔️ Clique aqui para baixar o ZIP",
                            data=zip_buffer.getvalue(),
                            file_name=f"certidoes_{usuario_atual}.zip",
                            mime="application/zip",
                            use_container_width=True
                        )
            
            with c_btn2:
                if st.button("🗑️ Excluir Selecionadas", use_container_width=True):
                    caminhos_para_excluir = [f"{usuario_atual}/{arq}" for arq in arquivos_selecionados]
                    supabase.storage.from_("certidoes_usuarios").remove(caminhos_para_excluir)
                    st.success("Arquivos excluídos da nuvem com sucesso!")
                    st.rerun()

# ==========================================
# 6. TELA: GERADOR DE CERTIDÃO (OTIMIZADO PARA MOBILE)
# ==========================================
elif menu == "📝 Gerar Certidão":
    st.title("Certidão Negativa (Detalhada)")
    
    if not dados_usuario.get("nome"):
        st.warning("⚠️ Você ainda não configurou seu perfil! Vá em 'Meu Perfil' no menu lateral e preencha seus dados antes de gerar certidões.")
        st.stop()

    # --- CAMPOS PRINCIPAIS: Lado a lado (Mandado + Processo + Ano) ---
    c_mandado, c_proc, c_ano = st.columns([1, 2.5, 1])
    with c_mandado:
        mandado = st.text_input("Mandado nº:", placeholder="Ex: 01")
    with c_proc:
        processo = st.text_input("Informe o Processo:", placeholder="Ex: 4400281-16")
    with c_ano:
        ano = st.text_input("Ano:", placeholder="Ex: 2026")

    c_end, c_pes = st.columns(2)
    with c_end:
        endereco = st.text_input("Endereço (opcional):", placeholder="Se vazio: 'informado no mesmo'")
    with c_pes:
        pessoa = st.text_input("Pessoa procurada:", placeholder="Deixe vazio para termo genérico")

    st.markdown("---")
    st.write("**Dias e Horários das Diligências:**")
    
    # --- DIAS E HORAS: Emparelhados perfeitamente lado a lado por dia ---
    c_d1, c_h1 = st.columns(2)
    with c_d1:
        d1 = st.text_input("Dia 1", placeholder="Ex: 08/08")
    with c_h1:
        h1 = st.text_input("Hora 1", placeholder="Ex: 14:55hs")
        
    c_d2, c_h2 = st.columns(2)
    with c_d2:
        d2 = st.text_input("Dia 2", placeholder="Ex: 11/08")
    with c_h2:
        h2 = st.text_input("Hora 2", placeholder="Ex: 16:58hs")
        
    c_d3, c_h3 = st.columns(2)
    with c_d3:
        d3 = st.text_input("Dia 3", placeholder="Ex: 12/08")
    with c_h3:
        h3 = st.text_input("Hora 3", placeholder="Ex: 11:15hs")

    st.divider()
    st.write("**Deixei de cumprir o ato uma vez que:**")
    sit_c1, sit_c2 = st.columns(2)
    with sit_c1:
        nao_loc_dest = st.checkbox("O destinatário do mandado não foi localizado")
    with sit_c2:
        nao_loc_bens = st.checkbox("O(s) bem(ns) indicados não foi(ram) localizado(s)")

    # --- BLOCOS RETRÁTEIS (Expansores) ---
    
    motivos_selecionados = []
    with st.expander("📌 Clique aqui para selecionar os Motivos da Negativa (Opcional)", expanded=False):
        motivos_list = [
            "mudou-se", "não reside", "é desconhecido", "dificilmente fica ali", "trabalha em tempo integral",
            "não trabalha no local", "está viajando", "local inabitado", "antigo(a) inquilino(a)", 
            "antigo(a) morador(a)", "antigo(a) proprietário(a)", "rotatividade de inquilinos",
            "foi repassado para terceiros", "encontra-se internado", "foi transferido", "encontra-se preso",
            "faleceu", "faliu", "não exerce(em) atividades no local", "o local estava fechado", 
            "o número não foi localizado", "a rua/av não foi localizada", "o ap/bloco não foi localizado", 
            "aparece por lá esporadicamente", "utiliza o endereço para fins de recebimento de correspondências",
            "\"salvo melhor juízo\" não tem condições psíquicas de entender o conteúdo do presente mandado",
            "encontrei no endereço, apenas bens que, \"salvo melhor juízo\", guarnecem a residência amparados pela Lei 8.009/90",
            "\"salvo melhor juízo\" são insuficientes para saldar o débito e/ou acréscimos legais"
        ]
        cols_mot = st.columns(2)
        for idx, m in enumerate(motivos_list):
            with cols_mot[idx % 2]:
                if st.checkbox(m, key=f"mot_{idx}"):
                    motivos_selecionados.append(m)

    relacoes_selecionadas = []
    nao_sabe_selecionados = []
    sabe_tel = ""
    sabe_end = ""
    
    with st.expander("👤 Informações sobre o Informante (Se houver)", expanded=False):
        nome_inf = st.text_input("Nome do Sr(a):", placeholder="Deixe em branco se não houver informante")

        st.caption("Relação / Qualidade:")
        relacoes_list = [
            "morador(a)", "proprietário(a)", "inquilino(a)", "funcionário(a)", "vizinho(a)", "pai", "mãe",
            "padrasto", "madrasta", "filho(a)", "irmão(a)", "tio(a)", "avô(ó)", "neto(a)", "sobrinho(a)",
            "primo(a)", "transeunte", "viúvo(a)", "ex", "esposo(a)", "companheiro(a)", "sogro(a)", "enteado(a)",
            "genro", "nora", "cunhado(a)", "concunhado(a)", "amigo(a)"
        ]
        cols_rel = st.columns(3)
        for idx, r in enumerate(relacoes_list):
            with cols_rel[idx % 3]:
                if st.checkbox(r, key=f"rel_{idx}"):
                    relacoes_selecionadas.append(r)

        st.markdown("---")
        st.write("**Não sabendo o informante indicar o(a):**")
        nao_sabe_list = [
            "endereço completo", "paradeiro", "o dia e nem o horário exato de localizá-lo(a)", 
            "telefone de contato", "dia e nem o horário exato de retorno", "o presídio", 
            "os dados da certidão de óbito", "previsão de alta"
        ]
        cols_ns = st.columns(2)
        for idx, ns in enumerate(nao_sabe_list):
            with cols_ns[idx % 2]:
                if st.checkbox(ns, key=f"ns_{idx}"):
                    nao_sabe_selecionados.append(ns)

        st.markdown("---")
        st.write("**Sabendo o informante indicar o:**")
        c_sab1, c_sab2 = st.columns(2)
        with c_sab1:
            sabe_tel = st.text_input("Telefone indicado:")
        with c_sab2:
            sabe_end = st.text_input("Endereço correto indicado:")
    
    if 'nome_inf' not in locals():
        nome_inf = ""

    with st.expander("📝 Certificações Adicionais e Observações", expanded=False):
        cert_extras = []
        if st.checkbox("Procurei obter informações junto aos moradores/vizinhos locais e não obtive êxito."):
            cert_extras.append("procurei obter informações junto aos moradores/vizinhos locais e não obtive êxito.")
        if st.checkbox("Devido à importância do mandado, deixei a cópia para ciência do prazo/data."):
            cert_extras.append("devido à importância do mandado e da dificuldade de encontrar a pessoa procurada, deixei a cópia do mandado com o(a) senhor(a) acima mencionado(a) para que a parte/testemunha tome ciência do prazo/data que deverá comparecer em juízo.")
        if st.checkbox("O imóvel é residencial e contém apenas móveis e utensílios domésticos comuns."):
            cert_extras.append("o imóvel é residencial e contém apenas móveis e utensílios domésticos que guarnecem a residência do réu.")

        observacoes = st.text_area("Observações Livres:")

    st.divider()

    if st.button("Salvar na Nuvem / Gerar DOCX", type="primary", use_container_width=True):
        with st.spinner("Construindo certidão e salvando na nuvem..."):
            
            dias_validos = [d for d in [d1, d2, d3] if d]
            horas_validas = [h for h in [h1, h2, h3] if h]
            
            texto_data_hora = ""
            if len(dias_validos) == 1:
                texto_data_hora = f", por volta das {horas_validas[0]}, do dia {dias_validos[0]},"
            elif len(dias_validos) > 1:
                str_horas = ", ".join(horas_validas[:-1]) + f" e {horas_validas[-1]}"
                str_dias = ", ".join(dias_validos[:-1]) + f" e {dias_validos[-1]}"
                texto_data_hora = f", por volta das {str_horas}, dos dias {str_dias}, respectivamente,"

            txt_endereco = f"à {endereco}" if endereco else "ao endereço/local/região/bairro indicado(a)"
            txt_pessoa = f" de {pessoa}" if pessoa else ""

            paragrafo = f"Certifico e dou fé que, em cumprimento ao mandado anexo, desloquei-me {txt_endereco}{texto_data_hora} onde deixei de cumprir o ato emanado no mandado{txt_pessoa}, uma vez que "
            
            sits = []
            if nao_loc_dest: sits.append("o destinatário do mandado não foi localizado")
            if nao_loc_bens: sits.append("o(s) bem(ns) indicados não foi(ram) localizado(s)")
            
            if sits:
                paragrafo += " e ".join(sits) + ". "
            else:
                paragrafo += "não foi possível a sua realização. "

            if motivos_selecionados:
                paragrafo += f"Constatou-se no local que o(a) mesmo(a) {', '.join(motivos_selecionados)}. "

            if nome_inf or relacoes_selecionadas:
                nome_str = nome_inf if nome_inf else "pessoa não identificada"
                rel_str = f", na qualidade de {', '.join(relacoes_selecionadas)}," if relacoes_selecionadas else ""
                
                paragrafo += f"Conforme informações prestadas no local pelo(a) Sr(a). {nome_str}{rel_str} "
                
                if nao_sabe_selecionados:
                    paragrafo += f"este(a) declarou não saber indicar: {', '.join(nao_sabe_selecionados)}. "
                else:
                    paragrafo += "este(a) prestou as devidas informações no local. "
                    
                if sabe_tel or sabe_end:
                    sabes_list = []
                    if sabe_tel: sabes_list.append(f"o telefone de contato {sabe_tel}")
                    if sabe_end: sabes_list.append(f"o endereço atual/correto sendo {sabe_end}")
                    paragrafo += f"Por outro lado, o informante soube indicar {' e '.join(sabes_list)}. "

            if cert_extras:
                paragrafo += f"Certifico também que {'; '.join(cert_extras)}. "
                
            if observacoes:
                paragrafo += f"{observacoes.strip()} "

            doc = Document()
            style = doc.styles['Normal']
            font = style.font
            font.name = 'Times New Roman'
            font.size = Pt(12)

            if os.path.exists("cabecalho.png"):
                p_img_cabecalho = doc.add_paragraph()
                p_img_cabecalho.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run_img_cab = p_img_cabecalho.add_run()
                run_img_cab.add_picture("cabecalho.png", width=Cm(16))
            elif os.path.exists("cabecalho.jpg"):
                p_img_cabecalho = doc.add_paragraph()
                p_img_cabecalho.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run_img_cab = p_img_cabecalho.add_run()
                run_img_cab.add_picture("cabecalho.jpg", width=Cm(16))

            if processo:
                texto_processo = f"Processo: {processo}"
                if ano:
                    texto_processo += f".{ano}.8.13.0245"
                doc.add_paragraph(texto_processo)
                
            if mandado:
                doc.add_paragraph(f"Mandado nº: {mandado}")
                
            doc.add_paragraph("")

            p_titulo = doc.add_paragraph()
            run_titulo = p_titulo.add_run("CERTIDÃO NEGATIVA")
            run_titulo.bold = True
            p_titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            doc.add_paragraph("")

            p_corpo = doc.add_paragraph(paragrafo.strip())
            p_corpo.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p_corpo.paragraph_format.first_line_indent = Pt(35.4)
            p_corpo.paragraph_format.line_spacing = 1.5 
            
            doc.add_paragraph("")

            p_fechamento = doc.add_paragraph("Devolvo o mandado para os devidos fins. O referido é verdade. Dou fé.")
            p_fechamento.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            hoje = datetime.datetime.now()
            meses = ["janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"]
            data_extenso = f"Santa Luzia, {hoje.day} de {meses[hoje.month - 1]} de {hoje.year}."
            
            p_data = doc.add_paragraph(data_extenso)
            p_data.alignment = WD_ALIGN_PARAGRAPH.CENTER

            doc.add_paragraph("")
            
            try:
                assinatura_bytes = supabase.storage.from_("assinaturas_usuarios").download(f"{usuario_atual}.png")
                assinatura_stream = BytesIO(assinatura_bytes)
                
                p_img_assinatura = doc.add_paragraph()
                p_img_assinatura.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run_img_ass = p_img_assinatura.add_run()
                run_img_ass.add_picture(assinatura_stream, width=Cm(6))
            except:
                pass 
            
            p_assinatura = doc.add_paragraph()
            p_assinatura.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            run_nome = p_assinatura.add_run(f"{dados_usuario['nome']}\n")
            run_nome.bold = True
            run_nome.font.size = Pt(8)
            
            run_cargo = p_assinatura.add_run(f"{dados_usuario['cargo']}\n")
            run_cargo.font.size = Pt(8)
            
            run_matricula = p_assinatura.add_run(f"{dados_usuario['matricula']}")
            run_matricula.font.size = Pt(8)

            buffer = BytesIO()
            doc.save(buffer)
            buffer.seek(0)

            data_arquivo = hoje.strftime("%d-%m-%Y_%Hh%M")
            nome_arquivo = f"Certidao_Negativa_{processo}_{data_arquivo}.docx" if processo else f"Certidao_Negativa_{data_arquivo}.docx"
            
            caminho_salvamento = f"{usuario_atual}/{nome_arquivo}"
            
            supabase.storage.from_("certidoes_usuarios").upload(
                file=buffer.getvalue(),
                path=caminho_salvamento,
                file_options={"content-type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document"}
            )

        st.success(f"✅ Certidão salva na sua conta na Nuvem!")
        st.download_button(
            label="📥 Baixar Documento Word Agora",
            data=buffer,
            file_name=nome_arquivo,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary"
        )
