
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
# 🚀 MELHORIA: Layout "wide" ativado e o código abaixo irá reorganizar o layout de fato
st.set_page_config(page_title="Sistema de Certidões", layout="wide")

# --- 🚀 CSS SUPER COMPACTO E ADAPTADO PARA LAYOUT LARGO (WIDE) E REORGANIZADO ---
st.markdown("""
    <style>
    /* 1. Global Page Compaction (WIDE Layout) */
    [data-testid="stAppViewContainer"] .main .block-container { padding: 0.2rem 1rem !important; }
    div[data-testid="stVerticalBlock"] { gap: 0.05rem !important; }

    /* 2. Compact Sidebar (Menu and Login) */
    section[data-testid="stSidebar"] { width: 14rem !important; }
    section[data-testid="stSidebar"] div[data-testid="stVerticalBlock"] { gap: 0.2rem !important; }
    section[data-testid="stSidebar"] h1 { font-size: 1.2rem !important; margin-top: -1rem !important;}
    section[data-testid="stSidebar"] label { font-size: 0.8rem !important; margin-bottom: 0.05rem !important; }
    section[data-testid="stSidebar"] .stButton > button { padding: 0.2rem 0.4rem; font-size: 0.8rem; }
    section[data-testid="stSidebar"] .stRadio label[data-testid="stMarkdownContainer"] { display: none; }

    /* 3. Compact Main Content (Form) */
    .main h1 { font-size: 1.6rem !important; margin-top: 0.05rem !important; margin-bottom: 0.2rem !important; text-align: center; }
    .main h3 { font-size: 1.0rem !important; margin-top: 0.1rem !important; margin-bottom: 0.05rem !important; }
    .main label { font-size: 0.75rem !important; margin-bottom: 0.01rem !important; margin-top: -0.05rem !important; }
    .main .stTextInput input, .main .stSelectbox > div > div > div { font-size: 0.85rem !important; padding: 0.1rem 0.3rem; }

    /* 🚀 NOVO CSS PARA ADAPTAR LARGURAS NO LAYOUT LARGO (Aproveitar mais horizontal) */
    /* Faz com que labels em colunas lado a lado tenham espaçamentos melhores */
    [data-testid="stColumn"] { padding-left: 0.2rem !important; padding-right: 0.2rem !important; }

    /* 4. Aggressively Compact and Increase Radio Button Usability (Touch and Wide) */
    div[role="radiogroup"] { gap: 0.15rem !important; }
    div[role="radiogroup"] div[class^="st-"] { margin-bottom: 0.05rem !important; }
    
    div[role="radiogroup"] div[class^="st-"] > label > div[class^="st-"] > input[type="radio"] { 
        transform: scale(1.6); 
        margin-top: 0.05rem;
        cursor: pointer;
    }
    div[role="radiogroup"] div[class^="st-"] > label > div[data-testid="stMarkdownContainer"] {
        font-size: 0.95rem !important; 
        margin-left: 0.5rem; 
        line-height: 1.2 !important;
    }

    /* 5. Hide Streamlit Branding (cleaner) */
    #MainMenu { visibility: hidden; }
    footer { visibility: hidden; }
    </style>
""", unsafe_allow_html=True)

# Conexão com o Supabase (usa st.cache_resource para manter a conexão ativa)
@st.cache_resource
def iniciar_conexao():
    # Certifique-se de que estas chaves estão configuradas no Render (Environment)
    url = st.secrets["SUPABASE_URL"]
    key = st.secrets["SUPABASE_KEY"]
    return create_client(url, key)

supabase: Client = iniciar_conexao()

# Função para criptografar senhas
def gerar_hash_senha(senha):
    return hashlib.sha256(senha.encode()).hexdigest()

# ==========================================
# 2. CONTROLE DE SESSÃO E LOGIN
# ==========================================
# Inicializa o estado de login se não existir
if "usuario_logado" not in st.session_state:
    st.session_state["usuario_logado"] = None

# Se o usuário não estiver logado, exibe a tela de Login/Cadastro
if st.session_state["usuario_logado"] is None:
    st.title("⚖️ Sistema de Certidões - TJMG")
    
    aba_login, aba_cadastro = st.tabs(["Entrar", "Criar Nova Conta"])
    
    with aba_login:
        st.write("Acesse sua conta para gerar certidões.")
        # Chaves únicas para os inputs da aba de login
        usuario_login = st.text_input("Usuário:", key="log_usr_input").lower().strip()
        senha_login = st.text_input("Senha:", type="password", key="log_pwd_input")
        
        if st.button("Entrar", type="primary", use_container_width=True, key="btn_entrar"):
            if usuario_login and senha_login:
                # Busca usuário no banco
                resposta = supabase.table("banco_usuarios").select("*").eq("usuario", usuario_login).execute()
                
                if len(resposta.data) > 0:
                    dados_bd = resposta.data[0]
                    # Criptografa a senha digitada e compara com a do banco
                    senha_criptografada = gerar_hash_senha(senha_login)
                    if dados_bd["senha"] == senha_criptografada:
                        # Define o usuário logado no session_state e recarrega a página
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
        # Chaves únicas para os inputs da aba de cadastro
        novo_usuario = st.text_input("Novo Usuário (sem espaços):", key="cad_usr_input").lower().strip()
        nova_senha = st.text_input("Crie uma Senha:", type="password", key="cad_pwd_input")
        
        if st.button("Criar Conta", use_container_width=True, key="btn_criar_conta"):
            if novo_usuario and nova_senha:
                # Checa se o usuário já existe
                checar = supabase.table("banco_usuarios").select("*").eq("usuario", novo_usuario).execute()
                if len(checar.data) > 0:
                    st.error("⚠️ Este nome de usuário já está em uso. Escolha outro.")
                else:
                    # Insere o novo usuário com cargo padrão
                    supabase.table("banco_usuarios").insert({
                        "usuario": novo_usuario,
                        "senha": gerar_hash_senha(nova_senha),
                        "nome": "", # Dados do perfil ficam vazios inicialmente
                        "cargo": "Oficial de Justiça Avaliador",
                        "matricula": ""
                    }).execute()
                    st.success("✅ Conta criada com sucesso! Vá na aba 'Entrar' para acessar o sistema.")
            else:
                st.error("Preencha o usuário e a senha para criar a conta.")
                
    # Interrompe a execução aqui se não estiver logado
    st.stop()

# ==========================================
# 3. DADOS DO USUÁRIO E MENU LATERAL
# ==========================================
# Se chegou aqui, o usuário está logado. Busca os dados atualizados dele.
usuario_atual = st.session_state["usuario_logado"]
resposta_usuario = supabase.table("banco_usuarios").select("*").eq("usuario", usuario_atual).execute()
dados_usuario = resposta_usuario.data[0]

with st.sidebar:
    st.write(f"👤 Olá, **{usuario_atual.title()}**!")
    st.divider()
    
    # Opções padrão do menu
    opcoes_menu = ["📝 Gerar Certidão", "📂 Minhas Certidões", "⚙️ Meu Perfil"]
    
    # Adiciona o menu de administrador se o usuário for o específico (Rafael)
    if usuario_atual == "10228429":
        opcoes_menu.append("🛡️ Painel do Administrador")
        
    # MENU LATERAL COMPACTADO PELO CSS ACIMA
    menu = st.radio("Navegação:", opcoes_menu)
    st.divider()
    
    # Botão de Logout
    if st.button("Sair (Logout)", key="btn_logout"):
        st.session_state["usuario_logado"] = None
        st.rerun()

# ==========================================
# 4. TELA: MEU PERFIL
# ==========================================
if menu == "⚙️ Meu Perfil":
    st.title("⚙️ Configurar Meu Perfil")
    st.write("Estes dados serão inseridos no final das suas certidões (Fonte tamanho 8).")
    
    # Inputs pré-preenchidos com os dados atuais do banco
    novo_nome = st.text_input("Nome Completo:", value=dados_usuario.get("nome", ""), key="input_perfil_nome")
    novo_cargo = st.text_input("Cargo:", value=dados_usuario.get("cargo", ""), key="input_perfil_cargo")
    nova_matricula = st.text_input("Matrícula (ex: PJPI: 12345):", value=dados_usuario.get("matricula", ""), key="input_perfil_matricula")
    
    st.write("**Sua Assinatura (Fundo branco ou transparente):**")
    arquivo_assinatura = st.file_uploader("Envie a foto da sua assinatura", type=["png", "jpg", "jpeg"], key="uploader_perfil")
    
    if st.button("💾 Salvar Perfil", type="primary", use_container_width=True, key="btn_salvar_perfil"):
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
                
        st.success("✅ Perfil atualizado e salvo na nuvem com sucesso!")
        st.rerun()

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
                # Pega a data do banco e ajusta o fuso (-3h)
                data_str = item["created_at"].replace("Z", "+00:00")
                data_obj = datetime.datetime.fromisoformat(data_str)
                data_br_obj = data_obj.replace(tzinfo=None) - datetime.timedelta(hours=3)
                data_br = data_br_obj.strftime("%d/%m/%Y às %H:%M")
            except:
                data_br = "Data desconhecida"

            with c1:
                if st.checkbox("", key=f"chk_file_{item['name']}"):
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
                if st.button("📥 Preparar Download (ZIP)", type="primary", use_container_width=True, key="btn_zip_download"):
                    with st.spinner("Baixando e compactando arquivos da nuvem..."):
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
                            use_container_width=True,
                            key="download_zip_real"
                        )
            
            with c_btn2:
                if st.button("🗑️ Excluir Selecionadas", use_container_width=True, key="btn_excluir_certidoes"):
                    caminhos_para_excluir = [f"{usuario_atual}/{arq}" for arq in arquivos_selecionados]
                    supabase.storage.from_("certidoes_usuarios").remove(caminhos_para_excluir)
                    st.success("✅ Arquivos excluídos da nuvem com sucesso!")
                    st.rerun()

# ==========================================
# 6. TELA: PAINEL DO ADMINISTRADOR
# ==========================================
elif menu == "🛡️ Painel do Administrador":
    if usuario_atual != "10228429":
        st.error("Acesso restrito apenas ao Administrador.")
        st.stop()
        
    st.title("🛡️ Painel de Administração")
    
    aba_adm1, aba_adm2 = st.tabs(["👥 Gerenciar Usuários", "📊 Auditoria de Certidões Gerais"])
    
    with aba_adm1:
        st.subheader("Oficiais Cadastrados")
        res_todos = supabase.table("banco_usuarios").select("usuario, nome, cargo, matricula").execute()
        usuarios_cadastrados = res_todos.data
        
        if usuarios_cadastrados:
            for u in usuarios_cadastrados:
                with st.expander(f"👤 Usuário: {u['usuario']} — Nome: {u.get('nome') or 'Não preenchido'}"):
                    st.write(f"**Cargo:** {u.get('cargo')}")
                    st.write(f"**Matrícula:** {u.get('matricula')}")
                    
                    if u['usuario'] != usuario_atual:
                        if st.button(f"🗑️ Excluir usuário {u['usuario']}", key=f"del_adm_usr_{u['usuario']}", use_container_width=True):
                            supabase.table("banco_usuarios").delete().eq("usuario", u['usuario']).execute()
                            st.success(f"Usuário {u['usuario']} removido com sucesso!")
                            st.rerun()
                    else:
                        st.caption("*(Administrador principal)*")
        else:
            st.info("Nenhum usuário encontrado.")

    with aba_adm2:
        st.subheader("Auditoria de Certidões Gerais")
        try:
            pastas_usuarios = supabase.storage.from_("certidoes_usuarios").list()
        except:
            pastas_usuarios = []
            
        if not pastas_usuarios:
            st.info("Nenhuma pasta de certidão encontrada na nuvem.")
        else:
            for pasta in pastas_usuarios:
                nome_oficial = pasta["name"]
                if nome_oficial and nome_oficial != ".emptyFolder":
                    st.markdown(f"### 📂 Oficial: `{nome_oficial}`")
                    try:
                        arquivos_oficial = supabase.storage.from_("certidoes_usuarios").list(nome_oficial)
                    except:
                        arquivos_oficial = []
                    certidoes_validas = [f for f in arquivos_oficial if f["name"] != ".emptyFolder" and f["name"] != ""]
                    
                    if not certidoes_validas:
                        st.caption("Nenhuma certidão gerada.")
                    else:
                        for arq in certidoes_validas:
                            c_arq_nome, c_btn_dl, c_btn_del = st.columns([4, 2, 2])
                            with c_arq_nome: st.text(arq["name"])
                            with c_btn_dl:
                                if st.button("📥 Baixar", key=f"dl_adm_f_{nome_oficial}_{arq['name']}", use_container_width=True):
                                    file_bytes = supabase.storage.from_("certidoes_usuarios").download(f"{nome_oficial}/{arq['name']}")
                                    st.download_button(
                                        label="Confirmar",
                                        data=file_bytes,
                                        file_name=arq["name"],
                                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                        key=f"btn_dl_real_{nome_oficial}_{arq['name']}",
                                        use_container_width=True
                                    )
                            with c_btn_del:
                                if st.button("🗑️ Excluir", key=f"del_adm_f_{nome_oficial}_{arq['name']}", use_container_width=True):
                                    supabase.storage.from_("certidoes_usuarios").remove([f"{nome_oficial}/{arq['name']}"])
                                    st.success("Excluído!")
                                    st.rerun()
                    st.divider()

# ==========================================
# 7. TELA: GERADOR DE CERTIDÃO (CORAÇÃO DA MELHORIA)
# ==========================================
elif menu == "📝 Gerar Certidão":
    st.title("Gerador de Certidão Negativa")
    
    if not dados_usuario.get("nome"):
        st.warning("⚠️ Você ainda não configurou seu perfil! Vá em 'Meu Perfil' antes de gerar certidões.")
        st.stop()

    # --- INPUT DO MODELO ---
    tipo_certidao = st.selectbox(
        "Selecione o Modelo de Certidão:", 
        ["Certidão Negativa Detalhada", "Certidão Negativa Simples (Opções Rápidas)"]
    )
    
    st.divider()

    # --- 📝 FORMULÁRIO GERAL COMPACTADO E REORGANIZADO ---
    # Linha 1: 🚀 MELHORIA: Agrupados Mandado, Processo, Ano, Comarca em UMA LINHA (4 Colunas)
    c_mandado, c_proc, c_ano, c_comarca = st.columns([1, 1, 0.5, 1])
    with c_mandado:
        mandado = st.text_input("Mandado:", placeholder="Ex: 01", key="mandado_det")
    with c_proc:
        processo = st.text_input("Informe o Processo:", placeholder="Ex: 4400281-16", key="processo_det")
    with c_ano:
        ano = st.text_input("Ano:", placeholder="Ex: 2026", key="ano_det")
    with c_comarca:
        comarca = st.text_input("Código Comarca:", value="0245", placeholder="Ex: 0245", key="comarca_det")

    # Linha 2: 🚀 MELHORIA: Agrupados Endereço e Pessoa procurada em UMA LINHA (2 Colunas, mais proporcionais)
    c_end, c_pes = st.columns([1.5, 1])
    with c_end:
        endereco = st.text_input("Endereço (opcional):", placeholder="Se vazio: 'informado no mesmo'", key="endereco_det")
    with c_pes:
        pessoa = st.text_input("Pessoa procurada:", placeholder="Deixe vazio para termo genérico", key="pessoa_det")

    st.markdown("---")
    st.subheader("Dias e Horários das Diligências")
    
    # --- 📝 SEÇÃO DIAS E HORÁRIOS - COMPACTADA DRASTICAMENTE (6 LINHAS -> 2 LINHAS) ---
    # 🚀 MELHORIA: Em vez de 3 linhas completas (Dia/Hora), criamos uma legenda compacta e duas linhas de campos.
    
    # Legendas compactas
    cl_1, cl_2, cl_3 = st.columns(3)
    cl_1.caption("Diligência 1")
    cl_2.caption("Diligência 2")
    cl_3.caption("Diligência 3")

    # Linha única para todos os DIAS (3 Colunas)
    cd1, cd2, cd3 = st.columns(3)
    with cd1:
        d1 = st.text_input("Dia 1", placeholder="Ex: 08/08", label_visibility="collapsed", key="d1_det")
    with cd2:
        d2 = st.text_input("Dia 2", placeholder="Ex: 11/08", label_visibility="collapsed", key="d2_det")
    with cd3:
        d3 = st.text_input("Dia 3", placeholder="Ex: 12/08", label_visibility="collapsed", key="d3_det")

    # Linha única para todas as HORAS (3 Colunas)
    ch1, ch2, ch3 = st.columns(3)
    with ch1:
        h1 = st.text_input("Hora 1", placeholder="Ex: 14:55hs", label_visibility="collapsed", key="h1_det")
    with ch2:
        h2 = st.text_input("Hora 2", placeholder="Ex: 16:58hs", label_visibility="collapsed", key="h2_det")
    with ch3:
        h3 = st.text_input("Hora 3", placeholder="Ex: 11:15hs", label_visibility="collapsed", key="h3_det")

    st.divider()

    # ==========================================
    # OPÇÃO A: CERTIDÃO DETALHADA
    # ==========================================
    if tipo_certidao == "Certidão Negativa Detalhada":
        st.subheader("Motivos da Negativa")
        sit_c1, sit_c2 = st.columns(2)
        with sit_c1:
            nao_loc_dest = st.checkbox("Destinatário não localizado", key="nao_loc_dest")
        with sit_c2:
            nao_loc_bens = st.checkbox("Bem(ns) não localizado(s)", key="nao_loc_bens")

        motivos_selecionados = []
        with st.expander("📌 Selecionar Motivos Detalhados (Se necessário)", expanded=False):
            # ... (lista de motivos permanece igual, mas compactada pelo CSS global)
            motivos_list = [
                "mudou-se", "não reside", "é desconhecido", "dificilmente fica ali", "trabalha em tempo integral",
                "não trabalha no local", "está viajando", "local inabitado", "antigo(a) inquilino(a)", 
                "antigo(a) morador(a)", "antigo(a) proprietário(a)", "rotatividade de inquilinos",
                "Repassado para terceiros", "internado", "transferido", "encontra-se preso",
                "faleceu", "faliu", "não exerce atividades", "local fechado", 
                "número não localizado", "rua/av não localizada", "ap/bloco não localizado", 
                "aparece esporadicamente", "utiliza endereço para correspondências",
                "sem condições psíquicas de entender conteúdo mandado",
                "guarnecem a residência amparados pela Lei 8.009/90",
                "são insuficientes para saldar o débito"
            ]
            cols_mot = st.columns(2)
            for idx, m in enumerate(motivos_list):
                with cols_mot[idx % 2]:
                    if st.checkbox(m, key=f"mot_det_{idx}"):
                        motivos_selecionados.append(m)

        st.markdown("---")
        # Informações sobre o informante
        relacoes_selecionadas = []
        nao_sabe_selecionados = []
        sabe_tel = ""
        sabe_end = ""
        
        with st.expander("👤 Informações sobre o Informante (Se houver)", expanded=False):
            nome_inf_det = st.text_input("Nome do Sr(a):", placeholder="Vazio se não houver informante", key="nome_inf_det")

            st.caption("Relação / Qualidade:")
            relacoes_list = [
                "morador", "proprietário", "inquilino", "funcionário", "vizinho", "pai", "mãe",
                "padrasto", "madrasta", "filho", "irmão", "tio", "avô(ó)", "neto", "sobrinho",
                "primo", "transeunte", "viúvo", "ex", "esposo", "companheiro", "sogro", "enteado",
                "genro", "nora", "cunhado", "concunhado", "amigo"
            ]
            cols_rel = st.columns(3)
            for idx, r in enumerate(relacoes_list):
                with cols_rel[idx % 3]:
                    if st.checkbox(r, key=f"rel_det_{idx}"):
                        relacoes_selecionadas.append(r)

            st.write("**Não sabendo o informante indicar:**")
            nao_sabe_list = [
                "endereço completo", "paradeiro", "o dia/horário exato", 
                "telefone", "dia/horário de retorno", "o presídio", 
                "dados do óbito", "previsão de alta"
            ]
            cols_ns = st.columns(2)
            for idx, ns in enumerate(nao_sabe_list):
                with cols_ns[idx % 2]:
                    if st.checkbox(ns, key=f"ns_det_{idx}"):
                        nao_sabe_selecionados.append(ns)

            st.write("**Sabendo o informante indicar:**")
            sabe_tel = st.text_input("Telefone indicado:", key="sabe_tel_det")
            sabe_end = st.text_input("Endereço correto indicado:", key="sabe_end_det")
        
        with st.expander("📝 Certificações Adicionais", expanded=False):
            cert_extras = []
            if st.checkbox("Procurei informações com moradores e não obtive êxito.", key="cert_vizinhos_det"):
                cert_extras.append("procurei obter informações junto aos moradores/vizinhos locais e não obtive êxito.")
            if st.checkbox("Cópia do mandado com informante para ciência do prazo/data.", key="cert_copia_det"):
                cert_extras.append("devido à importância do mandado e da dificuldade de encontrar a pessoa procurada, deixei a cópia do mandado com o(a) senhor(a) acima mencionado(a) para que a parte/testemunha tome ciência do prazo/data que deverá comparecer em juízo.")
            if st.checkbox("Imóvel residencial contém apenas móveis/utensílios comuns.", key="cert_moveis_det"):
                cert_extras.append("o imóvel é residencial e contém apenas móveis e utensílios domésticos que guarnecem a residência do réu.")

            observacoes_det = st.text_area("Observações Livres:", key="obs_livres_det")

        st.divider()

        # Botão de Geração
        if st.button("Salvar na Nuvem / Gerar DOCX (Detalhada)", type="primary", use_container_width=True, key="btn_gerar_docx_det"):
            with st.spinner("Gerando detalhada..."):
                # Lógica de montagem do DOCX detalhado
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
                paragrafo = f"Certifico que, em cumprimento ao mandado anexo, desloquei-me {txt_endereco}{texto_data_hora} onde deixei de cumprir o ato emanado no mandado{txt_pessoa}, uma vez que "
                sits = []
                if nao_loc_dest: sits.append("o destinatário do mandado não foi localizado")
                if nao_loc_bens: sits.append("o(s) bem(ns) indicados não foi(ram) localizado(s)")
                paragrafo += " e ".join(sits) + ". " if sits else "não foi possível a sua realização. "
                if motivos_selecionados:
                    paragrafo += f"Constatou-se no local que o(a) mesmo(a) {', '.join(motivos_selecionados)}. "
                if nome_inf_det or relacoes_selecionadas:
                    nome_str = nome_inf_det if nome_inf_det else "pessoa não identificada"
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
                if cert_extras: paragrafo += f"Certifico também que {'; '.join(cert_extras)}. "
                if observacoes_det: paragrafo += f"{observacoes_det.strip()} "
                doc = Document()
                style = doc.styles['Normal']; font = style.font; font.name = 'Times New Roman'; font.size = Pt(12)
                try:
                    cabecalho_bytes = supabase.storage.from_("imagens_sistema").download("cabecalho.png")
                    p_img_cabecalho = doc.add_paragraph(); p_img_cabecalho.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    p_img_cabecalho.add_run().add_picture(BytesIO(cabecalho_bytes), width=Cm(16))
                except: pass
                if processo:
                    texto_processo = f"Processo: {processo}"
                    if ano: texto_processo += f".{ano}.8.13.{comarca}"
                    doc.add_paragraph(texto_processo)
                if mandado: doc.add_paragraph(f"Mandado nº: {mandado}")
                doc.add_paragraph("")
                p_titulo = doc.add_paragraph(); run_titulo = p_titulo.add_run("CERTIDÃO NEGATIVA"); run_titulo.bold = True; run_titulo.font.size = Pt(16); p_titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
                doc.add_paragraph("")
                p_corpo = doc.add_paragraph(paragrafo.strip()); p_corpo.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY; p_corpo.paragraph_format.first_line_indent = Pt(35.4); p_corpo.paragraph_format.line_spacing = 1.5 
                doc.add_paragraph("")
                p_fechamento = doc.add_paragraph("Devolvo o mandado para os devidos fins. O referido é verdade. Dou fé."); p_fechamento.alignment = WD_ALIGN_PARAGRAPH.CENTER
                hoje = datetime.datetime.utcnow() - datetime.timedelta(hours=3); meses = ["janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"]
                local_data = dados_usuario.get("matricula", "").split(":")[0].strip() or "Santa Luzia"
                doc.add_paragraph(f"{local_data}, {hoje.day} de {meses[hoje.month - 1]} de {hoje.year}.").alignment = WD_ALIGN_PARAGRAPH.CENTER
                doc.add_paragraph("")
                try:
                    assinatura_bytes = supabase.storage.from_("assinaturas_usuarios").download(f"{usuario_atual}.png")
                    p_img_assinatura = doc.add_paragraph(); p_img_assinatura.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    p_img_assinatura.add_run().add_picture(BytesIO(assinatura_bytes), width=Cm(6))
                except: pass 
                p_assinatura = doc.add_paragraph(); p_assinatura.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run_nome = p_assinatura.add_run(f"{dados_usuario['nome']}\n"); run_nome.bold = True; run_nome.font.size = Pt(8)
                run_cargo = p_assinatura.add_run(f"{dados_usuario['cargo']}\n"); run_cargo.font.size = Pt(8)
                run_matricula = p_assinatura.add_run(f"{dados_usuario['matricula']}"); run_matricula.font.size = Pt(8)
                buffer = BytesIO(); doc.save(buffer); buffer.seek(0)
                data_arquivo = hoje.strftime("%d-%m-%Y_%Hh%M")
                nome_arquivo = f"Certidao_Negativa_{processo}_{data_arquivo}.docx" if processo else f"Certidao_Negativa_{data_arquivo}.docx"
                supabase.storage.from_("certidoes_usuarios").upload(file=buffer.getvalue(), path=f"{usuario_atual}/{nome_arquivo}", file_options={"content-type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document"})
            st.success(f"✅ Certidão detalhada salva na sua conta na Nuvem!")
            st.download_button(label="📥 Baixar Documento Word Agora", data=buffer, file_name=nome_arquivo, mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", type="primary", use_container_width=True, key="btn_dl_real_det")

    # ==========================================
    # OPÇÃO B: CERTIDÃO SIMPLES (MELHORIA TOUCH E WIDE)
    # ==========================================
    elif tipo_certidao == "Certidão Negativa Simples (Opções Rápidas)":
        
        # --- INPUTS ---
        st.subheader("Situação Principal")
        situacao_simples = st.radio(
            "Selecione uma opção:", 
            ["Local Fechado", "Pessoa Não Encontrada", "Não Localizei a Pessoa"],
            index=None, key="sit_radio_simples"
        )

        st.markdown("---")
        # Informante
        obteve_inf_simples = st.radio("Obteve Informações?", ["Sim", "Não", "NQI"], index=None, key="obteve_inf_radio_simples")
        nome_inf_simples = st.text_input("Nome do informante:", disabled=(obteve_inf_simples != "Sim"), key="nome_inf_input_simples")

        st.markdown("---")
        st.write("**Detalhes das Informações Obtidas:**")
        
        # Motivo
        st.caption("Motivo")
        motivo_simples = st.radio(
            "Selecione uma opção:", 
            ["Mudou-se", "Não Reside", "Não fica ali", "Não trabalha ali", "Falecido"], 
            index=None, key="motivo_radio_simples"
        )
        
        # Não sabe
        st.caption("O que não sabe indicar?")
        nao_sabe_simples = st.radio(
            "Selecione uma opção:", 
            ["Não Conhece", "Não sabe informar", "Não sabe endereço"], 
            index=None, key="naosabe_radio_simples"
        )
        
        # Paradeiro
        st.caption("Paradeiro")
        paradeiro_simples = st.radio(
            "Selecione uma opção:", 
            ["Não sabe o paradeiro", "Incerto e Não Sabido"], 
            index=None, key="paradeiro_radio_simples"
        )

        st.markdown("---")
        st.write("**Condições Extras**")
        # Condições do local
        condicao_simples = st.radio(
            "Selecione uma opção:", 
            ["Local Perigoso", "Medo Processo", "Zona Rural", "Blocos", "Chuva"], 
            index=None, key="condicao_radio_simples"
        )

        st.markdown("---")
        # Observações Extra Compacta (Área de texto compactada pelo CSS global)
        observacoes_simples = st.text_area("Observações Extras:", height=60, key="obs_simples")
        st.divider()

        # --- LÓGICA DO BOTÃO GERAR SIMPLES (Restaurada e Dinâmica) ---
        if st.button("Salvar na Nuvem / Gerar DOCX (Simples)", type="primary", use_container_width=True, key="btn_gerar_simples"):
            with st.spinner("Construindo certidão simples e salvando na nuvem..."):
                # Lógica de montagem do DOCX simples
                dias_validos = [d for d in [d1, d2, d3] if d]
                horas_validas = [h for h in [h1, h2, h3] if h]
                texto_data_hora = ""
                if len(dias_validos) == 1:
                    texto_data_hora = f", onde às {horas_validas[0]}, do dia {dias_validos[0]},"
                elif len(dias_validos) > 1:
                    str_horas = ", ".join(horas_validas[:-1]) + f" e {horas_validas[-1]}"
                    str_dias = ", ".join(dias_validos[:-1]) + f" e {dias_validos[-1]}"
                    texto_data_hora = f", onde às {str_horas}, dos dias {str_dias},"
                txt_endereco = f"à {endereco}" if endereco else "ao endereço informado no mesmo"
                txt_pessoa = f" a pessoa, Sr(a). {pessoa}" if pessoa else "a pessoa referida no mandado"
                txt_situacao = ""
                if situacao_simples == "Local Fechado": txt_situacao = "porque o local foi encontrado fechado e mesmo após chamar várias vezes, ninguém atendeu. "
                elif situacao_simples == "Pessoa Não Encontrada": txt_situacao = "porque não a encontrei no local. "
                elif situacao_simples == "Não Localizei a Pessoa": txt_situacao = "porque não a localizei. "
                paragrafo_unico = f"Certifico e dou fé que, em cumprimento ao mandado anexo, dirigi-me {txt_endereco}{texto_data_hora} e, deixei de citar/intimar/notificar {txt_pessoa}, {txt_situacao}"
                if obteve_inf_simples == "Sim": paragrafo_unico += f"Conforme informações obtidas no local com Sr.(a) {nome_inf_simples}, informou que, "
                elif obteve_inf_simples == "Não": paragrafo_unico += "Procurei obter informações junto aos moradores vizinhos locais, e não obtive êxito, uma vez que ninguém forneceu informações. "
                elif obteve_inf_simples == "NQI": paragrafo_unico += "Conforme informações prestadas pelo seu vizinho(a), que não quis se identificar, este afirmou que "
                if obteve_inf_simples in ["Sim", "NQI"]:
                    if motivo_simples == "Mudou-se": paragrafo_unico += "a pessoa procurada não reside mais no local, tendo se mudado sem deixar meios para contato; "
                    elif motivo_simples == "Não Reside": paragrafo_unico += "a pessoa procurada não reside no local referido; "
                    elif motivo_simples == "Não fica ali": paragrafo_unico += "a pessoa procurada reside no local, mas quase não fica no mesmo, onde nos dias e horários acima não foi localizada; "
                    elif motivo_simples == "Não trabalha ali": paragrafo_unico += "a pessoa procurada não trabalha no local; "
                    elif motivo_simples == "Falecido": paragrafo_unico += "a pessoa procurada já se encontra falecida. "
                    if nao_sabe_simples == "Não Conhece": paragrafo_unico += "não conhece a pessoa procurada, não sabendo informar o local/horário para encontrá-la. "
                    elif nao_sabe_simples == "Não sabe informar": paragrafo_unico += "que não sabe informar o dia e horário para encontrá-lo(a). "
                    elif nao_sabe_simples == "Não sabe endereço": paragrafo_unico += "que não sabe informar o endereço para encontrá-lo(a). "
                    if paradeiro_simples == "Não sabe o paradeiro": paragrafo_unico += "não sabe informar seu paradeiro, bem como o local para encontrá-lo. "
                    elif paradeiro_simples == "Incerto e Não Sabido": paragrafo_unico += "Certifico assim, que PESSOA PROCURADA SE ENCONTRA EM LOCAL INCERTO E NÃO SABIDO. "
                obs_extra = ""
                if condicao_simples == "Chuva": obs_extra = "Certifico que a execução restou dificultada em virtude das adversas condições meteorológicas no momento do ato, caracterizadas por intensa precipitação pluviométrica. "
                elif condicao_simples == "Local Perigoso": obs_extra = "local é conhecidamente de grande periculosidade, os moradores ficam receosos de envolvimento. "
                elif condicao_simples == "Zona Rural": obs_extra = "zona rural com difícil acesso, localização difícil, numeração irregular com muitas casas sem números. "
                elif condicao_simples == "Blocos": obs_extra = "local é um condomínio com vários blocos, portaria vazia, interfone aparentemente não está funcionando. "
                elif condicao_simples == "Medo Processo": obs_extra = "Procurei informações com vizinhos, mas os moradores ficam receosos de envolvimento com o processo. "
                if obs_extra or observacoes_simples: paragrafo_unico += obs_extra + (" " + observacoes_simples if observacoes_simples else "")
                doc = Document(); style = doc.styles['Normal']; font = style.font; font.name = 'Times New Roman'; font.size = Pt(12)
                try:
                    cabecalho_bytes = supabase.storage.from_("imagens_sistema").download("cabecalho.png")
                    p_img_cabecalho = doc.add_paragraph(); p_img_cabecalho.alignment = WD_ALIGN_PARAGRAPH.CENTER; p_img_cabecalho.add_run().add_picture(BytesIO(cabecalho_bytes), width=Cm(16))
                except: pass
                if processo:
                    texto_processo = f"Processo: {processo}"
                    if ano: texto_processo += f".{ano or '2026'}.8.13.{comarca}"
                    doc.add_paragraph(texto_processo)
                if mandado: doc.add_paragraph(f"Mandado nº: {mandado}")
                doc.add_paragraph(""); p_titulo = doc.add_paragraph(); run_titulo = p_titulo.add_run("CERTIDÃO"); run_titulo.bold = True; run_titulo.font.size = Pt(16); p_titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER; doc.add_paragraph("")
                doc.add_paragraph(paragrafo_unico.strip()).alignment = WD_ALIGN_PARAGRAPH.JUSTIFY; doc.paragraphs[-1].paragraph_format.first_line_indent = Pt(35.4); doc.add_paragraph("")
                doc.add_paragraph("Devolvo o mandado para os devidos fins. É verdade. Dou fé.").alignment = WD_ALIGN_PARAGRAPH.CENTER
                hoje = datetime.datetime.utcnow() - datetime.timedelta(hours=3); meses = ["janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"]
                doc.add_paragraph(f"{dados_usuario.get('matricula', '').split(':')[0].strip() or 'Santa Luzia'}, {hoje.day} mes {hoje.month - 1} de {hoje.year}.").alignment = WD_ALIGN_PARAGRAPH.CENTER; doc.add_paragraph("")
                try:
                    assinatura_bytes = supabase.storage.from_("assinaturas_usuarios").download(f"{usuario_atual}.png")
                    p_img_assinatura = doc.add_paragraph(); p_img_assinatura.alignment = WD_ALIGN_PARAGRAPH.CENTER; p_img_assinatura.add_run().add_picture(BytesIO(assinatura_bytes), width=Cm(6))
                except: pass 
                p_assinatura = doc.add_paragraph(); p_assinatura.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run_nome = p_assinatura.add_run(f"{dados_usuario['nome']}\n"); run_nome.bold = True; run_nome.font.size = Pt(8)
                run_cargo = p_assinatura.add_run(f"{dados_usuario['cargo']}\n"); run_cargo.font.size = Pt(8)
                run_matricula = p_assinatura.add_run(f"{dados_usuario['matricula']}"); run_matricula.font.size = Pt(8)
                buffer = BytesIO(); doc.save(buffer); buffer.seek(0)
                data_arquivo = hoje.strftime("%d-%m-%Y_%Hh%M")
                nome_arquivo = f"Certidao_Simples_{processo}_{data_arquivo}.docx" if processo else f"Certidao_Simples_{data_arquivo}.docx"
                supabase.storage.from_("certidoes_usuarios").upload(file=buffer.getvalue(), path=f"{usuario_atual}/{nome_arquivo}", file_options={"content-type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document"})
            st.success(f"✅ Certidão simples salva na sua conta na Nuvem!")
            st.download_button(label="📥 Baixar Documento Word Agora", data=buffer, file_name=nome_arquivo, mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", type="primary", use_container_width=True, key="btn_dl_simples")
