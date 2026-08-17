import streamlit as st
import os
import hashlib
import zipfile
import subprocess
import tempfile
import re
import time
import uuid
from io import BytesIO
import datetime
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, Cm
from supabase import create_client, Client

# ==========================================
# FUNÇÕES AUXILIARES
# ==========================================
def limpar_cpf(cpf_str):
    """Remove tudo que não for número do CPF."""
    return re.sub(r'\D', '', cpf_str)

def formatar_data_completa(data_str, ano_padrao):
    """Recebe uma string de data (dd/mm ou dd/mm/aa) e retorna dd/mm/aaaa."""
    if not data_str or "/" not in data_str:
        return data_str
    
    partes = [p.strip() for p in data_str.split('/')]
    
    if len(partes) == 2:
        return f"{partes[0].zfill(2)}/{partes[1].zfill(2)}/{ano_padrao}"
    
    if len(partes) == 3:
        dia, mes, ano_part = partes[0].zfill(2), partes[1].zfill(2), partes[2]
        if len(ano_part) == 2:
            ano_part = "20" + ano_part
        return f"{dia}/{mes}/{ano_part}"
        
    return data_str

def converter_docx_para_pdf(docx_bytes):
    """Converte DOCX para PDF usando LibreOffice oculto."""
    with tempfile.TemporaryDirectory() as temp_dir:
        caminho_docx = os.path.join(temp_dir, "temp_certidao.docx")
        
        with open(caminho_docx, "wb") as f:
            f.write(docx_bytes)
        
        comando = [
            "libreoffice", "--headless", "--convert-to", "pdf",
            "--outdir", temp_dir, caminho_docx
        ]
        
        subprocess.run(comando, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        caminho_pdf = os.path.join(temp_dir, "temp_certidao.pdf")
        
        if os.path.exists(caminho_pdf):
            with open(caminho_pdf, "rb") as f:
                return f.read()
        return None

# ==========================================
# CONFIGURAÇÃO DA PÁGINA E BANCO DE DADOS
# ==========================================
st.set_page_config(page_title="Sistema de Certidões", layout="wide")

st.markdown("""
    <style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    
    .block-container { 
        padding-top: 2rem; 
        padding-bottom: 2rem; 
        max-width: 1200px;
    }
    div[data-testid="stVerticalBlock"] { 
        gap: 1rem !important; 
    }

    .main h1, h1[data-testid="stHeader"] { 
        font-size: 26px !important; 
        font-weight: 700;
        text-align: center !important; 
        color: #1E293B;
        margin-top: 0rem !important; 
        margin-bottom: 1.5rem !important; 
        display: block;
        width: 100%;
    }

    /* Ajuste do fundo das caixas */
    div[data-testid="stExpander"], div.stTextInput, div.stSelectbox, div.stRadio {
        background-color: #FFFFFF !important;
        border-radius: 10px;
        padding: 0.2rem;
    }

    /* CORREÇÃO DO MODO ESCURO: Força o texto e rótulos a ficarem escuros dentro do fundo branco */
    div[data-testid="stExpander"] *, 
    div.stTextInput *, 
    div.stSelectbox *, 
    div.stRadio * {
        color: #1E293B !important;
    }

    /* Ajuste das caixas de digitação */
    input[type="text"], input[type="password"] {
        border-radius: 8px !important;
        border: 1px solid #CBD5E1 !important;
        background-color: #FFFFFF !important;
        color: #1E293B !important; /* Texto digitado sempre escuro */
    }
    input[type="text"]:focus, input[type="password"]:focus {
        border-color: #0F172A !important;
        box-shadow: 0 0 0 1px #0F172A !important;
    }

    button[kind="primary"] {
        background-color: #0F172A !important;
        color: #FFFFFF !important;
        border-radius: 8px !important;
        font-weight: 600 !important;
        border: none !important;
        transition: all 0.3s ease !important;
    }
    button[kind="primary"]:hover {
        background-color: #1E293B !important;
        box-shadow: 0 4px 12px rgba(15, 23, 42, 0.2) !important;
    }

    section[data-testid="stSidebar"] { 
        background-color: #F8FAFC !important;
        border-right: 1px solid #E2E8F0;
        width: 18rem !important; 
    }
    section[data-testid="stSidebar"] div[data-testid="stVerticalBlock"] { 
        gap: 0.8rem !important; 
    }
    </style>
""", unsafe_allow_html=True)

@st.cache_resource
def iniciar_conexao():
    url = st.secrets["SUPABASE_URL"]
    key = st.secrets["SUPABASE_KEY"]
    return create_client(url, key)

supabase: Client = iniciar_conexao()

def gerar_hash_senha(senha):
    return hashlib.sha256(senha.encode()).hexdigest()

# ==========================================
# CONTROLE DE SESSÃO E LOGIN (VIA URL)
# ==========================================
if "usuario_logado" not in st.session_state or st.session_state["usuario_logado"] is None:
    usuario_url = st.query_params.get("user")
    if usuario_url:
        st.session_state["usuario_logado"] = usuario_url
    else:
        st.session_state["usuario_logado"] = None

if st.session_state["usuario_logado"] is None:
    st.title("⚖️ Sistema de Certidões - TJMG")
    
    aba_login, aba_cadastro = st.tabs(["Entrar", "Criar Nova Conta"])
    
    with aba_login:
        cpf_input_bruto = st.text_input("CPF (Apenas números):", key="log_usr_input")
        senha_login = st.text_input("Senha:", type="password", key="log_pwd_input")
        
        if st.button("Entrar", type="primary", use_container_width=True, key="btn_entrar"):
            usuario_login = limpar_cpf(cpf_input_bruto)
            
            if not usuario_login or len(usuario_login) != 11:
                st.warning("⚠️ O login deve ser um CPF válido contendo 11 números.")
            elif senha_login:
                resposta = supabase.table("banco_usuarios").select("*").eq("usuario", usuario_login).execute()
                
                if len(resposta.data) > 0:
                    dados_bd = resposta.data[0]
                    senha_criptografada = gerar_hash_senha(senha_login)
                    if dados_bd["senha"] == senha_criptografada:
                        st.session_state["usuario_logado"] = usuario_login
                        st.query_params["user"] = usuario_login
                        st.rerun()
                    else:
                        st.error("Senha incorreta!")
                else:
                    st.error("CPF não cadastrado no sistema.")
            else:
                st.warning("Preencha a senha.")
                
    with aba_cadastro:
        st.subheader("Criar Nova Conta")
        st.write("Preencha os dados abaixo para iniciar seu período gratuito.")
        
        cpf_cad_bruto = st.text_input("CPF (Apenas números):", key="cad_cpf_input")
        senha_cad = st.text_input("Crie uma Senha:", type="password", key="cad_pwd_input")
        senha_cad_conf = st.text_input("Confirme a Senha:", type="password", key="cad_pwd_conf_input")
        
        st.markdown("---")
        
        # O texto claro do Termo de Aceite e Regra do 1 Ano
        texto_termos = """
        **Termos de Uso e Assinatura:**
        Declaro que li e concordo com as normas de uso do sistema. 
        Estou ciente de que esta ferramenta será oferecida de forma **100% gratuita pelo prazo de 1 (um) ano** a partir da data deste cadastro. 
        Após este período de testes, para continuar utilizando o sistema, será cobrada uma mensalidade no valor de **R$ 30,00**. Não haverá cobrança automática sem aviso prévio.
        """
        st.info(texto_termos)
        aceite_termos = st.checkbox("Li e aceito os Termos de Uso e o período de gratuidade de 1 ano.", key="chk_termos")
        
        if st.button("Criar Conta e Iniciar Teste", type="primary", use_container_width=True, key="btn_cadastrar"):
            usuario_cad = limpar_cpf(cpf_cad_bruto)
            
            if not usuario_cad or len(usuario_cad) != 11:
                st.error("⚠️ O CPF deve conter 11 números válidos.")
            elif not senha_cad or len(senha_cad) < 6:
                st.error("⚠️ A senha deve ter pelo menos 6 caracteres.")
            elif senha_cad != senha_cad_conf:
                st.error("⚠️ As senhas não coincidem.")
            elif not aceite_termos:
                st.error("⚠️ Você deve marcar a caixa aceitando os Termos de Uso para criar a conta.")
            else:
                # 1. Verifica se o CPF já existe
                busca = supabase.table("banco_usuarios").select("usuario").eq("usuario", usuario_cad).execute()
                if len(busca.data) > 0:
                    st.error("❌ Este CPF já está cadastrado no sistema.")
                else:
                    # 2. Calcula as datas
                    hoje = datetime.datetime.utcnow() - datetime.timedelta(hours=3)
                    data_cadastro = hoje.date()
                    # Adiciona 365 dias (1 ano)
                    vencimento_trial = data_cadastro + datetime.timedelta(days=365)
                    
                    # 3. Insere no banco
                    senha_cripto = gerar_hash_senha(senha_cad)
                    novo_usuario = {
                        "usuario": usuario_cad,
                        "senha": senha_cripto,
                        "data_cadastro": str(data_cadastro),
                        "vencimento_trial": str(vencimento_trial),
                        "status_assinatura": "trial",
                        "aceitou_termos": True
                    }
                    
                    supabase.table("banco_usuarios").insert(novo_usuario).execute()
                    
                    st.success("✅ Conta criada com sucesso! Faça login na aba ao lado para configurar seu perfil.")
                    time.sleep(3)
                    st.rerun()
                    
    st.stop()

# ==========================================
# DADOS DO USUÁRIO E MENU LATERAL
# ==========================================
usuario_atual = st.session_state["usuario_logado"]
resposta_usuario = supabase.table("banco_usuarios").select("*").eq("usuario", usuario_atual).execute()

# Trava de segurança: se a sessão atual não existir mais no banco de dados, força o logout
if not resposta_usuario.data:
    st.session_state["usuario_logado"] = None
    st.query_params.clear()
    st.rerun()

dados_usuario = resposta_usuario.data[0]

# ==========================================
# VERIFICAÇÃO DE ASSINATURA / TRIAL
# ==========================================
hoje_verificacao = (datetime.datetime.utcnow() - datetime.timedelta(hours=3)).date()
vencimento_str = dados_usuario.get("vencimento_trial")

# Se o usuário tiver data de vencimento preenchida, verificamos
if vencimento_str:
    try:
        vencimento_obj = datetime.datetime.strptime(vencimento_str, "%Y-%m-%d").date()
        if hoje_verificacao > vencimento_obj and dados_usuario.get("status_assinatura") != "ativo":
            st.warning("⚠️ **Seu período de 1 ano gratuito expirou.**")
            st.write("Obrigado por utilizar o Gerador de Certidões - TJMG durante este último ano!")
            st.write("Para continuar economizando tempo e gerando suas certidões com facilidade, ative sua assinatura mensal por apenas **R$ 30,00**.")
            
            # Futuramente, aqui você colocará o link/botão do Mercado Pago ou Stripe
            st.button("Assinar agora via PIX / Cartão", type="primary")
            
            if st.button("Sair da conta"):
                st.session_state["usuario_logado"] = None
                st.query_params.clear()
                st.rerun()
                
            st.stop() # Isso impede que o resto do sistema e o menu lateral carreguem
    except:
        pass # Caso haja erro de conversão de data, permite o acesso temporário

with st.sidebar:
    st.write(f"👤 Olá, **{usuario_atual}**!")
    st.divider()
    
    opcoes_menu = ["📝 Gerar Certidão", "📂 Minhas Certidões", "⚙️ Meu Perfil"]
    # CPF do Administrador Atualizado
    if usuario_atual == "05042687670":
        opcoes_menu.append("🛡️ Painel do Administrador")
        
    menu = st.radio("Navegação:", opcoes_menu)
    st.divider()
    
    if st.button("Sair (Logout)", key="btn_logout"):
        st.session_state["usuario_logado"] = None
        st.query_params.clear()
        st.rerun()

# ==========================================
# TELA: MEU PERFIL
# ==========================================
if menu == "⚙️ Meu Perfil":
    st.title("⚙️ Configurar Meu Perfil")
    st.write("Estes dados são **obrigatórios** para que você possa gerar certidões.")
    
    novo_nome = st.text_input("Nome Completo:", value=dados_usuario.get("nome", ""), placeholder="Ex: Rafael", key="input_perfil_nome")
    novo_cargo = st.text_input("Cargo:", value=dados_usuario.get("cargo", ""), placeholder="Ex: Oficial de Justiça - TJMG", key="input_perfil_cargo")
    nova_matricula = st.text_input("Matrícula (ex: PJPI: 12345):", value=dados_usuario.get("matricula", ""), key="input_perfil_matricula")
    novo_email = st.text_input("E-mail Profissional:", value=dados_usuario.get("email", ""), placeholder="Ex: rafael@tjmg.jus.br", key="input_perfil_email")
    
    c_cid, c_est = st.columns([3, 1])
    with c_cid:
        nova_cidade = st.text_input("Comarca / Cidade de Lotação:", value=dados_usuario.get("cidade", ""), placeholder="Ex: Belo Horizonte", key="input_perfil_cidade")
    with c_est:
        novo_estado = st.text_input("Estado (Sigla):", value=dados_usuario.get("estado", ""), max_chars=2, placeholder="Ex: MG", key="input_perfil_estado").upper()
    
    st.markdown("---")
    st.write("**Sua Assinatura (Fundo branco ou transparente):**")
    
    try:
        assinatura_salva = supabase.storage.from_("assinaturas_usuarios").download(f"{usuario_atual}.png")
        if assinatura_salva:
            st.success("✅ **Você já possui uma assinatura salva no sistema:**")
            st.image(assinatura_salva, width=250)
            st.write("*(Envie um novo arquivo abaixo apenas se desejar substituí-la)*")
    except:
        st.warning("❌ **Nenhuma assinatura salva.** Por favor, envie sua assinatura abaixo.")

    arquivo_assinatura = st.file_uploader("Envie a foto da sua assinatura", type=["png", "jpg", "jpeg"], key="uploader_perfil")
    
    if st.button("💾 Salvar Perfil", type="primary", use_container_width=True, key="btn_salvar_perfil"):
        if not (novo_nome and novo_cargo and nova_matricula and novo_email and nova_cidade and novo_estado):
            st.error("⚠️ Atenção: Todos os campos de texto são obrigatórios. Preencha todos antes de salvar!")
        else:
            supabase.table("banco_usuarios").update({
                "nome": novo_nome,
                "cargo": novo_cargo,
                "matricula": nova_matricula,
                "email": novo_email,
                "cidade": nova_cidade,
                "estado": novo_estado
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
            time.sleep(2) 
            st.rerun()

# ==========================================
# TELA: MINHAS CERTIDÕES
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
        c_filtro, c_btn1, c_btn2 = st.columns([2, 1.5, 1.5])
        with c_filtro:
            ativar_filtro = st.checkbox("Filtrar por data", key="ativar_filtro_data")
            hoje_real = datetime.datetime.utcnow() - datetime.timedelta(hours=3)
            data_filtro = st.date_input("Escolha a data:", value=hoje_real.date(), format="DD/MM/YYYY", disabled=not ativar_filtro, label_visibility="collapsed")
            
        arquivos_filtrados = []
        for item in arquivos:
            try:
                data_str = item["created_at"].replace("Z", "+00:00")
                data_obj = datetime.datetime.fromisoformat(data_str)
                data_br_obj = data_obj.replace(tzinfo=None) - datetime.timedelta(hours=3)
                data_br_date = data_br_obj.date()
                data_br = data_br_obj.strftime("%d/%m/%Y às %H:%M")
            except:
                data_br_date = None
                data_br = "Data desconhecida"
                
            item['data_br_date'] = data_br_date
            item['data_br'] = data_br
            
            if ativar_filtro and data_br_date != data_filtro:
                continue
            arquivos_filtrados.append(item)
            
        with c_btn1:
            if st.button("✓ Marcar Todos Abaixo", use_container_width=True):
                for arq in arquivos_filtrados:
                    st.session_state[f"chk_file_{arq['name']}"] = True
                st.rerun()
        with c_btn2:
            if st.button("✕ Desmarcar Todos", use_container_width=True):
                for arq in arquivos_filtrados:
                    st.session_state[f"chk_file_{arq['name']}"] = False
                st.rerun()
                
        arquivos_filtrados.sort(key=lambda x: x["created_at"], reverse=True)
        st.divider()
        
        c_sel, c_nome, c_data = st.columns([1, 4, 3])
        c_sel.write("**Selecionar**")
        c_nome.write("**Nome do Arquivo**")
        c_data.write("**Data de Criação**")
        st.divider()
        
        arquivos_selecionados = []
        
        for item in arquivos_filtrados:
            c1, c2, c3 = st.columns([1, 4, 3])
            with c1:
                if st.checkbox("", key=f"chk_file_{item['name']}"):
                    arquivos_selecionados.append(item['name'])
            with c2:
                st.write(item['name'])
            with c3:
                st.write(item['data_br'])
                
        st.divider()
        
        if arquivos_selecionados:
            st.write(f"**{len(arquivos_selecionados)} arquivo(s) selecionado(s)**")
            c_btn1, c_btn2 = st.columns(2)
            
            with c_btn1:
                if st.button("📥 Preparar Download (ZIP)", type="primary", use_container_width=True, key="btn_zip_download"):
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
# TELA: PAINEL DO ADMINISTRADOR
# ==========================================
elif menu == "🛡️ Painel do Administrador":
    if usuario_atual != "05042687670":
        st.error("Acesso restrito apenas ao Administrador.")
        st.stop()
        
    st.title("🛡️ Painel de Administração")
    st.write("Área restrita para gestão de oficiais e auditoria de certidões em nuvem.")
    
    aba_adm1, aba_adm2 = st.tabs(["👥 Gerenciar Usuários", "📊 Auditoria de Certidões Gerais"])
    
    with aba_adm1:
        st.subheader("Oficiais Cadastrados no Sistema")
        res_todos = supabase.table("banco_usuarios").select("usuario, nome, cargo, matricula, email, cidade, estado").execute()
        usuarios_cadastrados = res_todos.data
        
        if usuarios_cadastrados:
            for u in usuarios_cadastrados:
                with st.expander(f"👤 Usuário/CPF: {u['usuario']} — Nome: {u.get('nome') or 'Não preenchido'}"):
                    st.write(f"**Cargo:** {u.get('cargo')}")
                    st.write(f"**Matrícula:** {u.get('matricula')}")
                    st.write(f"**E-mail:** {u.get('email') or 'Não informado'}")
                    st.write(f"**Comarca/Cidade:** {u.get('cidade') or 'Não informada'} / {u.get('estado') or '-'}")
                    
                    if u['usuario'] != usuario_atual:
                        if st.button(f"🗑️ Excluir usuário {u['usuario']}", key=f"del_adm_usr_{u['usuario']}"):
                            supabase.table("banco_usuarios").delete().eq("usuario", u['usuario']).execute()
                            st.success(f"Usuário {u['usuario']} removido com sucesso!")
                            st.rerun()
                    else:
                        st.caption("*(Esta é a sua conta de Administrador principal)*")
        else:
            st.info("Nenhum usuário encontrado.")

    with aba_adm2:
        st.subheader("Certidões Geradas por Todos os Oficiais")
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
                    st.markdown(f"### 📂 Oficial (CPF): `{nome_oficial}`")
                    
                    try:
                        arquivos_oficial = supabase.storage.from_("certidoes_usuarios").list(nome_oficial)
                    except:
                        arquivos_oficial = []
                        
                    certioes_validas = [f for f in arquivos_oficial if f["name"] != ".emptyFolder" and f["name"] != ""]
                    
                    if not certioes_validas:
                        st.caption("Nenhuma certidão gerada por este oficial ainda.")
                    else:
                        for arq in certioes_validas:
                            c_arq_nome, c_btn_dl, c_btn_del = st.columns([4, 2, 2])
                            
                            with c_arq_nome:
                                st.text(arq["name"])
                                
                            with c_btn_dl:
                                if st.button("📥 Baixar", key=f"dl_adm_f_{nome_oficial}_{arq['name']}", use_container_width=True):
                                    file_bytes = supabase.storage.from_("certidoes_usuarios").download(f"{nome_oficial}/{arq['name']}")
                                    mime_tipo = "application/pdf" if arq["name"].endswith(".pdf") else "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                                    
                                    st.download_button(
                                        label="Confirmar",
                                        data=file_bytes,
                                        file_name=arq["name"],
                                        mime=mime_tipo,
                                        key=f"btn_dl_real_{nome_oficial}_{arq['name']}"
                                    )
                                    
                            with c_btn_del:
                                if st.button("🗑️ Excluir", key=f"del_adm_f_{nome_oficial}_{arq['name']}", use_container_width=True):
                                    supabase.storage.from_("certidoes_usuarios").remove([f"{nome_oficial}/{arq['name']}"])
                                    st.success("Excluído!")
                                    st.rerun()
                    st.divider()

# ==========================================
# TELA: GERADOR DE CERTIDÃO
# ==========================================
elif menu == "📝 Gerar Certidão":
    st.title("Gerador de Certidões - TJMG")
    
    if not (dados_usuario.get("nome") and dados_usuario.get("cargo") and dados_usuario.get("matricula") and dados_usuario.get("email") and dados_usuario.get("cidade") and dados_usuario.get("estado")):
        st.warning("⚠️ Acesso restrito! Vá em 'Meu Perfil' e preencha **todos os dados obrigatórios** (Nome, Cargo, Matrícula, E-mail, Comarca e Estado) para liberar a geração de certidões.")
        st.stop()

    cidade_certidao = dados_usuario.get("cidade").strip().title()
    estado_certidao = dados_usuario.get("estado").strip().upper()

    c_tipo, c_formato = st.columns([3, 1])
    with c_tipo:
        tipo_certidao = st.selectbox(
            "Selecione o Modelo de Certidão:", 
            [
                "Certidão Negativa Detalhada",
                "Certidão Positiva",
                "Certidão Positiva por Hora Certa"
            ]
        )
    with c_formato:
        formato_saida = st.radio(
            "Formato de exportação:", 
            ["Word (.docx)", "PDF (.pdf)"], 
            key="formato_global"
        )
        
    st.divider()

    c_mandado, c_proc, c_ano, c_comarca = st.columns([1, 2.5, 1, 1])
    
    with c_mandado:
        mandado = st.text_input("Mandado:", placeholder="Ex: 01", key="mandado_geral")
    with c_proc:
        processo = st.text_input("Processo:", placeholder="Ex: 4400281-16", key="processo_geral")
    with c_ano:
        ano = st.text_input("Ano:", value="2026", placeholder="Ex: 2026", key="ano_geral")
    with c_comarca:
        comarca = st.text_input("Cód. Comarca:", value="0245", placeholder="Ex: 0245", key="comarca_geral")

    c_end, c_pes = st.columns(2)
    with c_end:
        endereco = st.text_input("Endereço (opcional):", placeholder="Se vazio: 'informado no mesmo'", key="endereco_geral")
    with c_pes:
        pessoa = st.text_input("Pessoa procurada: (opcional)", placeholder="Deixe vazio para termo genérico", key="pessoa_geral")

    st.markdown("---")
    st.subheader("Data do Documento e Diligências")
    
    hoje_real = datetime.datetime.utcnow() - datetime.timedelta(hours=3)
    
    dias_validos_temp = [d for d in [st.session_state.get("d1_geral"), st.session_state.get("d2_geral"), st.session_state.get("d3_geral")] if d]
    data_padrao_calculada = hoje_real.date()
    if dias_validos_temp:
        ultimo_dia_str = dias_validos_temp[-1].strip()
        try:
            partes = ultimo_dia_str.split('/')
            if len(partes) == 2:
                dia_num, mes_num = int(partes[0]), int(partes[1])
                ano_num = int(ano) if ano and ano.isdigit() else hoje_real.year
                data_padrao_calculada = datetime.date(ano_num, mes_num, dia_num)
        except:
            pass

    usar_ultima_diligencia = st.checkbox("Usar automaticamente a data da última diligência", value=True, key="chk_usar_ultima_dil")

    if usar_ultima_diligencia:
        data_certidao = data_padrao_calculada
        st.info(f"📅 Data da certidão definida automaticamente pela última diligência: **{data_certidao.strftime('%d/%m/%Y')}**")
    else:
        data_certidao = st.date_input("Escolha a data da certidão:", value=data_padrao_calculada, format="DD/MM/YYYY", key="data_certidao_manual_escolha")
    
    st.write("**Informe os Dias e Horários**")
    
    cd1, cd2, cd3 = st.columns(3)

    with cd1:
        st.write("**Diligência 1**")
        d1 = st.text_input("Dia 1", placeholder="Ex: 08/08", key="d1_geral")
        h1 = st.text_input("Hora 1", placeholder="Ex: 14:55", key="h1_geral")
        
    with cd2:
        st.write("**Diligência 2**")
        d2 = st.text_input("Dia 2", placeholder="Ex: 11/08", key="d2_geral")
        h2 = st.text_input("Hora 2", placeholder="Ex: 16:58", key="h2_geral")
        
    with cd3:
        st.write("**Diligência 3**")
        d3 = st.text_input("Dia 3", placeholder="Ex: 12/08", key="d3_geral")
        h3 = st.text_input("Hora 3", placeholder="Ex: 11:15", key="h3_geral")

    st.divider()

    # ==========================================
    # OPÇÃO A: CERTIDÃO DETALHADA
    # ==========================================
    if tipo_certidao == "Certidão Negativa Detalhada":
        
        if st.session_state.get('limpar_detalhada_nova'):
            for k in list(st.session_state.keys()):
                if k.startswith(("mot_detn_", "rel_detn_", "ns_detn_", "inf_informante_det_", "cert_n_")) or k in ["nao_loc_dest_n", "nao_loc_bens_n"]:
                    st.session_state[k] = False 
                elif k in ["nome_inf_det_n", "sabe_tel_det_n", "sabe_end_det_n", "obs_livres_det_n"]:
                    st.session_state[k] = ""
            st.session_state['limpar_detalhada_nova'] = False

        st.write("**Deixei de cumprir a ordem descrita uma vez que:**")
        sit_c1, sit_c2 = st.columns(2)
        with sit_c1:
            nao_loc_dest_n = st.checkbox("O destinatário não foi localizado", key="nao_loc_dest_n")
        with sit_c2:
            nao_loc_bens_n = st.checkbox("Bem(ns) não localizado(s)", key="nao_loc_bens_n")

        motivos_selecionados = []
        with st.expander("📌 Selecionar Motivos Detalhados", expanded=False):
            motivos_list = [
                "local fechado", 
                "número não localizado", 
                "ap/bloco não localizado", 
                "local inabitado", 
                "rua/av não localizada", 
                "são insuficientes para saldar o débito",  
                "guarnecem a residência amparados pela Lei 8.009/90", 
                "sem condições psíquicas de entender conteúdo mandado"
            ]
            cols_mot = st.columns(3)
            for idx, m in enumerate(motivos_list):
                with cols_mot[idx % 3]:
                    if st.checkbox(m, key=f"mot_detn_{idx}"):
                        motivos_selecionados.append(m)

        st.markdown("---")
        relacoes_selecionadas = []
        nao_sabe_selecionados = []
        informacoes_informante_selecionadas = []
        sabe_tel = ""
        sabe_end = ""
        
        with st.expander("👤 Informações sobre o Informante", expanded=False):
            nome_inf_det_n = st.text_input("Nome do Sr(a):", placeholder="Vazio se não houver informante", key="nome_inf_det_n")
            st.caption("Relação / Qualidade:")
            relacoes_list = [
                "morador", "proprietário", "inquilino", "funcionário", "vizinho", "pai", "mãe",
                "padrasto", "madrasta", "filho", "irmão", "tio", "avô(ó)", "neto", "sobrinho",
                "primo", "transeunte", "viúvo", "ex", "esposo", "companheiro", "sogro", "enteado",
                "genro", "nora", "cunhado", "concunhado", "amigo"
            ]
            cols_rel = st.columns(4)
            for idx, r in enumerate(relacoes_list):
                with cols_rel[idx % 4]:
                    if st.checkbox(r, key=f"rel_detn_{idx}"):
                        relacoes_selecionadas.append(r)

            st.write("**Informações recebidas pelos informantes (Migradas):**")
            info_informantes_list = [
                "mudou-se",
                "não trabalha no local",
                "é desconhecida", 
                "não reside no local", 
                "dificilmente fica ali", 
                "antigo inquilino", 
                "rotatividade de inquilinos", 
                "trabalha em tempo integral", 
                "transferido", 
                "faliu", 
                "aparece esporadicamente", 
                "está viajando", 
                "antigo morador", 
                "repassado a terceiros", 
                "encontra-se preso", 
                "não exerce atividades", 
                "utiliza endereço para correspondências", 
                "antigo proprietário", 
                "internado", 
                "faleceu"
            ]
            cols_inf_info = st.columns(3)
            for idx, inf_item in enumerate(info_informantes_list):
                with cols_inf_info[idx % 3]:
                    if st.checkbox(inf_item, key=f"inf_informante_det_{idx}"):
                        informacoes_informante_selecionadas.append(inf_item)

            st.write("**Não sabendo o informante indicar:**")
            nao_sabe_list = [
                "endereço completo", "o paradeiro da mesma", "o dia/horário exato para encontrá-lo(a)", 
                "telefone", "dia/horário de retorno", "o presídio", 
                "dados do óbito", "previsão de alta", "o paradeiro do bem procurado"
            ]
            cols_ns = st.columns(3)
            for idx, ns in enumerate(nao_sabe_list):
                with cols_ns[idx % 3]:
                    if st.checkbox(ns, key=f"ns_detn_{idx}"):
                        nao_sabe_selecionados.append(ns)

            st.write("**Sabendo o informante indicar:**")
            sabe_tel = st.text_input("Telefone indicado:", key="sabe_tel_det_n")
            sabe_end = st.text_input("Endereço correto indicado:", key="sabe_end_det_n")
        
        with st.expander("📝 Certificações Adicionais", expanded=False):
            cert_extras = []
            c_extra1, c_extra2 = st.columns(2)

            with c_extra1:
                if st.checkbox("Procurei informações com moradores", key="cert_vizinhos_det_n"):
                    cert_extras.append("Procurei obter informações junto aos moradores/vizinhos locais e não obtive êxito")
                if st.checkbox("Cópia do mandado com informante", key="cert_copia_det_n"):
                    cert_extras.append("Devido à importância do mandado e da dificuldade de encontrar a pessoa procurada, deixei a cópia do mandado com o(a) senhor(a) acima mencionado(a)")
                if st.checkbox("Local Perigoso", key="cert_perigoso_det_n"):
                    cert_extras.append("Informo que o local é conhecidamente de grande periculosidade, o que quase sempre inviabiliza a obtenção de informações, pois os moradores ficam receosos de envolvimento com o processo e suas consequências, onde conversei com alguns vizinhos, que não quiseram se identificar, e ninguém soube informar detalhes sobre o possível horário/local para encontrar a pessoa procurada")
                if st.checkbox("Medo do Processo", key="cert_medo_det_n"):
                    cert_extras.append("Informo que os moradores ficam receosos de envolvimento com o processo e suas consequências, onde conversei com alguns vizinhos, que não quiseram se identificar, e ninguém soube informar detalhes sobre o possível horário/local para encontrar a pessoa procurada")
            with c_extra2:
                if st.checkbox("Apenas bens domésticos", key="cert_moveis_det_n"):
                    cert_extras.append("Informo que o imóvel é residencial e contém apenas móveis e utensílios domésticos")
                if st.checkbox("Zona Rural", key="cert_rural_det_n"):
                    cert_extras.append("Informo que o local é uma zona rural com difícil acesso, localização difícil, numeração irregular com muitas casas sem números na porta, o que causa desconforto nos moradores em fornecer informações precisas sobre o local/horário para encontrar a pessoa procurada")
                if st.checkbox("Condomínio s/ Porteiro", key="cert_blocos_det_n"):
                    cert_extras.append("Informo que o local é um condomínio de edifícios com vários blocos de apartamentos em seu interior; possui portaria na entrada do condomínio, mas não existe nenhum porteiro no local em nenhum horário; possui um interfone na entrada que é o único meio de contato com os apartamentos dentro do condomínio, mas aparentemente esse interfone não está funcionando, pois toquei várias vezes e ninguém atendeu; procurei informações com moradores que estavam saindo do condomínio sobre o possível contato com a pessoa procurada, mas ninguém soube informar se o mesmo reside no condomínio dizendo “são muitos moradores e não conhecemos todo mundo”, afirmando não saber informar também o possível horário para encontrá-la")
                if st.checkbox("Chuva Forte", key="cert_chuva_det_n"):
                    cert_extras.append("Informo que a execução da diligência restou dificultada em virtude das adversas condições meteorológicas no momento do ato, caracterizadas por intensa precipitação pluviométrica. Ressalto que tal circunstância, além de elevar significativamente o ruído ambiental comprometendo a audibilidade do chamamento realizado no portão, bem como ocasiona o natural recolhimento dos moradores no interior da residência com janelas e portas cerradas, o que obstaculizou a percepção da minha presença e, consequentemente, impediu o efetivo atendimento")

            observacoes_det = st.text_area("Observações Livres:", key="obs_livres_det_n")

        st.divider()

        if st.button("Gerar Certidão", type="primary", use_container_width=True, key="btn_gerar_docx_det_n"):
            with st.spinner("Construindo certidão e preparando arquivo..."):
                
                ano_base = ano if (ano and ano.isdigit()) else str(datetime.datetime.utcnow().year)
                dias_formatados = [formatar_data_completa(d.strip(), ano_base) for d in [d1, d2, d3] if d.strip()]
                
                horas_cruas = [h for h in [h1, h2, h3] if h]
                horas_validas = []
                for h in horas_cruas:
                    h_limpo = h.strip()
                    if h_limpo and not h_limpo.lower().endswith(('h', 'hs')):
                        h_limpo += 'hs' 
                    horas_validas.append(h_limpo)

                texto_data_hora = ""
                if len(dias_formatados) == 1:
                    h_str = horas_validas[0] if len(horas_validas) > 0 else "___hs"
                    texto_data_hora = f"no dia {dias_formatados[0]}, por volta das {h_str},"
                elif len(dias_formatados) > 1:
                    if len(horas_validas) > 1:
                        str_horas = ", ".join(horas_validas[:-1]) + f" e {horas_validas[-1]}"
                    elif len(horas_validas) == 1:
                        str_horas = horas_validas[0]
                    else:
                        str_horas = "___hs"
                        
                    str_dias = ", ".join(dias_formatados[:-1]) + f" e {dias_formatados[-1]}"
                    
                    if len(horas_validas) > 1:
                        texto_data_hora = f"nos dias {str_dias}, por volta das {str_horas}, respectivamente,"
                    else:
                        texto_data_hora = f"nos dias {str_dias}, por volta das {str_horas},"

                txt_endereco = f"à {endereco}" if endereco else "ao endereço indicado"
                txt_pessoa = f" em face de {pessoa}" if pessoa else ""
                
                texto_data_hora_formatado = f"{texto_data_hora} " if texto_data_hora else ""
                paragrafo = f"Certifico que, em cumprimento ao presente mandado, dirigi-me {txt_endereco}, {texto_data_hora_formatado}ocasião em que deixei de cumprir a ordem descrita{txt_pessoa}, uma vez que "
                sits = []
                if nao_loc_dest_n: sits.append("o destinatário não foi localizado")
                if nao_loc_bens_n: sits.append("o(s) bem(ns) indicados não foi(ram) localizado(s)")
                paragrafo += " e ".join(sits) + ". " if sits else "não foi possível a sua realização. "

                if motivos_selecionados:
                    frases_motivos = []
                    
                    mapa_motivos = {
                        "local fechado": "o imóvel encontrava-se fechado",
                        "número não localizado": "o número do imóvel não foi localizado",
                        "ap/bloco não localizado": "o apartamento ou bloco indicado não foi localizado",
                        "local inabitado": "o local encontra-se inabitado",
                        "rua/av não localizada": "a rua ou avenida indicada não foi localizada",
                        "são insuficientes para saldar o débito": "os bens encontrados são insuficientes para saldar o débito",
                        "guarnecem a residência amparados pela Lei 8.009/90": "os bens que guarnecem a residência estão amparados pela Lei 8.009/90",
                        "sem condições psíquicas de entender conteúdo mandado": "a pessoa procurada encontra-se sem condições psíquicas de compreender o conteúdo do mandado"
                    }

                    for m in motivos_selecionados:
                        frases_motivos.append(mapa_motivos.get(m, f"a pessoa procurada {m}"))
                    
                    # Lógica para não repetir o sujeito "a pessoa procurada"
                    frases_motivos_limpas = []
                    suj_motivo_usado = False
                    for f in frases_motivos:
                        if "a pessoa procurada" in f:
                            if not suj_motivo_usado:
                                suj_motivo_usado = True
                            else:
                                f = f.replace("a pessoa procurada ", "")
                        frases_motivos_limpas.append(f)
                    
                    if len(frases_motivos_limpas) > 1:
                        meio = ", que ".join(frases_motivos_limpas[:-1])
                        texto_motivos = f"{meio}, e que {frases_motivos_limpas[-1]}"
                    elif len(frases_motivos_limpas) == 1:
                        texto_motivos = frases_motivos_limpas[0]
                    else:
                        texto_motivos = ""
                        
                    if texto_motivos:
                        paragrafo += f"Constatou-se na diligência que {texto_motivos}. "

                # ==========================================
                # BLOCO MELHORADO DO INFORMANTE - Frase Contínua e Integrada
                # ==========================================
                if nome_inf_det_n or relacoes_selecionadas or informacoes_informante_selecionadas or nao_sabe_selecionados or sabe_tel or sabe_end:
                    
                    # Trata as relações para ficarem com "e" no final (ex: vizinho e pai)
                    if relacoes_selecionadas:
                        if len(relacoes_selecionadas) > 1:
                            relacoes_str = ", ".join(relacoes_selecionadas[:-1]) + " e " + relacoes_selecionadas[-1]
                        else:
                            relacoes_str = relacoes_selecionadas[0]
                    else:
                        relacoes_str = ""

                    if nome_inf_det_n:
                        txt_informante = f"pelo(a) Sr(a). {nome_inf_det_n}"
                        txt_informante += f", na qualidade de {relacoes_str}," if relacoes_str else ","
                    else:
                        if relacoes_str: 
                            # Removido o segundo "no local" para evitar cacofonia
                            txt_informante = f"por uma pessoa que se apresentou como {relacoes_str}, mas preferiu não se identificar, "
                        else: 
                            txt_informante = "por pessoa que preferiu não se identificar, "
                            
                    paragrafo += f"Conforme informações prestadas no local {txt_informante}"
                    
                    partes_informacao = []

                    # 1. Informações afirmativas do informante
                    if informacoes_informante_selecionadas:
                        # Dicionário com sujeitos explícitos para garantir gramática ao usar a conjunção "que"
                        mapa_inf_info = {
                            "mudou-se": "a pessoa procurada mudou-se dali",
                            "não trabalha no local": "a pessoa procurada não trabalha no local",
                            "é desconhecida": "a pessoa procurada é desconhecida no local",
                            "não reside no local": "a pessoa procurada não reside no endereço",
                            "dificilmente fica ali": "a pessoa procurada dificilmente se encontra ali",
                            "antigo inquilino": "a pessoa procurada trata-se de um antigo inquilino",
                            "rotatividade de inquilinos": "há uma alta rotatividade de inquilinos no local",
                            "trabalha em tempo integral": "a pessoa procurada trabalha em tempo integral",
                            "transferido": "a pessoa procurada foi transferida de localidade",
                            "faliu": "a empresa encerrou suas atividades ou faliu",
                            "aparece esporadicamente": "a pessoa procurada aparece apenas esporadicamente no endereço",
                            "está viajando": "a pessoa procurada encontra-se viajando",
                            "antigo morador": "a pessoa procurada trata-se de um antigo morador",
                            "repassado a terceiros": "o imóvel ou negócio foi repassado a terceiros",
                            "encontra-se preso": "a pessoa procurada encontra-se presa",
                            "não exerce atividades": "a pessoa ou empresa não exerce mais atividades no local",
                            "utiliza endereço para correspondências": "o endereço é utilizado apenas para o recebimento de correspondências",
                            "antigo proprietário": "a pessoa procurada trata-se do antigo proprietário",
                            "internado": "a pessoa procurada encontra-se internada",
                            "faleceu": "a pessoa procurada já é falecida"
                        }
                        
                        frases_inf_info = [mapa_inf_info.get(inf, inf) for inf in informacoes_informante_selecionadas]
                        
                        # Lógica para não repetir o sujeito nas falas do informante
                        frases_inf_limpas = []
                        suj_inf_usado = False
                        for f in frases_inf_info:
                            if "a pessoa procurada" in f:
                                if not suj_inf_usado:
                                    suj_inf_usado = True
                                else:
                                    f = f.replace("a pessoa procurada ", "")
                            frases_inf_limpas.append(f)

                        if len(frases_inf_limpas) > 1:
                            texto_inf_info = ", que ".join(frases_inf_limpas[:-1]) + f", e que {frases_inf_limpas[-1]}"
                        else:
                            texto_inf_info = frases_inf_limpas[0]
                        
                        partes_informacao.append(f"o(a) mesmo(a) declarou que {texto_inf_info}")

                    # 2. Informações negativas (Não sabe)
                    if nao_sabe_selecionados:
                        if len(nao_sabe_selecionados) > 1: 
                            texto_ns = ", ".join(nao_sabe_selecionados[:-1]) + f" e nem {nao_sabe_selecionados[-1]}"
                        else: 
                            texto_ns = nao_sabe_selecionados[0]
                        
                        if partes_informacao:
                            partes_informacao.append(f"acrescentando não saber informar {texto_ns}")
                        else:
                            partes_informacao.append(f"a qual declarou não saber informar {texto_ns}")

                    # 3. Informações positivas adicionais (Telefone / Endereço)
                    if sabe_tel or sabe_end:
                        sabes_list = []
                        if sabe_tel: sabes_list.append(f"o telefone de contato {sabe_tel}")
                        if sabe_end: sabes_list.append(f"o endereço atual como sendo: {sabe_end}")
                        texto_sabe = " e ".join(sabes_list)
                        
                        if partes_informacao:
                            partes_informacao.append(f"contudo, soube indicar {texto_sabe}")
                        else:
                            partes_informacao.append(f"a qual soube indicar {texto_sabe}")
                    
                    # Unindo todas as partes do informante de forma contínua
                    if not partes_informacao:
                        paragrafo += "que nada mais declarou. "
                    else:
                        texto_final_informante = ", ".join(partes_informacao)
                        paragrafo += f"{texto_final_informante}. "

                if not paragrafo.endswith(". "):
                    paragrafo = paragrafo.rstrip() + ". "

                if cert_extras: paragrafo += f"{'; '.join(cert_extras)}. "
                if observacoes_det: paragrafo += f"{observacoes_det.strip()} "

                doc = Document(); style = doc.styles['Normal']; font = style.font; font.name = 'Times New Roman'; font.size = Pt(12)
                try:
                    cabecalho_bytes = supabase.storage.from_("imagens_sistema").download("cabecalho.png")
                    p_img_cabecalho = doc.add_paragraph(); p_img_cabecalho.alignment = WD_ALIGN_PARAGRAPH.CENTER; p_img_cabecalho.add_run().add_picture(BytesIO(cabecalho_bytes), width=Cm(16))
                except: pass
                if processo:
                    texto_processo = f"Processo: {processo}"
                    if ano: texto_processo += f".{ano}.8.13.{comarca}"
                    doc.add_paragraph(texto_processo)
                if mandado: doc.add_paragraph(f"Mandado nº: {mandado}")
                doc.add_paragraph(""); p_titulo = doc.add_paragraph(); run_titulo = p_titulo.add_run("CERTIDÃO"); run_titulo.bold = True; run_titulo.font.size = Pt(16); p_titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER; doc.add_paragraph("")
                doc.add_paragraph(paragrafo.strip()).alignment = WD_ALIGN_PARAGRAPH.JUSTIFY; doc.paragraphs[-1].paragraph_format.first_line_indent = Pt(35.4); doc.add_paragraph("")
                doc.add_paragraph("Devolvo o mandado para os devidos fins. É verdade. Dou fé.").alignment = WD_ALIGN_PARAGRAPH.CENTER
                
                meses = ["janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"]
                doc.add_paragraph(f"{cidade_certidao}/{estado_certidao}, {data_certidao.day} de {meses[data_certidao.month - 1]} de {data_certidao.year}.").alignment = WD_ALIGN_PARAGRAPH.CENTER; doc.add_paragraph("")
                
                try:
                    assinatura_bytes = supabase.storage.from_("assinaturas_usuarios").download(f"{usuario_atual}.png")
                    p_img_assinatura = doc.add_paragraph(); p_img_assinatura.alignment = WD_ALIGN_PARAGRAPH.CENTER; p_img_assinatura.add_run().add_picture(BytesIO(assinatura_bytes), width=Cm(5))
                except: pass 
                p_assinatura = doc.add_paragraph(); p_assinatura.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run_nome = p_assinatura.add_run(f"{dados_usuario['nome']}\n"); run_nome.bold = True; run_nome.font.size = Pt(8)
                run_cargo = p_assinatura.add_run(f"{dados_usuario['cargo']}\n"); run_cargo.font.size = Pt(8)
                run_matricula = p_assinatura.add_run(f"{dados_usuario['matricula']}"); run_matricula.font.size = Pt(8)
                
                buffer = BytesIO(); doc.save(buffer); buffer.seek(0)
                docx_bytes = buffer.getvalue()
                
                data_arquivo = hoje_real.strftime("%d-%m-%Y_%Hh%Mm%Ss")
                sufixo_unico = str(uuid.uuid4())[:6]
                nome_base = f"Certidao_Negativa_Detalhada_{processo}_{data_arquivo}_{sufixo_unico}"
                
                if formato_saida == "PDF (.pdf)":
                    arquivo_final_bytes = converter_docx_para_pdf(docx_bytes)
                    nome_final = nome_base + ".pdf"
                    mime_final = "application/pdf"
                    if not arquivo_final_bytes:
                        st.error("Erro na conversão PDF. Baixando DOCX.")
                        arquivo_final_bytes, nome_final, mime_final = docx_bytes, nome_base + ".docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                else:
                    arquivo_final_bytes, nome_final, mime_final = docx_bytes, nome_base + ".docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

                supabase.storage.from_("certidoes_usuarios").upload(file=arquivo_final_bytes, path=f"{usuario_atual}/{nome_final}", file_options={"content-type": mime_final})
                
                st.session_state['doc_pronto_bytes_a2'] = arquivo_final_bytes
                st.session_state['doc_pronto_nome_a2'] = nome_final
                st.session_state['doc_pronto_mime_a2'] = mime_final
                st.session_state['piscar_tela'] = True
                st.session_state['limpar_detalhada_nova'] = True
                st.rerun()

        if 'doc_pronto_bytes_a2' in st.session_state:
            if st.session_state.get('piscar_tela'):
                st.balloons(); st.toast("✅ Certidão gerada!", icon="🎉")
                st.session_state['piscar_tela'] = False
            st.success("✅ Certidão salva na Nuvem!")
            st.download_button("📥 Baixar Arquivo", data=st.session_state['doc_pronto_bytes_a2'], file_name=st.session_state['doc_pronto_nome_a2'], mime=st.session_state['doc_pronto_mime_a2'], type="primary", use_container_width=True)
            
    # ==========================================
    # OPÇÃO B: CERTIDÃO POSITIVA
    # ==========================================
    elif tipo_certidao == "Certidão Positiva":
        
        if st.session_state.get('limpar_positiva'):
            st.session_state["mod_recebimento_pos"] = "Aceitou e assinou"
            st.session_state["recurso_pos"] = "Não informado"
            st.session_state["ciencia_pos"] = "Não informado"
            st.session_state["reconhece_pos"] = "Não informado"
            st.session_state["interesse_pos"] = "Não informado"
            st.session_state["adv_pos_novo"] = "Não informado"
            st.session_state["tel_pos"] = ""
            st.session_state["email_pos"] = ""
            st.session_state["tipo_realizacao_pos"] = "Pessoa procurada"
            st.session_state["nome_rep_pos"] = ""
            st.session_state["enunciado_pos"] = False
            st.session_state["obs_pos"] = ""
            st.session_state['limpar_positiva'] = False
            
        st.subheader("Detalhes da Certidão Positiva")
        st.markdown("---")
        
        st.write("**1. Condições do Recebimento:**")
        mod_recebimento_pos = st.radio(
            "Selecione a opção de recebimento e assinatura:",
            [
                "Aceitou e exarou sua assinatura no mandado",
                "Aceitou, mas não exarou sua assinatura",
                "Aceitou a contrafé, contudo, deixei de colher a assinatura por medida de precaução sanitária",
                "Não aceitou a contrafé",
                "Não aceitou e não exarou sua assinatura"
            ],
            key="mod_recebimento_pos"
        )
        
        st.markdown("---")
        st.subheader("2. Declarações e Informações do(a) Sra(a).")
        
        c_p1, c_p2 = st.columns(2)
        with c_p1:
            recurso_pos = st.radio("Sobre recorrer:", ["Não informado", "Deseja recorrer", "Não deseja recorrer"], key="recurso_pos")
            ciencia_pos = st.radio("Ciência da ação:", ["Não informado", "Tem ciência da ação em curso", "Não tem ciência da ação em curso"], key="ciencia_pos")
            reconhece_pos = st.radio("Reconhecimento de assinatura:", ["Não informado", "Reconhece como suas as assinaturas", "Não reconhece as assinaturas"], key="reconhece_pos")
        with c_p2:
            interesse_pos = st.radio("Interesse no prosseguimento:", ["Não informado", "Tem interesse no prosseguimento", "Não tem interesse no prosseguimento"], key="interesse_pos")
            adv_pos_novo = st.radio("Condições financeiras (Advogado):", ["Não informado", "Tem condições financeiras (advogado constituído)", "Não tem condições (hipossuficiente / defensor)"], key="adv_pos_novo")

        st.markdown("---")
        c_t1, c_t2 = st.columns(2)
        with c_t1:
            tel_pos = st.text_input("Telefone de contato:", placeholder="Ex: (31) 99999-9999", key="tel_pos")
        with c_t2:
            email_pos = st.text_input("E-mail:", placeholder="Ex: nome@email.com", key="email_pos")

        st.markdown("---")
        st.subheader("3. Especificidades da Realização do Ato")
        tipo_realizacao_pos = st.radio(
            "Como o ato foi realizado?",
            [
                "Pessoa procurada",
                "Representante legal",
                "Enunciados 5 e 38 do Fonaje"
            ],
            key="tipo_realizacao_pos"
        )

        nome_rep_pos = ""
        enunciado_pos = False

        if tipo_realizacao_pos == "Representante legal":
            nome_rep_pos = st.text_input("Nome do Representante Legal:", placeholder="Ex: Maria da Silva", key="nome_rep_pos")
        elif tipo_realizacao_pos == "Enunciados 5 e 38 do Fonaje":
            enunciado_pos = True
            nome_rep_pos = st.text_input("Nome da pessoa / dados complementares:", placeholder="Ex: Nome da pessoa que recebeu", key="nome_rep_pos")

        obs_pos = st.text_area("Observações Adicionais (opcional)", height=90, key="obs_pos")
            
        st.divider()
        
        if st.button("Gerar Certidão", type="primary", use_container_width=True, key="btn_gerar_positiva"):
            with st.spinner("Gerando certidão positiva..."):
                
                verbo_ato = "citei/intimei/notifiquei"
                
                ano_base = ano if (ano and ano.isdigit()) else str(datetime.datetime.utcnow().year)
                dias_formatados = [formatar_data_completa(d.strip(), ano_base) for d in [d1, d2, d3] if d.strip()]
                
                horas_cruas = [h for h in [h1, h2, h3] if h]
                horas_formatadas = []
                
                for h in horas_cruas:
                    h_limpo = h.strip()
                    if ":" in h_limpo:
                        partes_h = h_limpo.split(":")
                        horas_formatadas.append(f"{partes_h[0]}h{partes_h[1]}min")
                    elif not h_limpo.lower().endswith(('h', 'min')):
                        horas_formatadas.append(f"{h_limpo}h00min")
                    else:
                        horas_formatadas.append(h_limpo)
                
                str_horarios_dias = ""
                if len(dias_formatados) == 1 and len(horas_formatadas) == 1:
                    str_horarios_dias = f"por volta das {horas_formatadas[0]}, do dia {dias_formatados[0]}"
                elif len(dias_formatados) > 1 and len(horas_formatadas) > 1:
                    pares = []
                    for i in range(min(len(horas_formatadas), len(dias_formatados))):
                        pares.append(f"por volta das {horas_formatadas[i]}, do(s) dia(s) {dias_formatados[i]}")
                    if len(pares) == 2:
                        str_horarios_dias = f"{pares[0]} e {pares[1]}"
                    else:
                        str_horarios_dias = ", ".join(pares[:-1]) + f" e {pares[-1]}"
                else:
                     h_f = horas_formatadas[0] if horas_formatadas else "___hs 00min"
                     d_f = dias_formatados[0] if dias_formatados else f"___/___/{ano_base}"
                     str_horarios_dias = f"por volta das {h_f}, do(s) dia(s) {d_f}"

                txt_endereco = f"ao endereço indicado" if not endereco else f"à {endereco}"
                
                paragrafo = f"Certifico que, em cumprimento ao presente mandado, desloquei-me {txt_endereco}, {str_horarios_dias}, onde, {verbo_ato} o destinatário para todos os termos e conteúdo do mandado referido, que li e lhe dei para ler, do que ficou bem ciente. Dei-lhe a contrafé, que "
                
                if mod_recebimento_pos == "Aceitou e exarou sua assinatura no mandado":
                    paragrafo += "aceitou, exarando no mandado sua nota de ciência. "
                elif mod_recebimento_pos == "Aceitou, mas não exarou sua assinatura":
                    paragrafo += "aceitou, não exarando, contudo, no mandado sua nota de ciência. "
                elif mod_recebimento_pos == "Aceitou a contrafé, contudo, deixei de colher a assinatura por medida de precaução sanitária":
                    paragrafo += "aceitou, deixando eu de colher a assinatura física como medida de precaução contra a propagação de doenças infectocontagiosas, diante das circunstâncias verificadas no local ou das condições apresentadas pela pessoa. "
                elif mod_recebimento_pos == "Não aceitou a contrafé":
                    paragrafo += "não aceitou, exarando no mandado sua nota de ciência. "
                else:
                    paragrafo += "não aceitou e não exarou no mandado sua nota de ciência. "

                infos_adicionais = []
                if recurso_pos == "Deseja recorrer": infos_adicionais.append("deseja recorrer")
                elif recurso_pos == "Não deseja recorrer": infos_adicionais.append("não deseja recorrer")
                
                if ciencia_pos == "Tem ciência da ação em curso": infos_adicionais.append("tem ciência do ajuizamento da ação em curso")
                elif ciencia_pos == "Não tem ciência da ação em curso": infos_adicionais.append("não tem ciência do ajuizamento da ação em curso")

                if reconhece_pos == "Reconhece como suas as assinaturas": infos_adicionais.append("reconhece como suas as assinaturas constantes nos documentos juntados no processo")
                elif reconhece_pos == "Não reconhece as assinaturas": infos_adicionais.append("não reconhece como suas as assinaturas constantes nos documentos juntados no processo")

                if interesse_pos == "Tem interesse no prosseguimento": infos_adicionais.append("tem interesse no prosseguimento do feito/demanda")
                elif interesse_pos == "Não tem interesse no prosseguimento": infos_adicionais.append("não tem interesse no prosseguimento do feito/demanda")

                if adv_pos_novo == "Tem condições financeiras (advogado constituído)": 
                    infos_adicionais.append("tem condições financeiras de apresentar sua defesa através de advogado constituído")
                elif adv_pos_novo == "Não tem condições (hipossuficiente / defensor)": 
                    infos_adicionais.append("não tem condições financeiras de apresentar sua defesa através de advogado constituído, declarando sua hipossuficiência e requerendo a nomeação de um defensor público ou dativo para fazê-la")

                contatos = []
                if tel_pos: contatos.append(f"telefone de contato ({tel_pos})")
                if email_pos: contatos.append(f"e-mail ({email_pos})")

                # Ajuste no Bloco da Positiva para fundir as informações em uma única frase fluida
                if infos_adicionais:
                    txt_infos = ", ".join(infos_adicionais[:-1]) + " e " + infos_adicionais[-1] if len(infos_adicionais) > 1 else infos_adicionais[0]
                    paragrafo += f"Certifico, ainda, que o(a) supracitado(a) informou que {txt_infos}"
                    
                    if contatos:
                        paragrafo += f", bem como indicou seu {' e '.join(contatos)}. "
                    else:
                        paragrafo += ". "
                else:
                    if contatos:
                        paragrafo += f"Certifico, ainda, que o(a) supracitado(a) indicou seu {' e '.join(contatos)}. "
                
                if tipo_realizacao_pos == "Representante legal":
                    rep_txt = nome_rep_pos if nome_rep_pos else "quem de direito"
                    # Ajusta a concordância dependendo se o nome foi digitado ou não
                    txt_alvo = f"do(a) Sr(a). {pessoa}" if pessoa else "da pessoa referida no mandado"
                    paragrafo += f"Certifico também que o ato foi realizado na pessoa {txt_alvo}, que se apresentou como representante legal ({rep_txt}). "
                elif tipo_realizacao_pos == "Enunciados 5 e 38 do Fonaje":
                    enq_txt = nome_rep_pos if nome_rep_pos else "quem de direito"
                    paragrafo += f"Certifico também que o ato foi realizado na pessoa do(a) Sr(a). {enq_txt}, de acordo com os Enunciados 5 e 38 do Fórum Permanente de Juízes Coordenadores dos Juizados Especiais. "

                if obs_pos:
                    paragrafo += f"{obs_pos.strip()} "

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
                
                doc.add_paragraph("")
                p_titulo = doc.add_paragraph()
                run_titulo = p_titulo.add_run("CERTIDÃO POSITIVA")
                run_titulo.bold = True
                run_titulo.font.size = Pt(16)
                p_titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
                doc.add_paragraph("")

                p_Linha = doc.add_paragraph(paragrafo.strip())
                p_Linha.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                p_Linha.paragraph_format.first_line_indent = Pt(35.4) 

                doc.add_paragraph("")
                doc.add_paragraph("Devolvo o mandado para os devidos fins. É verdade. Dou fé.").alignment = WD_ALIGN_PARAGRAPH.CENTER
                
                meses = ["janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"]
                doc.add_paragraph(f"{cidade_certidao}/{estado_certidao}, {data_certidao.day} de {meses[data_certidao.month - 1]} de {data_certidao.year}.").alignment = WD_ALIGN_PARAGRAPH.CENTER
                doc.add_paragraph("")
                
                try:
                    assinatura_bytes = supabase.storage.from_("assinaturas_usuarios").download(f"{usuario_atual}.png")
                    p_img_assinatura = doc.add_paragraph(); p_img_assinatura.alignment = WD_ALIGN_PARAGRAPH.CENTER; p_img_assinatura.add_run().add_picture(BytesIO(assinatura_bytes), width=Cm(6))
                except: pass 
                
                p_assinatura = doc.add_paragraph(); p_assinatura.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run_nome = p_assinatura.add_run(f"{dados_usuario['nome']}\n"); run_nome.bold = True; run_nome.font.size = Pt(8)
                run_cargo = p_assinatura.add_run(f"{dados_usuario['cargo']}\n"); run_cargo.font.size = Pt(8)
                run_matricula = p_assinatura.add_run(f"{dados_usuario['matricula']}"); run_matricula.font.size = Pt(8)
                
                buffer = BytesIO(); doc.save(buffer); buffer.seek(0)
                docx_bytes = buffer.getvalue()
                
                data_arquivo = hoje_real.strftime("%d-%m-%Y_%Hh%M")
                nome_base = f"Certidao_Positiva_{processo}_{data_arquivo}"
                
                if formato_saida == "PDF (.pdf)":
                    arquivo_final_bytes = converter_docx_para_pdf(docx_bytes)
                    nome_final = nome_base + ".pdf"
                    mime_final = "application/pdf"
                    if not arquivo_final_bytes:
                        st.error("Erro na conversão PDF. Baixando DOCX.")
                        arquivo_final_bytes, nome_final, mime_final = docx_bytes, nome_base + ".docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                else:
                    arquivo_final_bytes, nome_final, mime_final = docx_bytes, nome_base + ".docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

                supabase.storage.from_("certidoes_usuarios").upload(file=arquivo_final_bytes, path=f"{usuario_atual}/{nome_final}", file_options={"content-type": mime_final})
                
                st.session_state['doc_pronto_bytes_c'] = arquivo_final_bytes
                st.session_state['doc_pronto_nome_c'] = nome_final
                st.session_state['doc_pronto_mime_c'] = mime_final
                st.session_state['piscar_tela'] = True
                st.session_state['limpar_positiva'] = True
                st.rerun()

        if 'doc_pronto_bytes_c' in st.session_state:
            if st.session_state.get('piscar_tela'):
                st.balloons(); st.toast("✅ Certidão Positiva gerada!", icon="🎉")
                st.session_state['piscar_tela'] = False
            st.success("✅ Certidão salva na Nuvem!")
            st.download_button("📥 Baixar Arquivo", data=st.session_state['doc_pronto_bytes_c'], file_name=st.session_state['doc_pronto_nome_c'], mime=st.session_state['doc_pronto_mime_c'], type="primary", use_container_width=True)

    # ==========================================
    # OPÇÃO C: CERTIDÃO POSITIVA POR HORA CERTA
    # ==========================================
    elif tipo_certidao == "Certidão Positiva por Hora Certa":
        
        if st.session_state.get('limpar_horacerta'):
            st.session_state["fin_hc"] = "Citação"
            for k in ["hc_nome_terceiro", "hc_relacao", "hc_data_retorno", "hc_hora_retorno"]:
                st.session_state[k] = ""
            st.session_state["hc_encontrou_alvo"] = "Não"
            st.session_state["hc_aceitou"] = "Sim"
            st.session_state["hc_assinou"] = "Não"
            st.session_state['limpar_horacerta'] = False
            
        st.subheader("1. Suspeita de Ocultação e Agendamento")
        finalidade_hc = st.selectbox("Ato sendo praticado:", ["Citação", "Intimação", "Notificação"], key="fin_hc")
        
        c_hc1, c_hc2 = st.columns(2)
        with c_hc1:
            hc_nome_terceiro = st.text_input("Com quem marcou a hora certa?", placeholder="Ex: Sra. Teresinha dos Santos", key="hc_nome_terceiro")
            hc_data_retorno = st.text_input("Data marcada para o retorno:", placeholder="Ex: 19/05", key="hc_data_retorno")
        with c_hc2:
            hc_relacao = st.text_input("Qualidade/Relação:", placeholder="Ex: mãe, vizinha, porteiro", key="hc_relacao")
            hc_hora_retorno = st.text_input("Hora marcada para o retorno:", placeholder="Ex: 18:15", key="hc_hora_retorno")

        st.markdown("---")
        st.subheader("2. Desfecho do Retorno")
        c_hc3, c_hc4, c_hc5 = st.columns(3)
        with c_hc3: hc_encontrou_alvo = st.radio("Encontrou a pessoa?", ["Não", "Sim"], key="hc_encontrou_alvo")
        with c_hc4: hc_aceitou = st.radio("O terceiro/alvo aceitou receber?", ["Sim", "Não"], key="hc_aceitou")
        with c_hc5: hc_assinou = st.radio("O terceiro/alvo assinou?", ["Não", "Sim", "Covid-19"], key="hc_assinou")

        st.divider()

        if st.button("Gerar Certidão", type="primary", use_container_width=True, key="btn_gerar_horacerta"):
            with st.spinner("Construindo certidão de Hora Certa..."):
                
                txt_endereco = f"à {endereco}" if endereco else "ao endereço indicado"
                txt_pessoa = f" de {pessoa}" if pessoa else " da pessoa referida no mandado"
                
                ano_base = ano if (ano and ano.isdigit()) else str(datetime.datetime.utcnow().year)
                dias_formatados = [formatar_data_completa(d.strip(), ano_base) for d in [d1, d2, d3] if d.strip()]
                
                hc_data_retorno_formatada = formatar_data_completa(hc_data_retorno.strip(), ano_base) if hc_data_retorno else ""
                
                horas_cruas = [h for h in [h1, h2, h3] if h]
                horas_validas = []
                for h in horas_cruas:
                    h_limpo = h.strip()
                    if h_limpo and not h_limpo.lower().endswith(('h', 'hs', 'min')): h_limpo += 'hs'
                    horas_validas.append(h_limpo)
                
                texto_data_hora = ""
                if len(dias_formatados) == 1:
                    h_str = horas_validas[0] if len(horas_validas) > 0 else "___hs"
                    texto_data_hora = f"no dia {dias_formatados[0]}, por volta das {h_str},"
                elif len(dias_formatados) > 1:
                    if len(horas_validas) > 1:
                        str_horas = ", ".join(horas_validas[:-1]) + f" e {horas_validas[-1]}"
                    elif len(horas_validas) == 1:
                        str_horas = horas_validas[0]
                    else:
                        str_horas = "___hs"
                        
                    str_dias = ", ".join(dias_formatados[:-1]) + f" e {dias_formatados[-1]}"
                    
                    if len(horas_validas) > 1:
                        texto_data_hora = f"nos dias {str_dias}, por volta das {str_horas}, respectivamente,"
                    else:
                        texto_data_hora = f"nos dias {str_dias}, por volta das {str_horas},"
                
                hr_limpo = hc_hora_retorno.strip()
                if hr_limpo and not hr_limpo.lower().endswith(('h', 'hs', 'min')): hr_limpo += 'hs'

                nome_ato = finalidade_hc.upper()

                txt_terceiro = f"na pessoa do(a) Sr(a). {hc_nome_terceiro}" if hc_nome_terceiro else "na pessoa de um terceiro ali presente"
                txt_relacao = f", na qualidade de {hc_relacao}," if hc_relacao else ","
                txt_retorno_alvo = "ali não a encontrando" if hc_encontrou_alvo == "Não" else "ali a encontrando"
                
                if hc_aceitou == "Sim":
                    if hc_assinou == "Sim": txt_final = "a qual aceitou o documento e exarou sua assinatura no respectivo mandado."
                    elif hc_assinou == "Não": txt_final = "a qual aceitou o documento, mas recusou-se a exarar sua assinatura no respectivo mandado."
                    else: txt_final = "a qual aceitou o documento, deixando eu de colher a assinatura física como medida de prevenção sanitária/Covid-19."
                else: txt_final = "a qual se recusou a receber a contrafé e a assinar o respectivo mandado."

                paragrafo = f"Certifico que, em cumprimento ao presente mandado, dirigi-me {txt_endereco}, onde {texto_data_hora} não encontrei a pessoa procurada. Diante das diligências frustradas e havendo fundada suspeita de ocultação, efetuei o agendamento de HORA CERTA {txt_terceiro}{txt_relacao} intimando-o(a) de que retornaria no dia {hc_data_retorno_formatada}, pontualmente às {hr_limpo}, para efetivar o ato judicial. Retornando no dia e hora estritamente designados, {txt_retorno_alvo}, dei por realizada a {nome_ato}{txt_pessoa}, deixando a respectiva contrafé com a pessoa mencionada, {txt_final}"
                
                if not paragrafo.endswith(". "):
                    paragrafo = paragrafo.rstrip() + ". "

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
                doc.add_paragraph(""); p_titulo = doc.add_paragraph(); run_titulo = p_titulo.add_run("CERTIDÃO POSITIVA POR HORA CERTA"); run_titulo.bold = True; run_titulo.font.size = Pt(16); p_titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER; doc.add_paragraph("")
                doc.add_paragraph(paragrafo.strip()).alignment = WD_ALIGN_PARAGRAPH.JUSTIFY; doc.paragraphs[-1].paragraph_format.first_line_indent = Pt(35.4); doc.add_paragraph("")
                doc.add_paragraph("Devolvo o mandado para os devidos fins. É verdade. Dou fé.").alignment = WD_ALIGN_PARAGRAPH.CENTER
                
                meses = ["janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"]
                doc.add_paragraph(f"{cidade_certidao}/{estado_certidao}, {data_certidao.day} de {meses[data_certidao.month - 1]} de {data_certidao.year}.").alignment = WD_ALIGN_PARAGRAPH.CENTER; doc.add_paragraph("")
                try:
                    assinatura_bytes = supabase.storage.from_("assinaturas_usuarios").download(f"{usuario_atual}.png")
                    p_img_assinatura = doc.add_paragraph(); p_img_assinatura.alignment = WD_ALIGN_PARAGRAPH.CENTER; p_img_assinatura.add_run().add_picture(BytesIO(assinatura_bytes), width=Cm(6))
                except: pass 
                p_assinatura = doc.add_paragraph(); p_assinatura.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run_nome = p_assinatura.add_run(f"{dados_usuario['nome']}\n"); run_nome.bold = True; run_nome.font.size = Pt(8)
                run_cargo = p_assinatura.add_run(f"{dados_usuario['cargo']}\n"); run_cargo.font.size = Pt(8)
                run_matricula = p_assinatura.add_run(f"{dados_usuario['matricula']}"); run_matricula.font.size = Pt(8)
                
                buffer = BytesIO(); doc.save(buffer); buffer.seek(0)
                docx_bytes = buffer.getvalue()
                
                data_arquivo = hoje_real.strftime("%d-%m-%Y_%Hh%M")
                nome_base = f"Certidao_HoraCerta_{processo}_{data_arquivo}"
                
                if formato_saida == "PDF (.pdf)":
                    arquivo_final_bytes = converter_docx_para_pdf(docx_bytes)
                    nome_final = nome_base + ".pdf"
                    mime_final = "application/pdf"
                    if not arquivo_final_bytes:
                        st.error("Erro na conversão PDF. Baixando DOCX.")
                        arquivo_final_bytes, nome_final, mime_final = docx_bytes, nome_base + ".docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                else:
                    arquivo_final_bytes, nome_final, mime_final = docx_bytes, nome_base + ".docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

                supabase.storage.from_("certidoes_usuarios").upload(file=arquivo_final_bytes, path=f"{usuario_atual}/{nome_final}", file_options={"content-type": mime_final})
                
                st.session_state['doc_pronto_bytes_d'] = arquivo_final_bytes
                st.session_state['doc_pronto_nome_d'] = nome_final
                st.session_state['doc_pronto_mime_d'] = mime_final
                st.session_state['piscar_tela'] = True
                st.session_state['limpar_horacerta'] = True 
                st.rerun() 

        if 'doc_pronto_bytes_d' in st.session_state:
            if st.session_state.get('piscar_tela'):
                st.balloons(); st.toast("✅ Certidão de Hora Certa gerada!", icon="🎉")
                st.session_state['piscar_tela'] = False 
            st.success("✅ Certidão salva na Nuvem!")
            st.download_button("📥 Baixar Arquivo", data=st.session_state['doc_pronto_bytes_d'], file_name=st.session_state['doc_pronto_nome_d'], mime=st.session_state['doc_pronto_mime_d'], type="primary", use_container_width=True)