import streamlit as st
import os
import hashlib
import zipfile
from io import BytesIO
import datetime
import uuid  # Para gerar as chaves únicas
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
    /* Oculta marcações padrão do Streamlit */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    
    .block-container { padding-top: 1rem; padding-bottom: 1rem; }
    h1 { font-size: 22px; text-align: center; margin-bottom: 0; padding-bottom: 0;}
    .stCheckbox { margin-top: -5px; margin-bottom: -5px; }
    div[role="radiogroup"] { margin-top: -10px; }
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

# === FUNÇÕES DE PERSISTÊNCIA DE LOGIN POR URL ===

def criar_sessao_segura(usuario):
    """Gera uma chave única, salva no banco e retorna"""
    chave = str(uuid.uuid4()) # Gera um ID único aleatório (ex: 550e8400-e29b...)
    # Define expiração para 30 dias a partir de agora (UTC)
    expira = datetime.datetime.utcnow() + datetime.timedelta(days=30)
    
    # Salva na tabela banco_sessoes
    supabase.table("banco_sessoes").insert({
        "usuario": usuario,
        "chave_acesso": chave,
        "expira_em": expira.isoformat()
    }).execute()
    
    return chave

def verificar_sessao_url():
    """Lê a chave da URL e tenta logar o usuário"""
    # Pega os parâmetros da URL (ex: ?access_key=...)
    params = st.query_params
    
    if "access_key" in params:
        chave_url = params["access_key"]
        
        # Busca a chave no banco de dados
        agora = datetime.datetime.utcnow().isoformat()
        resposta = supabase.table("banco_sessoes").select("usuario").eq("chave_acesso", chave_url).gt("expira_em", agora).execute()
        
        if len(resposta.data) > 0:
            # Chave válida e não expirada! Loga o usuário na sessão do Streamlit
            st.session_state["usuario_logado"] = resposta.data[0]["usuario"]
            return True
    return False

def encerrar_sessao():
    """Remove a chave da URL e limpa o login"""
    params = st.query_params
    if "access_key" in params:
        chave_url = params["access_key"]
        # Deleta do banco de dados (opcional, mas boa prática)
        supabase.table("banco_sessoes").delete().eq("chave_acesso", chave_url).execute()
    
    # Limpa o parâmetro da URL no navegador
    st.query_params.clear()
    # Limpa a sessão do Streamlit
    st.session_state["usuario_logado"] = None

# ==========================================
# 2. CONTROLE DE SESSÃO E LOGIN
# ==========================================
if "usuario_logado" not in st.session_state:
    st.session_state["usuario_logado"] = None

# Tenta recuperar o login pela URL antes de mostrar a tela de login
if st.session_state["usuario_logado"] is None:
    if verificar_sessao_url():
        st.rerun() # Recarrega a página já logado

# Se ainda não estiver logado, mostra a tela de login
if st.session_state["usuario_logado"] is None:
    st.title("⚖️ Sistema de Certidões - TJMG")
    
    aba_login, aba_cadastro = st.tabs(["Entrar", "Criar Nova Conta"])
    
    with aba_login:
        st.write("Acesse sua conta para gerar certidões.")
        usuario_login = st.text_input("Usuário:", key="log_usr").lower().strip()
        senha_login = st.text_input("Senha:", type="password", key="log_pwd")
        
        # Checkbox opcional para o usuário decidir se quer manter logado
        manter_logado = st.checkbox("Manter-me logado neste dispositivo (30 dias)", value=True)
        
        if st.button("Entrar", type="primary", use_container_width=True):
            if usuario_login and senha_login:
                with st.spinner("Autenticando..."):
                    resposta = supabase.table("banco_usuarios").select("*").eq("usuario", usuario_login).execute()
                    
                    if len(resposta.data) > 0:
                        dados_bd = resposta.data[0]
                        senha_criptografada = gerar_hash_senha(senha_login)
                        if dados_bd["senha"] == senha_criptografada:
                            # Login com senha correto!
                            st.session_state["usuario_logado"] = usuario_login
                            
                            if manter_logado:
                                # Gera a chave, salva no banco e injeta na URL
                                nova_chave = criar_sessao_segura(usuario_login)
                                st.query_params["access_key"] = nova_chave
                            
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
                with st.spinner("Criando conta..."):
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
                
    st.stop() # Interrompe o script aqui se não estiver logado

# ==========================================
# 3. DADOS DO USUÁRIO E MENU LATERAL (SÓ ENTRA AQUI SE LOGADO)
# ==========================================
usuario_atual = st.session_state["usuario_logado"]

# Busca dados do usuário (garantindo que ele ainda existe)
resposta_usuario = supabase.table("banco_usuarios").select("*").eq("usuario", usuario_atual).execute()
if len(resposta_usuario.data) == 0:
    # Usuário foi deletado do banco, força logout
    encerrar_sessao()
    st.rerun()

dados_usuario = resposta_usuario.data[0]

with st.sidebar:
    st.write(f"👤 Olá, **{usuario_atual.title()}**!")
    st.divider()
    
    opcoes_menu = ["📝 Gerar Certidão", "📂 Minhas Certidões", "⚙️ Meu Perfil"]
    
    # Adiciona o menu de administrador se o usuário for '10228429'
    if usuario_atual == "10228429":
        opcoes_menu.append("🛡️ Painel do Administrador")
        
    menu = st.radio("Navegação:", opcoes_menu)
    st.divider()
    
    # Botão de Sair atualizado para limpar a URL
    if st.button("Sair (Logout)", use_container_width=True):
        with st.spinner("Saindo..."):
            encerrar_sessao()
            st.rerun()

# ==========================================
# 4. TELA: MEU PERFIL
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
        with st.spinner("Salvando..."):
            supabase.table("banco_usuarios").update({
                "nome": novo_nome,
                "cargo": novo_cargo,
                "matricula": nova_matricula
            }).eq("usuario", usuario_atual).execute()
            
            if arquivo_assinatura is not None:
                try:
                    # Tenta remover a antiga (pode falhar se não existir)
                    supabase.storage.from_("assinaturas_usuarios").remove([f"{usuario_atual}.png"])
                except:
                    pass
                # Upload da nova assinatura
                supabase.storage.from_("assinaturas_usuarios").upload(
                    file=arquivo_assinatura.getvalue(),
                    path=f"{usuario_atual}.png",
                    file_options={"content-type": arquivo_assinatura.type}
                )
                    
            st.success("Perfil atualizado e salvo na nuvem com sucesso!")
            st.rerun() # Recarrega para atualizar os dados na sidebar

# ==========================================
# 5. TELA: MINHAS CERTIDÕES
# ==========================================
elif menu == "📂 Minhas Certidões":
    st.title("📂 Minhas Certidões Salvas")
    st.write("Baixe ou exclua seus arquivos salvos na nuvem.")
    
    with st.spinner("Carregando lista de arquivos..."):
        try:
            arquivos_nuvem = supabase.storage.from_("certidoes_usuarios").list(usuario_atual)
        except:
            arquivos_nuvem = []
    
    arquivos = [arq for arq in arquivos_nuvem if arq["name"] != ".emptyFolder" and arq["name"] != ""]
    
    if not arquivos:
        st.info("Nenhuma certidão salva ainda.")
    else:
        # Ordena por data de criação (mais recente primeiro)
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
                # Pega a data do banco e ajusta o formato (UTC -> Brasília -3h)
                data_str = item["created_at"].replace("Z", "+00:00")
                data_obj = datetime.datetime.fromisoformat(data_str)
                
                # Remove a "etiqueta" do fuso do servidor e subtrai 3 horas cravadas
                data_br_obj = data_obj.replace(tzinfo=None) - datetime.timedelta(hours=3)
                data_br = data_br_obj.strftime("%d/%m/%Y às %H:%M")
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
                                try:
                                    arquivo_bytes = supabase.storage.from_("certidoes_usuarios").download(f"{usuario_atual}/{arq}")
                                    zip_file.writestr(arq, arquivo_bytes)
                                except:
                                    st.error(f"Falha ao baixar o arquivo: {arq}")
                                
                        st.download_button(
                            label="✔️ Clique aqui para baixar o ZIP",
                            data=zip_buffer.getvalue(),
                            file_name=f"certidoes_{usuario_atual}.zip",
                            mime="application/zip",
                            use_container_width=True
                        )
            
            with c_btn2:
                if st.button("🗑️ Excluir Selecionadas", use_container_width=True):
                    with st.spinner("Excluindo arquivos..."):
                        caminhos_para_excluir = [f"{usuario_atual}/{arq}" for arq in arquivos_selecionados]
                        supabase.storage.from_("certidoes_usuarios").remove(caminhos_para_excluir)
                        st.success("Arquivos excluídos da nuvem com sucesso!")
                        st.rerun()

# ==========================================
# 6. TELA: PAINEL DO ADMINISTRADOR
# ==========================================
elif menu == "🛡️ Painel do Administrador":
    if usuario_atual != "10228429":
        st.error("Acesso restrito apenas ao Administrador.")
        st.stop()
        
    st.title("🛡️ Painel de Administração")
    st.write("Área restrita para gestão de oficiais e auditoria de certidões em nuvem.")
    
    aba_adm1, aba_adm2 = st.tabs(["👥 Gerenciar Usuários", "📊 Auditoria de Certidões Gerais"])
    
    # ABA 1: GERENCIAR USUÁRIOS
    with aba_adm1:
        st.subheader("Oficiais Cadastrados no Sistema")
        with st.spinner("Buscando usuários..."):
            res_todos = supabase.table("banco_usuarios").select("usuario, nome, cargo, matricula").execute()
            usuarios_cadastrados = res_todos.data
        
        if usuarios_cadastrados:
            for u in usuarios_cadastrados:
                with st.expander(f"👤 Usuário: {u['usuario']} — Nome: {u.get('nome') or 'Não preenchido'}"):
                    st.write(f"**Cargo:** {u.get('cargo')}")
                    st.write(f"**Matrícula:** {u.get('matricula')}")
                    
                    if u['usuario'] != usuario_atual:
                        if st.button(f"🗑️ Excluir usuário {u['usuario']}", key=f"del_usr_{u['usuario']}"):
                            with st.spinner("Excluindo..."):
                                supabase.table("banco_usuarios").delete().eq("usuario", u['usuario']).execute()
                                # Limpa as sessões ativas desse usuário para deslogá-lo imediatamente
                                supabase.table("banco_sessoes").delete().eq("usuario", u['usuario']).execute()
                                st.success(f"Usuário {u['usuario']} removido com sucesso!")
                                st.rerun()
                    else:
                        st.caption("*(Esta é a sua conta de Administrador principal)*")
        else:
            st.info("Nenhum usuário encontrado.")

    # ABA 2: AUDITORIA DE CERTIDÕES
    with aba_adm2:
        st.subheader("Certidões Geradas por Todos os Oficiais")
        st.write("Inspecione, baixe ou exclua os arquivos salvos por qualquer oficial.")
        
        with st.spinner("Buscando pastas de oficiais..."):
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
                        
                    certioes_validas = [f for f in arquivos_oficial if f["name"] != ".emptyFolder" and f["name"] != ""]
                    
                    if not certioes_validas:
                        st.caption("Nenhuma certidão gerada por este oficial ainda.")
                    else:
                        for arq in certioes_validas:
                            # Colunas ajustadas para acomodar Nome, Download e Excluir
                            c_arq_nome, c_btn_dl, c_btn_del = st.columns([4, 2, 2])
                            
                            with c_arq_nome:
                                st.text(arq["name"])
                                
                            with c_btn_dl:
                                if st.button("📥 Baixar", key=f"dl_adm_{nome_oficial}_{arq['name']}", use_container_width=True):
                                    with st.spinner("Baixando..."):
                                        try:
                                            file_bytes = supabase.storage.from_("certidoes_usuarios").download(f"{nome_oficial}/{arq['name']}")
                                            st.download_button(
                                                label="Confirmar",
                                                data=file_bytes,
                                                file_name=arq["name"],
                                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                                key=f"btn_dl_real_{nome_oficial}_{arq['name']}"
                                            )
                                        except:
                                            st.error("Falha ao baixar.")
                                    
                            with c_btn_del:
                                if st.button("🗑️ Excluir", key=f"del_adm_{nome_oficial}_{arq['name']}", use_container_width=True):
                                    with st.spinner("Excluindo..."):
                                        try:
                                            supabase.storage.from_("certidoes_usuarios").remove([f"{nome_oficial}/{arq['name']}"])
                                            st.success("Excluído!")
                                            st.rerun()
                                        except:
                                            st.error("Falha ao excluir.")
                    st.divider()

# ==========================================
# 7. TELA: GERADOR DE CERTIDÃO
# ==========================================
elif menu == "📝 Gerar Certidão":
    st.title("Gerador de Certidão Negativa")
    
    if not dados_usuario.get("nome"):
        st.warning("⚠️ Você ainda não configurou seu perfil! Vá em 'Meu Perfil' no menu lateral e preencha seus dados antes de gerar certidões.")
        st.stop()

    tipo_certidao = st.selectbox(
        "Selecione o Modelo de Certidão:", 
        ["Certidão Negativa Detalhada", "Certidão Negativa Simples (Opções Rápidas)"]
    )
    
    st.divider()

    # --- CAMPOS COMPARTILHADOS (Cabeçalho e Datas) ---
    c_mandado, c_proc = st.columns([1, 3])
    with c_mandado:
        mandado = st.text_input("Mandado:", placeholder="Ex: 01")
    with c_proc:
        processo = st.text_input("Informe o Processo:", placeholder="Ex: 4400281-16")
    
    c_ano, c_comarca = st.columns(2)
    with c_ano:
        ano = st.text_input("Ano:", placeholder="Ex: 2026")
    with c_comarca:
        # NOVO CAMPO: Comarca com padrão 0245
        comarca = st.text_input("Código Comarca:", value="0245", placeholder="Ex: 0245")

    c_end, c_pes = st.columns(2)
    with c_end:
        endereco = st.text_input("Endereço (opcional):", placeholder="Se vazio: 'informado no mesmo'")
    with c_pes:
        pessoa = st.text_input("Pessoa procurada:", placeholder="Deixe vazio para termo genérico")

    st.markdown("---")
    st.write("**Dias e Horários das Diligências:**")
    
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

    # ==========================================
    # OPÇÃO A: CERTIDÃO DETALHADA
    # ==========================================
    if tipo_certidao == "Certidão Negativa Detalhada":
        st.write("**Deixei de cumprir o ato uma vez que:**")
        sit_c1, sit_c2 = st.columns(2)
        with sit_c1:
            nao_loc_dest = st.checkbox("O destinatário do mandado não foi localizado")
        with sit_c2:
            nao_loc_bens = st.checkbox("O(s) bem(ns) indicados não foi(ram) localizado(s)")

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

        if st.button("Salvar na Nuvem / Gerar DOCX (Detalhada)", type="primary", use_container_width=True):
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

                # Geração do DOCX
                doc = Document()
                style = doc.styles['Normal']
                font = style.font
                font.name = 'Times New Roman'
                font.size = Pt(12)

                # Busca o cabeçalho no Storage do Supabase (para ser dinâmico)
                try:
                    cabecalho_bytes = supabase.storage.from_("imagens_sistema").download("cabecalho.png")
                    cabecalho_stream = BytesIO(cabecalho_bytes)
                    p_img_cabecalho = doc.add_paragraph()
                    p_img_cabecalho.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    run_img_cab = p_img_cabecalho.add_run()
                    run_img_cab.add_picture(cabecalho_stream, width=Cm(16))
                except:
                    # Se falhar ao baixar o cabeçalho da nuvem, ignora
                    pass

                if processo:
                    texto_processo = f"Processo: {processo}"
                    if ano:
                        # USA O VALOR DO CAMPO COMARCA AQUI
                        texto_processo += f".{ano}.8.13.{comarca}"
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
                
                # Data ajustada para fuso do Brasil (-3h)
                hoje = datetime.datetime.utcnow() - datetime.timedelta(hours=3)
                meses = ["janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"]
                # Pega o local do perfil do usuário para a data (ex: Santa Luzia)
                local_data = dados_usuario.get("matricula", "").split(":")[0].strip() or "Santa Luzia"
                data_extenso = f"{local_data}, {hoje.day} de {meses[hoje.month - 1]} de {hoje.year}."
                
                p_data = doc.add_paragraph(data_extenso)
                p_data.alignment = WD_ALIGN_PARAGRAPH.CENTER

                doc.add_paragraph("")
                
                # Assinatura dinâmica da Nuvem
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
                
                # Salva o arquivo na pasta do usuário na Nuvem
                caminho_salvamento = f"{usuario_atual}/{nome_arquivo}"
                
                supabase.storage.from_("certidoes_usuarios").upload(
                    file=buffer.getvalue(),
                    path=caminho_salvamento,
                    file_options={"content-type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document"}
                )

            st.success(f"✅ Certidão detalhada salva na sua conta na Nuvem!")
            st.download_button(
                label="📥 Baixar Documento Word Agora",
                data=buffer,
                file_name=nome_arquivo,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                type="primary",
                use_container_width=True
            )

    # ==========================================
    # OPÇÃO B: CERTIDÃO SIMPLES
    # ==========================================
    elif tipo_certidao == "Certidão Negativa Simples (Opções Rápidas)":
        
        # Situação Principal em colunas lado a lado
        st.write("**Situação Principal:**")
        c1, c2 = st.columns(2)
        with c1:
            local_fechado = st.checkbox("Local Fechado", key="sit1")
        with c2:
            pessoa_nao_enc = st.checkbox("Pessoa Não Encontrada", key="sit2")
            
        c1, c2 = st.columns(2)
        with c1:
            nao_localizei = st.checkbox("Não Localizei a Pessoa", key="sit3")
        with c2:
            st.empty() # Espaço vazio para alinhar

        st.divider()
        
        c_inf1, c_inf2 = st.columns([2, 1])
        with c_inf1:
            nome_inf = st.text_input("Nome do Informante:", key="inf_nome")
        with c_inf2:
            obteve_inf = st.checkbox("Obteve Informações?", key="inf_sim", value=True)

        st.write("**Detalhes das Informações Obtidas:**")
        c_m1, c_m2 = st.columns(2)
        with c_m1:
            mudou_se = st.checkbox("Mudou-se", key="mot1")
            nao_reside = st.checkbox("Não Reside no Local", key="mot2")
            nao_fica = st.checkbox("Não fica ali", key="mot3")
        with c_m2:
            nao_trabalha = st.checkbox("Não trabalha ali", key="mot4")
            falecido = st.checkbox("Falecido", key="mot5")
            st.empty()

        st.divider()
        st.write("**Condições do Local:**")
        c1, c2 = st.columns(2)
        with c1:
            local_perigoso = st.checkbox("Local Perigoso", key="cond1")
            zona_rural = st.checkbox("Zona Rural", key="cond2")
            chuva = st.checkbox("Chuva/Meteorológico", key="cond3")
        with c2:
            blocos = st.checkbox("Blocos/Portaria Vazia", key="cond4")
            receio = st.checkbox("Medo/Receio do Processo", key="cond5")
            st.empty()

        st.markdown("---")
        observacoes = st.text_area("Observações Extras:", height=68, key="obs_simples")
        st.divider()

        if st.button("Salvar na Nuvem / Gerar DOCX (Simples)", type="primary", use_container_width=True):
            with st.spinner("Construindo certidão simples e salvando na nuvem..."):
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

                sits = []
                if local_fechado: sits.append("porque o local foi encontrado fechado e mesmo após chamar várias vezes, ninguém atendeu")
                if pessoa_nao_enc: sits.append("porque não a encontrei no local")
                if nao_localizei: sits.append("porque não a localizei")
                
                txt_situacao = ""
                if sits:
                    txt_situacao = ", ".join(sits[:-1]) + (" e " + sits[-1] if len(sits) > 1 else sits[0]) + ". "
                else:
                    txt_situacao = "não foi possível a sua realização. "

                paragrafo_unico = (
                    f"Certifico e dou fé que, em cumprimento ao mandado anexo, dirigi-me {txt_endereco}{texto_data_hora} "
                    f"e, deixei de citar/intimar/notificar {txt_pessoa}, {txt_situacao}"
                )

                if obteve_inf and nome_inf:
                    paragrafo_unico += f"Conforme informações obtidas no local com Sr.(a) {nome_inf}, informou que, "
                elif obteve_inf and not nome_inf:
                    paragrafo_unico += "Conforme informações prestadas no local por pessoa que não quis se identificar, informou que, "
                else:
                    paragrafo_unico += "Procurei obter informações junto aos moradores vizinhos locais, e não obtive êxito, uma vez que ninguém forneceu informações. "

                mots = []
                if mudou_se: mots.append("a pessoa procurada não reside mais no local, tendo se mudado sem deixar meios para contato")
                if nao_reside: mots.append("a pessoa procurada não reside no local referido")
                if nao_fica: mots.append("a pessoa procurada reside no local, mas quase não fica no mesmo, onde nos dias e horários acima não foi localizada")
                if nao_trabalha: mots.append("a pessoa procurada não trabalha no local")
                if falecido: mots.append("a pessoa procurada já se encontra falecida")
                
                if obteve_inf and mots:
                    paragrafo_unico += ", ".join(mots[:-1]) + (" e " + mots[-1] if len(mots) > 1 else mots[0]) + ". "

                conds = []
                if local_perigoso: conds.append("Informo também que o local é conhecidamente de grande periculosidade, o que quase sempre inviabiliza a obtenção de informações, pois os moradores ficam receosos de envolvimento com o processo e suas consequências, onde conversei com alguns vizinhos, que não quiseram se identificar, e ninguém soube informar detalhes sobre o possível horário/local para encontrar a pessoa procurada")
                if zona_rural: conds.append("Informo que o local é uma zona rural com difícil acesso, localização difícil, numeração irregular com muitas casas sem números na porta, o que causa desconforto nos moradores em fornecer informações precisas sobre o local/horário para encontrar a pessoa procurada")
                if chuva: conds.append("Certifico que a execução da diligência restou dificultada em virtude das adversas condições meteorológicas no momento do ato, caracterizadas por intensa precipitação pluviométrica. Ressalto que tal circunstância, além de elevar significativamente o ruído ambiental comprometendo a audibilidade do chamamento realizado no portão, bem como ocasiona o natural recolhimento dos moradores no interior da residência com janelas e portas cerradas, o que obstaculizou a percepção da minha presença e, consequentemente, impediu o efetivo atendimento")
                if blocos: conds.append("Informo também que o local é um condomínio de edifícios com vários blocos de apartamentos em seu interior; possui portaria na entrada do condomínio, mas não existe nenhum porteiro no local em nenhum horário; possui um interfone na entrada que é o único meio de contato com os apartamentos dentro do condomínio, mas aparentemente esse interfone não está funcionando, pois toquei várias vezes e ninguém atendeu; procurei informações com moradores que estavam saindo do condomínio sobre o possível contato com a pessoa procurada, mas ninguém soube informar se o mesmo reside no condomínio dizendo “são muitos moradores e não conhecemos todo mundo”, afirmando não saber informar também o possível horário para encontrá-la")
                if receio: conds.append("Procurei informações com vizinhos sobre o horário/local para encontrar a pessoa procurada, mas os moradores ficam receosos de envolvimento com o processo e suas consequências, onde conversei com alguns vizinhos, que não quiseram se identificar, e ninguém soube informar detalhes sobre o possível horário/local para encontrar a pessoa procurada")
                
                if conds:
                    paragrafo_unico += ". ".join(conds) + ". "

                if observacoes:
                    paragrafo_unico += observacoes.strip() + " "

                doc = Document()
                style = doc.styles['Normal']
                font = style.font
                font.name = 'Times New Roman'
                font.size = Pt(12)

                # Cabeçalho dinâmico da Nuvem
                try:
                    cabecalho_bytes = supabase.storage.from_("imagens_sistema").download("cabecalho.png")
                    cabecalho_stream = BytesIO(cabecalho_bytes)
                    p_img_cabecalho = doc.add_paragraph()
                    p_img_cabecalho.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    run_img_cab = p_img_cabecalho.add_run()
                    run_img_cab.add_picture(cabecalho_stream, width=Cm(16))
                except:
                    pass

                if processo:
                    texto_processo = f"Processo: {processo}"
                    if ano:
                        # USA O VALOR DO CAMPO COMARCA AQUI TAMBÉM
                        texto_processo += f".{ano}.8.13.{comarca}"
                    doc.add_paragraph(texto_processo)
                    
                if mandado:
                    doc.add_paragraph(f"Mandado nº: {mandado}")
                    
                doc.add_paragraph("")

                p_titulo = doc.add_paragraph()
                run_titulo = p_titulo.add_run("CERTIDÃO")
                run_titulo.bold = True
                p_titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
                
                doc.add_paragraph("")

                p_corpo = doc.add_paragraph(paragrafo_unico.strip())
                p_corpo.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                p_corpo.paragraph_format.first_line_indent = Pt(35.4) 
                
                doc.add_paragraph("")

                p_fechamento = doc.add_paragraph("Devolvo o mandado para os devidos fins. O referido é verdade. Dou fé.")
                p_fechamento.alignment = WD_ALIGN_PARAGRAPH.CENTER
                
                # Data ajustada para fuso do Brasil (-3h)
                hoje = datetime.datetime.utcnow() - datetime.timedelta(hours=3)
                meses = ["janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"]
                # Pega o local do perfil do usuário para a data
                local_data = dados_usuario.get("matricula", "").split(":")[0].strip() or "Santa Luzia"
                data_extenso = f"{local_data}, {hoje.day} de {meses[hoje.month - 1]} de {hoje.year}."
                
                p_data = doc.add_paragraph(data_extenso)
                p_data.alignment = WD_ALIGN_PARAGRAPH.CENTER

                doc.add_paragraph("")
                
                # Assinatura dinâmica da Nuvem
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
                nome_arquivo = f"Certidao_Simples_{processo}_{data_arquivo}.docx" if processo else f"Certidao_Simples_{data_arquivo}.docx"
                
                # Salva na Nuvem
                caminho_salvamento = f"{usuario_atual}/{nome_arquivo}"
                
                supabase.storage.from_("certidoes_usuarios").upload(
                    file=buffer.getvalue(),
                    path=caminho_salvamento,
                    file_options={"content-type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document"}
                )

            st.success(f"✅ Certidão simples salva na sua conta na Nuvem!")
            st.download_button(
                label="📥 Baixar Documento Word Agora",
                data=buffer,
                file_name=nome_arquivo,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                type="primary",
                use_container_width=True
            )
