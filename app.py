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

# CSS para esconder menus do Streamlit e compactar a visualização
st.markdown("""
    <style>
    /* Oculta marcações padrão do Streamlit */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    
    .block-container { padding-top: 1rem; padding-bottom: 1rem; }
    h1 { font-size: 22px; text-align: center; margin-bottom: 0; padding-bottom: 0;}
    /* Ajustes finos de margem para elementos específicos */
    .stCheckbox { margin-top: -5px; margin-bottom: -5px; }
    div[role="radiogroup"] { margin-top: -10px; }
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
    
    # Abas para separar Login de Cadastro
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
    # Upload da assinatura (chave única para evitar conflitos)
    arquivo_assinatura = st.file_uploader("Envie a foto da sua assinatura", type=["png", "jpg", "jpeg"], key="uploader_perfil")
    
    if st.button("💾 Salvar Perfil", type="primary", use_container_width=True, key="btn_salvar_perfil"):
        # Atualiza dados de texto no banco_usuarios
        supabase.table("banco_usuarios").update({
            "nome": novo_nome,
            "cargo": novo_cargo,
            "matricula": nova_matricula
        }).eq("usuario", usuario_atual).execute()
        
        # Se houve upload de arquivo de assinatura, gerencia no Storage
        if arquivo_assinatura is not None:
            try:
                # Tenta remover assinatura antiga (se houver) para limpar espaço
                supabase.storage.from_("assinaturas_usuarios").remove([f"{usuario_atual}.png"])
            except:
                pass
            # Faz o upload da nova assinatura, sempre com o nome [usuario].png
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
        # Lista os arquivos dentro da pasta do usuário no bucket certidoes_usuarios
        arquivos_nuvem = supabase.storage.from_("certidoes_usuarios").list(usuario_atual)
    except:
        arquivos_nuvem = []
    
    # Filtra arquivos válidos (descarta pastas vazias do sistema)
    arquivos = [arq for arq in arquivos_nuvem if arq["name"] != ".emptyFolder" and arq["name"] != ""]
    
    if not arquivos:
        st.info("Nenhuma certidão salva ainda.")
    else:
        # Ordena por data de criação (mais recente primeiro)
        arquivos.sort(key=lambda x: x["created_at"], reverse=True)
        
        # Cabeçalho da tabela
        c_sel, c_nome, c_data = st.columns([1, 4, 3])
        c_sel.write("**Selecionar**")
        c_nome.write("**Nome do Arquivo**")
        c_data.write("**Data de Criação**")
        st.divider()
        
        arquivos_selecionados = []
        
        # Loop para criar as linhas da tabela com checkboxes
        for item in arquivos:
            c1, c2, c3 = st.columns([1, 4, 3])
            try:
                # Pega a data do banco (UTC), remove 'Z' e ajusta o fuso do Brasil (-3h)
                data_str = item["created_at"].replace("Z", "+00:00")
                data_obj = datetime.datetime.fromisoformat(data_str)
                data_br_obj = data_obj.replace(tzinfo=None) - datetime.timedelta(hours=3)
                data_br = data_br_obj.strftime("%d/%m/%Y às %H:%M")
            except:
                data_br = "Data desconhecida"

            with c1:
                # Checkbox com chave única baseada no nome do arquivo
                if st.checkbox("", key=f"chk_file_{item['name']}"):
                    arquivos_selecionados.append(item['name'])
            with c2:
                st.write(item['name'])
            with c3:
                st.write(data_br)
                
        st.divider()
        
        # Ações para arquivos selecionados
        if arquivos_selecionados:
            st.write(f"**{len(arquivos_selecionados)} arquivo(s) selecionado(s)**")
            c_btn1, c_btn2 = st.columns(2)
            
            with c_btn1:
                if st.button("📥 Preparar Download (ZIP)", type="primary", use_container_width=True, key="btn_zip_download"):
                    with st.spinner("Baixando e compactando arquivos da nuvem..."):
                        zip_buffer = BytesIO()
                        with zipfile.ZipFile(zip_buffer, "w") as zip_file:
                            for arq in arquivos_selecionados:
                                # Baixa cada arquivo e adiciona ao ZIP
                                arquivo_bytes = supabase.storage.from_("certidoes_usuarios").download(f"{usuario_atual}/{arq}")
                                zip_file.writestr(arq, arquivo_bytes)
                                
                        # Oferece o arquivo ZIP para download
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
                    # Cria a lista de caminhos completos ([usuario]/[arquivo])
                    caminhos_para_excluir = [f"{usuario_atual}/{arq}" for arq in arquivos_selecionados]
                    # Remove os arquivos do storage
                    supabase.storage.from_("certidoes_usuarios").remove(caminhos_para_excluir)
                    st.success("✅ Arquivos excluídos da nuvem com sucesso!")
                    st.rerun()

# ==========================================
# 6. TELA: PAINEL DO ADMINISTRADOR
# ==========================================
elif menu == "🛡️ Painel do Administrador":
    # Proteção extra de acesso
    if usuario_atual != "10228429":
        st.error("Acesso restrito apenas ao Administrador.")
        st.stop()
        
    st.title("🛡️ Painel de Administração")
    st.write("Área restrita para gestão de oficiais e auditoria de certidões em nuvem.")
    
    # Abas administrativas
    aba_adm1, aba_adm2 = st.tabs(["👥 Gerenciar Usuários", "📊 Auditoria de Certidões Gerais"])
    
    # ABA 1: GERENCIAR USUÁRIOS
    with aba_adm1:
        st.subheader("Oficiais Cadastrados no Sistema")
        # Busca dados básicos de todos os usuários cadastrados
        res_todos = supabase.table("banco_usuarios").select("usuario, nome, cargo, matricula").execute()
        usuarios_cadastrados = res_todos.data
        
        if usuarios_cadastrados:
            for u in usuarios_cadastrados:
                # Usa expansor para detalhar informações de cada usuário
                with st.expander(f"👤 Usuário: {u['usuario']} — Nome: {u.get('nome') or 'Não preenchido'}"):
                    st.write(f"**Cargo:** {u.get('cargo')}")
                    st.write(f"**Matrícula:** {u.get('matricula')}")
                    
                    # Permite excluir outros usuários, mas não o próprio administrador
                    if u['usuario'] != usuario_atual:
                        if st.button(f"🗑️ Excluir usuário {u['usuario']}", key=f"del_adm_usr_{u['usuario']}", use_container_width=True):
                            # Exclui da tabela banco_usuarios
                            supabase.table("banco_usuarios").delete().eq("usuario", u['usuario']).execute()
                            # NOTA: Idealmente excluiria também a pasta no Storage, mas exige listagem e remoção recursiva arquivo por arquivo via API.
                            st.success(f"Usuário {u['usuario']} removido com sucesso!")
                            st.rerun()
                    else:
                        st.caption("*(Esta é a sua conta de Administrador principal)*")
        else:
            st.info("Nenhum usuário encontrado.")

    # ABA 2: AUDITORIA DE CERTIDÕES
    with aba_adm2:
        st.subheader("Certidões Geradas por Todos os Oficiais")
        st.write("Inspecione, baixe ou exclua os arquivos salvos por qualquer oficial na nuvem.")
        
        try:
            # Lista as pastas (nomes dos usuários) na raiz do bucket certidoes_usuarios
            pastas_usuarios = supabase.storage.from_("certidoes_usuarios").list()
        except:
            pastas_usuarios = []
            
        if not pastas_usuarios:
            st.info("Nenhuma pasta de certidão encontrada na nuvem.")
        else:
            # Loop por cada pasta de oficial
            for pasta in pastas_usuarios:
                nome_oficial = pasta["name"]
                # Filtra pastas inválidas ou do sistema
                if nome_oficial and nome_oficial != ".emptyFolder":
                    st.markdown(f"### 📂 Oficial: `{nome_oficial}`")
                    
                    try:
                        # Lista arquivos dentro da pasta deste oficial específico
                        arquivos_oficial = supabase.storage.from_("certidoes_usuarios").list(nome_oficial)
                    except:
                        arquivos_oficial = []
                        
                    # Filtra arquivos válidos dentro da pasta
                    certidoes_validas = [f for f in arquivos_oficial if f["name"] != ".emptyFolder" and f["name"] != ""]
                    
                    if not certidoes_validas:
                        st.caption("Nenhuma certidão gerada por este oficial ainda.")
                    else:
                        # Colunas para exibir nome, botão baixar e botão excluir
                        for arq in certidoes_validas:
                            c_arq_nome, c_btn_dl, c_btn_del = st.columns([4, 2, 2])
                            
                            with c_arq_nome:
                                st.text(arq["name"])
                                
                            with c_btn_dl:
                                # Botão baixar com chave única baseada no oficial e arquivo
                                if st.button("📥 Baixar", key=f"dl_adm_f_{nome_oficial}_{arq['name']}", use_container_width=True):
                                    file_bytes = supabase.storage.from_("certidoes_usuarios").download(f"{nome_oficial}/{arq['name']}")
                                    st.download_button(
                                        label="Confirmar",
                                        data=file_bytes,
                                        file_name=arq["name"],
                                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                        key=f"dl_adm_btn_real_{nome_oficial}_{arq['name']}",
                                        use_container_width=True
                                    )
                                    
                            with c_btn_del:
                                # Botão excluir com chave única
                                if st.button("🗑️ Excluir", key=f"del_adm_f_{nome_oficial}_{arq['name']}", use_container_width=True):
                                    supabase.storage.from_("certidoes_usuarios").remove([f"{nome_oficial}/{arq['name']}"])
                                    st.success("Excluído!")
                                    st.rerun()
                    st.divider()

# ==========================================
# 7. TELA: GERADOR DE CERTIDÃO
# ==========================================
elif menu == "📝 Gerar Certidão":
    st.title("Gerador de Certidão Negativa")
    
    # Bloqueio de segurança se o perfil não estiver configurado
    if not dados_usuario.get("nome"):
        st.warning("⚠️ Você ainda não configurou seu perfil! Vá em 'Meu Perfil' no menu lateral e preencha seus dados antes de gerar certidões.")
        st.stop()

    # Seletor de modelo de certidão
    tipo_certidao = st.selectbox(
        "Selecione o Modelo de Certidão:", 
        ["Certidão Negativa Detalhada", "Certidão Negativa Simples (Opções Rápidas)"]
    )
    
    st.divider()

    # --- CAMPOS COMPARTILHADOS (Cabeçalho e Datas) ---
    c_mandado, c_proc = st.columns([1, 3])
    with c_mandado:
        mandado = st.text_input("Mandado:", placeholder="Ex: 01", key="mandado_geral")
    with c_proc:
        processo = st.text_input("Informe o Processo:", placeholder="Ex: 4400281-16", key="processo_geral")
    
    c_ano, c_comarca = st.columns(2)
    with c_ano:
        ano = st.text_input("Ano:", placeholder="Ex: 2026", key="ano_geral")
    with c_comarca:
        comarca = st.text_input("Código Comarca:", value="0245", placeholder="Ex: 0245", key="comarca_geral")

    c_end, c_pes = st.columns(2)
    with c_end:
        endereco = st.text_input("Endereço (opcional):", placeholder="Se vazio: 'informado no mesmo'", key="endereco_geral")
    with c_pes:
        pessoa = st.text_input("Pessoa procurada:", placeholder="Deixe vazio para termo genérico", key="pessoa_geral")

    st.markdown("---")
    st.write("**Dias e Horários das Diligências:**")
    
    # Grid para entrada de 3 datas/horas
    c_d1, c_h1 = st.columns(2)
    with c_d1:
        d1 = st.text_input("Dia 1", placeholder="Ex: 08/08", key="d1_geral")
    with c_h1:
        h1 = st.text_input("Hora 1", placeholder="Ex: 14:55hs", key="h1_geral")
        
    c_d2, c_h2 = st.columns(2)
    with c_d2:
        d2 = st.text_input("Dia 2", placeholder="Ex: 11/08", key="d2_geral")
    with c_h2:
        h2 = st.text_input("Hora 2", placeholder="Ex: 16:58hs", key="h2_geral")
        
    c_d3, c_h3 = st.columns(2)
    with c_d3:
        d3 = st.text_input("Dia 3", placeholder="Ex: 12/08", key="d3_geral")
    with c_h3:
        h3 = st.text_input("Hora 3", placeholder="Ex: 11:15hs", key="h3_geral")

    st.divider()

    # ==========================================
    # OPÇÃO A: CERTIDÃO DETALHADA
    # ==========================================
    if tipo_certidao == "Certidão Negativa Detalhada":
        # ... (Toda a lógica e campos da certidão detalhada permanecem iguais à versão estável anterior) ...
        st.write("**Deixei de cumprir o ato uma vez que:**")
        sit_c1, sit_c2 = st.columns(2)
        with sit_c1:
            nao_loc_dest = st.checkbox("O destinatário do mandado não foi localizado", key="nao_loc_dest")
        with sit_c2:
            nao_loc_bens = st.checkbox("O(s) bem(ns) indicados não foi(ram) localizado(s)", key="nao_loc_bens")

        # Lista expandível para motivos (seleção múltipla)
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
                    # Chave única baseada no índice para motivos
                    if st.checkbox(m, key=f"mot_det_{idx}"):
                        motivos_selecionados.append(m)

        # Informações sobre o informante (seleção múltipla)
        relacoes_selecionadas = []
        nao_sabe_selecionados = []
        sabe_tel = ""
        sabe_end = ""
        
        with st.expander("👤 Informações sobre o Informante (Se houver)", expanded=False):
            nome_inf_det = st.text_input("Nome do Sr(a):", placeholder="Deixe em branco se não houver informante", key="nome_inf_det")

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
                    # Chave única baseada no índice para relações
                    if st.checkbox(r, key=f"rel_det_{idx}"):
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
                    # Chave única baseada no índice para 'não sabe'
                    if st.checkbox(ns, key=f"ns_det_{idx}"):
                        nao_sabe_selecionados.append(ns)

            st.markdown("---")
            st.write("**Sabendo o informante indicar o:**")
            c_sab1, c_sab2 = st.columns(2)
            with c_sab1:
                sabe_tel = st.text_input("Telefone indicado:", key="sabe_tel_det")
            with c_sab2:
                sabe_end = st.text_input("Endereço correto indicado:", key="sabe_end_det")
        
        # Garante que a variável exista para a lógica de montagem
        if 'nome_inf_det' not in locals():
            nome_inf_det = ""

        with st.expander("📝 Certificações Adicionais e Observações", expanded=False):
            cert_extras = []
            if st.checkbox("Procurei obter informações junto aos moradores/vizinhos locais e não obtive êxito.", key="cert_vizinhos_det"):
                cert_extras.append("procurei obter informações junto aos moradores/vizinhos locais e não obtive êxito.")
            if st.checkbox("Devido à importância do mandado, deixei a cópia para ciência do prazo/data.", key="cert_copia_det"):
                cert_extras.append("devido à importância do mandado e da dificuldade de encontrar a pessoa procurada, deixei a cópia do mandado com o(a) senhor(a) acima mencionado(a) para que a parte/testemunha tome ciência do prazo/data que deverá comparecer em juízo.")
            if st.checkbox("O imóvel é residencial e contém apenas móveis e utensílios domésticos comuns.", key="cert_moveis_det"):
                cert_extras.append("o imóvel é residencial e contém apenas móveis e utensílios domésticos que guarnecem a residência do réu.")

            observacoes_det = st.text_area("Observações Livres:", key="obs_livres_det")

        st.divider()

        # Botão de Geração Detalhada
        if st.button("Salvar na Nuvem / Gerar DOCX (Detalhada)", type="primary", use_container_width=True, key="btn_gerar_docx_det"):
            with st.spinner("Construindo certidão detalhada e salvando na nuvem..."):
                # ... (Lógica de montagem do DOCX detalhado permanece igual à versão funcional anterior) ...
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
                
                if sits:
                    paragrafo += " e ".join(sits) + ". "
                else:
                    paragrafo += "não foi possível a sua realização. "

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

                if cert_extras:
                    paragrafo += f"Certifico também que {'; '.join(cert_extras)}. "
                    
                if observacoes_det:
                    paragrafo += f"{observacoes_det.strip()} "

                # Criação do DOCX
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

                # Processo / Ano / Comarca dinâmicos
                if processo:
                    texto_processo = f"Processo: {processo}"
                    if ano:
                        texto_processo += f".{ano}.8.13.{comarca}"
                    doc.add_paragraph(texto_processo)
                    
                if mandado:
                    doc.add_paragraph(f"Mandado nº: {mandado}")
                    
                doc.add_paragraph("")

                # Título aumentado
                p_titulo = doc.add_paragraph()
                run_titulo = p_titulo.add_run("CERTIDÃO NEGATIVA")
                run_titulo.bold = True
                run_titulo.font.size = Pt(16)
                p_titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
                
                doc.add_paragraph("")

                # Corpo do texto formatado
                p_corpo = doc.add_paragraph(paragrafo.strip())
                p_corpo.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                p_corpo.paragraph_format.first_line_indent = Pt(35.4)
                p_corpo.paragraph_format.line_spacing = 1.5 
                
                doc.add_paragraph("")

                p_fechamento = doc.add_paragraph("Devolvo o mandado para os devidos fins. O referido é verdade. Dou fé.")
                p_fechamento.alignment = WD_ALIGN_PARAGRAPH.CENTER
                
                # Data e Local dinâmicos do Perfil (Fuso -3h)
                hoje = datetime.datetime.utcnow() - datetime.timedelta(hours=3)
                meses = ["janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"]
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
                
                # Rodapé dinâmico (Tamanho 8)
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

                # Nomenclatura dinâmica
                data_arquivo = hoje.strftime("%d-%m-%Y_%Hh%M")
                nome_arquivo = f"Certidao_Negativa_{processo}_{data_arquivo}.docx" if processo else f"Certidao_Negativa_{data_arquivo}.docx"
                
                # Salva na pasta do oficial na Nuvem
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
                use_container_width=True,
                key="dl_btn_real_det"
            )

    # ==========================================
    # OPÇÃO B: CERTIDÃO SIMPLES (CORRIGIDA COM RADIO)
    # ==========================================
    elif tipo_certidao == "Certidão Negativa Simples (Opções Rápidas)":
        
        # --- INPUTS (Restaurada a versão com 'st.radio' baseada no seu exemplo funcional) ---

        # Desfecho Principal (Horizontal)
        situacao_simples = st.radio(
            "Situação Principal:", 
            ["Local Fechado", "Pessoa Não Encontrada", "Não Localizei a Pessoa"],
            index=None, horizontal=True, key="sit_radio_simples"
        )

        st.divider()
        c_inf1, c_inf2 = st.columns([1, 2])
        with c_inf1:
            # Opções de Informações Obtidas (NQI = Não Quis se Identificar)
            obteve_inf_simples = st.radio("Obteve Informações?", ["Sim", "Não", "NQI"], index=None, horizontal=True, key="obteve_inf_radio_simples")
        with c_inf2:
            # O campo de nome fica habilitado apenas se obteve_inf_simples for "Sim"
            nome_inf_simples = st.text_input("Nome do Informante:", disabled=(obteve_inf_simples != "Sim"), key="nome_inf_input_simples")

        # Motivos e Paradeiro (Botões de rádio para seleção única)
        st.write("**Detalhes das Informações Obtidas:**")
        c_m1, c_m2 = st.columns(2)
        with c_m1:
            motivo_simples = st.radio(
                "Motivo:", 
                ["Mudou-se", "Não Reside no Local", "Não fica ali", "Não trabalha ali", "Falecido"], 
                index=None, key="motivo_radio_simples"
            )
        with c_m2:
            nao_sabe_simples = st.radio(
                "O que não sabe?", 
                ["Não Conhece ele", "Não sabe informar", "Não sabe seu endereço"], 
                index=None, key="naosabe_radio_simples"
            )
            paradeiro_simples = st.radio(
                "Paradeiro:", 
                ["Não sabe o paradeiro", "Incerto e Não Sabido"], 
                index=None, key="paradeiro_radio_simples"
            )

        st.divider()
        # Condições Extras do Local
        condicao_simples = st.radio(
            "Condições do Local:", 
            ["Local Perigoso", "Medo Processo", "Zona Rural", "Blocos", "Chuva"], 
            index=None, horizontal=True, key="condicao_radio_simples"
        )

        st.markdown("---")
        # Campo de observações mantido para compatibilidade com o Storage
        observacoes_simples = st.text_area("Observações Extras:", height=68, key="obs_simples_text_area")
        st.divider()

        # --- LÓGICA DO BOTÃO GERAR SIMPLES (Restaurada do seu exemplo funcional e adaptada para Storage) ---
        if st.button("Salvar na Nuvem / Gerar DOCX (Simples)", type="primary", use_container_width=True, key="btn_gerar_docx_simples"):
            with st.spinner("Construindo certidão simples e salvando na nuvem..."):
                # Lógica de montagem do texto (Idêntica ao código fornecido por você)
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
                if situacao_simples == "Local Fechado":
                    txt_situacao = "porque o local foi encontrado fechado e mesmo após chamar várias vezes, ninguém atendeu. "
                elif situacao_simples == "Pessoa Não Encontrada":
                    txt_situacao = "porque não a encontrei no local. "
                elif situacao_simples == "Não Localizei a Pessoa":
                    txt_situacao = "porque não a localizei. "

                paragrafo_unico = (
                    f"Certifico e dou fé que, em cumprimento ao mandado anexo, dirigi-me {txt_endereco}{texto_data_hora} "
                    f"e, deixei de citar/intimar/notificar {txt_pessoa}, {txt_situacao}"
                )

                if obteve_inf_simples == "Sim":
                    paragrafo_unico += f"Conforme informações obtidas no local com Sr.(a) {nome_inf_simples}, informou que, "
                elif obteve_inf_simples == "Não":
                    paragrafo_unico += "Procurei obter informações junto aos moradores vizinhos locais, e não obtive êxito, uma vez que ninguém forneceu informações. "
                elif obteve_inf_simples == "NQI":
                    paragrafo_unico += "Conforme informações prestadas pelo seu vizinho(a), que não quis se identificar, este afirmou que "

                # Lógica baseada nas seleções dos radios (Unica escolha)
                if obteve_inf_simples in ["Sim", "NQI"]:
                    # Motivo
                    if motivo_simples == "Mudou-se":
                        paragrafo_unico += "a pessoa procurada não reside mais no local, tendo se mudado sem deixar meios para contato; "
                    elif motivo_simples == "Não Reside no Local":
                        paragrafo_unico += "a pessoa procurada não reside no local referido; "
                    elif motivo_simples == "Não fica ali":
                        paragrafo_unico += "a pessoa procurada reside no local, mas quase não fica no mesmo, onde nos dias e horários acima não foi localizada; "
                    elif motivo_simples == "Não trabalha ali":
                        paragrafo_unico += "a pessoa procurada não trabalha no local; "
                    elif motivo_simples == "Falecido":
                        paragrafo_unico += "a pessoa procurada já se encontra falecida. "

                    # Não sabe
                    if nao_sabe_simples == "Não Conhece ele":
                        paragrafo_unico += "não conhece a pessoa procurada, não sabendo informar o local/horário para encontrá-la. "
                    elif nao_sabe_simples == "Não sabe informar":
                        paragrafo_unico += "que não sabe informar o dia e horário para encontrá-lo(a). "
                    elif nao_sabe_simples == "Não sabe seu endereço":
                        paragrafo_unico += "que não sabe informar o endereço para encontrá-lo(a). "

                    # Paradeiro
                    if paradeiro_simples == "Não sabe o paradeiro":
                        paragrafo_unico += "não sabe informar seu paradeiro, bem como o local para encontrá-lo. "
                    elif paradeiro_simples == "Incerto e Não Sabido":
                        paragrafo_unico += "Certifico assim, que, com relação ao presente mandado, endereço fornecido e informações obtidas no local, A PESSOA PROCURADA SE ENCONTRA EM LOCAL INCERTO E NÃO SABIDO. "

                # Condições extras (Mantido do exemplo funcional)
                obs_extra = ""
                if condicao_simples == "Chuva":
                    obs_extra = "Certifico que a execução da diligência restou dificultada em virtude das adversas condições meteorológicas no momento do ato, caracterizadas por intensa precipitação pluviométrica. Ressalto que tal circunstância, além de elevar significativamente o ruído ambiental comprometendo a audibilidade do chamamento realizado no portão, bem como ocasiona o natural recolhimento dos moradores no interior da residência com janelas e portas cerradas, o que obstaculizou a percepção da minha presença e, consequentemente, impediu o efetivo atendimento. "
                elif condicao_simples == "Local Perigoso":
                    obs_extra = "Informo também que o local é conhecidamente de grande periculosidade, o que quase sempre inviabiliza a obtenção de informações, pois os moradores ficam receosos de envolvimento com o processo e suas consequências, onde conversei com alguns vizinhos, que não quiseram se identificar, e ninguém soube informar detalhes sobre o possível horário/local para encontrar a pessoa procurada. "
                elif condicao_simples == "Zona Rural":
                    obs_extra = "Informo que o local é uma zona rural com difícil acesso, localização difícil, numeração irregular com muitas casas sem números na porta, o que causa desconforto nos moradores em fornecer informações precisas sobre o local/horário para encontrar a pessoa procurada. "
                elif condicao_simples == "Blocos":
                    obs_extra = "Informo também que o local é um condomínio de edifícios com vários blocos de apartamentos em seu interior; possui portaria na entrada do condomínio, mas não existe nenhum porteiro no local em nenhum horário; possui um interfone na entrada que é o único meio de contato com os apartamentos dentro do condomínio, mas aparentemente esse interfone não está funcionando, pois toquei várias vezes e ninguém atendeu; procurei informações com moradores que estavam saindo do condomínio sobre o possível contato com a pessoa procurada, mas ninguém soube informar se o mesmo reside no condomínio dizendo “são muitos moradores e não conhecemos todo mundo”, afirmando não saber informar também o possível horário para encontrá-la. "
                elif condicao_simples == "Medo Processo":
                    obs_extra = "Procurei informações com vizinhos sobre o horário/local para encontrar a pessoa procurada, mas os moradores ficam receosos de envolvimento com o processo e suas consequências, onde conversei com alguns vizinhos, que não quiseram se identificar, e ninguém soube informar detalhes sobre o possível horário/local para encontrar a pessoa procurada. "

                # Une observações extras e texto livre
                if obs_extra or observacoes_simples:
                    paragrafo_unico += obs_extra + (" " + observacoes_simples if observacoes_simples else "")

                # --- GERAÇÃO DO DOCX (Adaptada para compatibilidade com rodapé do Perfil e Storage) ---
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

                # Processo com Ano e Comarca dinâmicos
                if processo:
                    texto_processo = f"Processo: {processo}"
                    if ano:
                        texto_processo += f".{ano}.8.13.{comarca}"
                    doc.add_paragraph(texto_processo)
                    
                if mandado:
                    doc.add_paragraph(f"Mandado nº: {mandado}")
                    
                doc.add_paragraph("")

                p_titulo = doc.add_paragraph()
                run_titulo = p_titulo.add_run("CERTIDÃO")
                run_titulo.bold = True
                run_titulo.font.size = Pt(16) # Tamanho 16pt
                p_titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
                
                doc.add_paragraph("")

                # Corpo do texto formatado
                p_corpo = doc.add_paragraph(paragrafo_unico.strip())
                p_corpo.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                p_corpo.paragraph_format.first_line_indent = Pt(35.4) # Indentação 1,25cm
                
                doc.add_paragraph("")

                p_fechamento = doc.add_paragraph("Devolvo o mandado para os devidos fins. O referido é verdade. Dou fé.")
                p_fechamento.alignment = WD_ALIGN_PARAGRAPH.CENTER
                
                # Data e Local dinâmicos do Perfil (UTC-3 Brasil)
                hoje = datetime.datetime.utcnow() - datetime.timedelta(hours=3)
                meses = ["janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"]
                # Local padrão dinâmico baseado no perfil
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
                
                # Dados do Oficial (Tamanho 8 - Dinâmico do Perfil)
                p_assinatura = doc.add_paragraph()
                p_assinatura.alignment = WD_ALIGN_PARAGRAPH.CENTER
                
                run_nome = p_assinatura.add_run(f"{dados_usuario['nome']}\n")
                run_nome.bold = True
                run_nome.font.size = Pt(8)
                
                run_cargo = p_assinatura.add_run(f"{dados_usuario['cargo']}\n")
                run_cargo.font.size = Pt(8)
                
                run_matricula = p_assinatura.add_run(f"{dados_usuario['matricula']}")
                run_matricula.font.size = Pt(8)

                # Salva o documento no buffer
                buffer = BytesIO()
                doc.save(buffer)
                buffer.seek(0)

                # Nomenclatura dinâmica
                data_arquivo = hoje.strftime("%d-%m-%Y_%Hh%M")
                nome_arquivo = f"Certidao_Simples_{processo}_{data_arquivo}.docx" if processo else f"Certidao_Simples_{data_arquivo}.docx"
                
                # Salva na pasta do usuário na Nuvem
                caminho_salvamento = f"{usuario_atual}/{nome_arquivo}"
                
                supabase.storage.from_("certidoes_usuarios").upload(
                    file=buffer.getvalue(),
                    path=caminho_salvamento,
                    file_options={"content-type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document"}
                )

            st.success(f"✅ Certidão simples salva na sua conta na Nuvem!")
            # Oferece o arquivo Word gerado para download imediato
            st.download_button(
                label="📥 Baixar Documento Word Agora",
                data=buffer,
                file_name=nome_arquivo,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                type="primary",
                use_container_width=True,
                key="btn_download_real_simples_docx"
            )
