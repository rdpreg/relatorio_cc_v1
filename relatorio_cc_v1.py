import streamlit as st
import pandas as pd
from datetime import datetime
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText
from email.utils import formataddr
import io
import requests

# 🔥 Flag para alternar entre teste e produção
modo_teste = False  # True = envia só para Rafael; False = envia para os assessores

def formatar_brasileiro(valor):
    """Formata número no padrão brasileiro com R$"""
    texto = f"R$ {valor:,.2f}".replace(",", "v").replace(".", ",").replace("v", ".")
    if valor < 0:
        return f"<span style='color: red;'>{texto}</span>"
    return texto

def formatar_brasileiro_whatsapp(valor):
    """Formata número no padrão brasileiro com R$ (sem HTML para WhatsApp)"""
    return f"R$ {valor:,.2f}".replace(",", "v").replace(".", ",").replace("v", ".")

def buscar_assessor_secrets(nome_assessor, secrets_assessores):
    """
    Busca assessor nos secrets de forma inteligente:
    - Ignora maiúsculas/minúsculas
    - Busca por nome parcial (ex: "Rafael Dadoorian" encontra "RAFAEL DADOORIAN PREGNOLATI")
    
    Retorna: (chave_encontrada, dados_assessor) ou (None, None)
    """
    nome_busca = nome_assessor.lower().strip()
    
    # Tentar match exato primeiro
    for chave, dados in secrets_assessores.items():
        if chave.lower().strip() == nome_busca:
            return chave, dados
    
    # Tentar match parcial (busca palavras do secrets no nome da planilha)
    for chave, dados in secrets_assessores.items():
        palavras_chave = chave.lower().split()
        # Se todas as palavras da chave existem no nome da planilha
        if all(palavra in nome_busca for palavra in palavras_chave):
            return chave, dados
    
    # Tentar match reverso (busca palavras da planilha no secrets)
    palavras_busca = nome_busca.split()
    for chave, dados in secrets_assessores.items():
        chave_lower = chave.lower()
        # Se pelo menos 2 palavras da planilha existem na chave
        matches = sum(1 for palavra in palavras_busca if palavra in chave_lower)
        if matches >= 2:
            return chave, dados
    
    return None, None

def enviar_whatsapp(telefone, mensagem):
    """Envia mensagem via WhatsApp usando ZAPI"""
    try:
        url = st.secrets["zapi"]["url"]
        headers = {
            "Client-Token": st.secrets["zapi"]["client_token"],
            "Content-Type": "application/json"
        }
        payload = {
            "phone": telefone,
            "message": mensagem
        }
        
        # Log do que está sendo enviado
        st.write(f"🔍 DEBUG - URL: {url}")
        st.write(f"🔍 DEBUG - Telefone: {telefone}")
        st.write(f"🔍 DEBUG - Tamanho da mensagem: {len(mensagem)} caracteres")
        
        response = requests.post(url, json=payload, headers=headers, timeout=30)
        
        st.write(f"🔍 DEBUG - Status Code: {response.status_code}")
        st.write(f"🔍 DEBUG - Response: {response.text}")
        
        if response.status_code == 200 or response.status_code == 201:
            return True, "Mensagem enviada com sucesso"
        else:
            return False, f"Erro na API: {response.status_code} - {response.text}"
    except requests.exceptions.Timeout:
        return False, "Timeout na requisição (30s)"
    except requests.exceptions.ConnectionError:
        return False, "Erro de conexão com a API"
    except Exception as e:
        return False, f"Erro ao enviar: {str(e)}"

def executar():
    st.title("💸 Envio de Fluxo Financeiro (Produção)")

    st.write("📂 Faça o upload dos arquivos necessários:")
    btg_file = st.file_uploader("1️⃣ Base BTG (conta + assessor)", type=["xlsx"])
    saldo_file = st.file_uploader("2️⃣ Saldo D0 + Valores a Receber Projetados (RF + VT)", type=["xlsx"])

    if btg_file and saldo_file:
        # Carregar bases
        df_btg = pd.read_excel(btg_file)
        df_saldo = pd.read_excel(saldo_file)

        # Padronizar colunas
        df_btg = df_btg.rename(columns={"Conta": "Conta Cliente", "Nome": "Nome Cliente"})
        df_saldo = df_saldo.rename(columns={"Conta": "Conta Cliente", "Saldo": "Saldo CC"})

        # Mesclar dados
        df_merged = df_btg.merge(df_saldo, on="Conta Cliente", how="left")

        # Preencher valores nulos com 0
        df_merged.fillna(0, inplace=True)

        # Calcular coluna Saldo Projetado
        df_merged["Saldo Projetado"] = (
            df_merged["Saldo CC"] +
            df_merged["D+1"] +
            df_merged["D+2"] +
            df_merged["D+3"]
        )

        # 🔥 Filtrar clientes com Saldo Projetado diferente de zero
        df_final = df_merged[
            df_merged["Saldo Projetado"] != 0
        ][[
            "Conta Cliente", "Nome Cliente", "Assessor",
            "Saldo CC", "D+1", "D+2", "D+3", "Saldo Projetado"
        ]]

        # Mapear e-mails dos assessores
        emails_assessores = st.secrets["emails_assessores"]
        df_final["Email Assessor"] = df_final["Assessor"].map(emails_assessores)

        # 🖥️ Formatar valores no padrão brasileiro e aplicar cores (para exibir no app)
        df_formatado = df_final.copy()
        for col in ["Saldo CC", "D+1", "D+2", "D+3", "Saldo Projetado"]:
            df_formatado[col] = df_formatado[col].apply(formatar_brasileiro)

        # Exibir tabela com scroll e formatação
        st.subheader("📊 Dados Processados (Saldo Projetado ≠ 0)")
        tabela_html = df_formatado.drop(columns=["Email Assessor"]).to_html(escape=False, index=False)
        tabela_com_scroll = f"""
        <div style="overflow:auto; max-height:500px; border:1px solid #ddd; padding:8px">
            {tabela_html}
        </div>
        """
        st.markdown(tabela_com_scroll, unsafe_allow_html=True)

        st.success(f"✅ {df_final.shape[0]} clientes com Saldo Projetado ≠ 0 processados com sucesso.")
        
        # 🔍 Debug: Mostrar assessores únicos encontrados
        assessores_unicos = df_final["Assessor"].unique()
        st.info(f"📋 Assessores encontrados no arquivo: {', '.join(assessores_unicos)}")
        
        # 🔍 Debug: Verificar se há telefones cadastrados
        assessores_com_telefone = []
        assessores_sem_telefone = []
        for assessor in assessores_unicos:
            chave_assessor, dados_assessor = buscar_assessor_secrets(assessor, st.secrets["assessores"])
            if dados_assessor and dados_assessor.get("telefone"):
                telefone = dados_assessor.get("telefone")
                assessores_com_telefone.append(f"{assessor} → {chave_assessor} ({telefone})")
            else:
                assessores_sem_telefone.append(assessor)
        
        if assessores_com_telefone:
            st.success(f"✅ Assessores com telefone cadastrado:\n" + "\n".join([f"  • {a}" for a in assessores_com_telefone]))
        if assessores_sem_telefone:
            st.warning(f"⚠️ Assessores SEM telefone cadastrado: {', '.join(assessores_sem_telefone)}")

        if st.button("📧 Enviar e-mails e WhatsApp aos assessores"):
            email_remetente = st.secrets["email"]["remetente"]
            senha_app = st.secrets["email"]["senha_app"]
            data_hoje = datetime.now().strftime("%d-%m-%Y")

            enviados_email = 0
            enviados_whatsapp = 0

            # 🔄 Loop pelos assessores
            for assessor, grupo in df_final.groupby("Assessor"):
                # 🔍 Buscar assessor nos secrets de forma inteligente
                chave_assessor, dados_assessor = buscar_assessor_secrets(assessor, st.secrets["assessores"])
                
                # 🔥 Se modo_teste=True, envia tudo para Rafael
                if modo_teste:
                    email_destino = "rafael@convexainvestimentos.com"
                    primeiro_nome = "Rafael"
                    telefone_assessor = "5521980039394"  # Seu telefone para teste
                    nome_completo_assessor = "Rafael"
                else:
                    # Buscar email (tentar busca inteligente primeiro)
                    email_destino = None
                    if chave_assessor:
                        # Buscar email usando a chave encontrada
                        email_destino = st.secrets["emails_assessores"].get(chave_assessor)
                    
                    # Se não encontrou, tentar busca direta
                    if not email_destino:
                        email_destino = st.secrets["emails_assessores"].get(assessor)
                    
                    # Pegar primeiro nome do assessor
                    primeiro_nome = assessor.strip().split()[0].capitalize()
                    
                    # Buscar telefone e nome completo
                    if dados_assessor:
                        telefone_assessor = dados_assessor.get("telefone", None)
                        nome_completo_assessor = dados_assessor.get("nome", primeiro_nome)
                        st.info(f"✅ Assessor '{assessor}' mapeado para '{chave_assessor}' nos secrets")
                    else:
                        telefone_assessor = None
                        nome_completo_assessor = primeiro_nome
                        st.warning(f"⚠️ Assessor '{assessor}' não encontrado nos secrets")

                if pd.isna(email_destino):
                    st.warning(f"⚠️ Assessor {assessor} sem e-mail definido. Pulando envio.")
                    continue

                # 🧮 Resumo consolidado do assessor (EMAIL)
                saldo_cc_total = grupo["Saldo CC"].sum()
                saldo_d1_total = grupo["D+1"].sum()
                saldo_d2_total = grupo["D+2"].sum()
                saldo_d3_total = grupo["D+3"].sum()

                resumo_html = f"""
                <p>Olá {primeiro_nome},</p>
                <p>Aqui estão os dados de Saldo em Conta consolidados. O relatório detalhado está em anexo.</p>
                <ul>
                    <li><strong>Saldo em Conta:</strong> {formatar_brasileiro(saldo_cc_total)}</li>
                    <li><strong>Saldo a entrar em D+1:</strong> {formatar_brasileiro(saldo_d1_total)}</li>
                    <li><strong>Saldo a entrar em D+2:</strong> {formatar_brasileiro(saldo_d2_total)}</li>
                    <li><strong>Saldo a entrar em D+3:</strong> {formatar_brasileiro(saldo_d3_total)}</li>
                </ul>
                <p>Abraços,<br>Equipe Convexa</p>
                """

                # Gerar anexo Excel com números puros
                output = io.BytesIO()
                grupo.drop(columns=["Email Assessor"]).to_excel(output, index=False)
                output.seek(0)

                # 📎 Nome do arquivo com data
                nome_arquivo = f"Saldo_em_Conta_{data_hoje}.xlsx"

                # Montar e-mail
                msg = MIMEMultipart()
                msg["From"] = formataddr(("Backoffice Convexa", email_remetente))
                msg["To"] = email_destino
                msg["Subject"] = f"📩 Fluxo Financeiro – {data_hoje}"

                msg.attach(MIMEText(resumo_html, "html"))
                anexo = MIMEApplication(output.read(), Name=nome_arquivo)
                anexo["Content-Disposition"] = f'attachment; filename="{nome_arquivo}"'
                msg.attach(anexo)

                # 📧 ENVIAR EMAIL
                try:
                    with smtplib.SMTP("smtp.gmail.com", 587) as smtp:
                        smtp.starttls()
                        smtp.login(email_remetente, senha_app)
                        smtp.send_message(msg)
                    enviados_email += 1
                    st.success(f"📨 E-mail enviado para {assessor} ({email_destino})")
                except Exception as e:
                    st.error(f"❌ Erro ao enviar e-mail para {assessor}: {e}")

                # 📱 ENVIAR WHATSAPP
                if telefone_assessor:
                    # Montar lista de clientes para WhatsApp
                    lista_clientes = ""
                    for _, cliente in grupo.iterrows():
                        conta = cliente["Conta Cliente"]
                        nome = cliente["Nome Cliente"]
                        valor_cc = formatar_brasileiro_whatsapp(cliente["Saldo CC"])
                        lista_clientes += f"• {conta} - {nome} - {valor_cc}\n"

                    # Mensagem WhatsApp
                    mensagem_whatsapp = f"""Olá {nome_completo_assessor if not modo_teste else primeiro_nome}

Você tem o total de {formatar_brasileiro_whatsapp(saldo_cc_total)} em conta.

É importante trabalhar para alocar antes que o cliente envie para outro banco.

Segue a lista de clientes:
{lista_clientes}"""

                    # Preview da mensagem
                    with st.expander(f"📱 Preview da mensagem para {assessor}"):
                        st.text(mensagem_whatsapp)
                    
                    st.info(f"📱 Tentando enviar WhatsApp para {assessor} no número {telefone_assessor}...")
                    
                    # Enviar via ZAPI
                    sucesso, msg_retorno = enviar_whatsapp(telefone_assessor, mensagem_whatsapp)
                    if sucesso:
                        enviados_whatsapp += 1
                        st.success(f"✅ WhatsApp enviado para {assessor} ({telefone_assessor})")
                    else:
                        st.error(f"❌ Erro ao enviar WhatsApp para {assessor}: {msg_retorno}")
                else:
                    st.warning(f"⚠️ Assessor {assessor} sem telefone definido no secrets. Pulando envio de WhatsApp.")

            # ✅ Enviar relatório consolidado para Rafael (EMAIL)
            try:
                # 🧮 Resumo consolidado geral
                saldo_cc_total = df_final["Saldo CC"].sum()
                saldo_d1_total = df_final["D+1"].sum()
                saldo_d2_total = df_final["D+2"].sum()
                saldo_d3_total = df_final["D+3"].sum()

                resumo_geral_html = f"""
                <p>Olá Rafael,</p>
                <p>Segue o relatório consolidado com todos os dados enviados aos assessores:</p>
                <ul>
                    <li><strong>Saldo em Conta:</strong> {formatar_brasileiro(saldo_cc_total)}</li>
                    <li><strong>Saldo em D+1:</strong> {formatar_brasileiro(saldo_d1_total)}</li>
                    <li><strong>Saldo em D+2:</strong> {formatar_brasileiro(saldo_d2_total)}</li>
                    <li><strong>Saldo em D+3:</strong> {formatar_brasileiro(saldo_d3_total)}</li>
                </ul>
                <p>Relatório detalhado em anexo.</p>
                """

                output_consolidado = io.BytesIO()
                df_final.drop(columns=["Email Assessor"]).to_excel(output_consolidado, index=False)
                output_consolidado.seek(0)

                # 📎 Nome do arquivo consolidado com data
                nome_arquivo_consolidado = f"Saldo_em_Conta_{data_hoje}.xlsx"

                msg_resumo = MIMEMultipart()
                msg_resumo["From"] = email_remetente
                msg_resumo["To"] = "rafael@convexainvestimentos.com"
                msg_resumo["Subject"] = f"📊 Relatório Consolidado – {data_hoje}"

                msg_resumo.attach(MIMEText(resumo_geral_html, "html"))
                anexo_resumo = MIMEApplication(output_consolidado.read(), Name=nome_arquivo_consolidado)
                anexo_resumo["Content-Disposition"] = f'attachment; filename="{nome_arquivo_consolidado}"'
                msg_resumo.attach(anexo_resumo)

                with smtplib.SMTP("smtp.gmail.com", 587) as smtp:
                    smtp.starttls()
                    smtp.login(email_remetente, senha_app)
                    smtp.send_message(msg_resumo)

                st.success("📨 Relatório consolidado enviado para rafael@convexainvestimentos.com.")

            except Exception as e:
                st.error(f"❌ Erro ao enviar relatório consolidado: {e}")

            # 📊 Resumo final
            st.info(f"✅ {enviados_email} e-mails enviados com sucesso.")
            st.info(f"✅ {enviados_whatsapp} mensagens WhatsApp enviadas com sucesso.")

# Executar o aplicativo
if __name__ == "__main__":
    executar()
