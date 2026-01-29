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
modo_teste = True  # True = envia só para Rafael; False = envia para os assessores

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
    st.title("💸 Relatório de Saldo em Conta Corrente")

    st.write("📂 Faça o upload do arquivo:")
    btg_file = st.file_uploader("Base BTG (Conta + Nome + Assessor + Conta Corrente)", type=["xlsx"])

    if btg_file:
        # Carregar arquivo
        df_btg = pd.read_excel(btg_file)
        
        # Verificar colunas necessárias
        colunas_necessarias = ["Conta", "Nome", "Assessor", "Conta Corrente"]
        colunas_faltando = [col for col in colunas_necessarias if col not in df_btg.columns]
        
        if colunas_faltando:
            st.error(f"❌ Colunas faltando no arquivo: {', '.join(colunas_faltando)}")
            st.info(f"📋 Colunas encontradas: {', '.join(df_btg.columns.tolist())}")
            return
        
        # Renomear para padronizar
        df_btg = df_btg.rename(columns={
            "Conta": "Conta Cliente",
            "Nome": "Nome Cliente",
            "Conta Corrente": "Saldo CC"
        })
        
        # Converter Saldo CC para numérico e preencher valores nulos com 0
        df_btg["Saldo CC"] = pd.to_numeric(df_btg["Saldo CC"], errors='coerce').fillna(0)
        
        # 🔥 Filtrar clientes com Saldo em Conta diferente de zero
        df_final = df_btg[df_btg["Saldo CC"] != 0][["Conta Cliente", "Nome Cliente", "Assessor", "Saldo CC"]].copy()
        
        # Mapear e-mails dos assessores usando busca inteligente
        def buscar_email_assessor(nome_assessor):
            """Busca email usando a mesma lógica inteligente"""
            emails = st.secrets["emails_assessores"]
            
            # Tentar busca direta primeiro
            if nome_assessor in emails:
                return emails[nome_assessor]
            
            # Usar a função de busca inteligente
            chave_encontrada, _ = buscar_assessor_secrets(nome_assessor, 
                                                          {k: {"dummy": "data"} for k in emails.keys()})
            if chave_encontrada and chave_encontrada in emails:
                return emails[chave_encontrada]
            
            return None
        
        df_final["Email Assessor"] = df_final["Assessor"].apply(buscar_email_assessor)
        
        # 🖥️ Formatar valores no padrão brasileiro e aplicar cores (para exibir no app)
        df_formatado = df_final.copy()
        df_formatado["Saldo CC"] = df_formatado["Saldo CC"].apply(formatar_brasileiro)
        
        # Exibir tabela com scroll e formatação
        st.subheader("📊 Dados Processados (Saldo em Conta ≠ 0)")
        tabela_html = df_formatado.drop(columns=["Email Assessor"]).to_html(escape=False, index=False)
        tabela_com_scroll = f"""
        <div style="overflow:auto; max-height:500px; border:1px solid #ddd; padding:8px">
            {tabela_html}
        </div>
        """
        st.markdown(tabela_com_scroll, unsafe_allow_html=True)
        
        st.success(f"✅ {df_final.shape[0]} clientes com Saldo em Conta ≠ 0 processados.")
        
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

        if st.button("📱 Testar envio de WhatsApp (emails desabilitados)"):
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
                    nome_completo_assessor = "Rafael"
                    # Telefone apenas se for realmente o Rafael
                    if dados_assessor:
                        telefone_assessor = dados_assessor.get("telefone", None)
                    else:
                        telefone_assessor = None
                else:
                    # Email já foi mapeado no dataframe com busca inteligente
                    email_destino = grupo["Email Assessor"].iloc[0]
                    
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
                total_clientes = len(grupo)

                resumo_html = f"""
                <p>Olá {primeiro_nome},</p>
                <p>Aqui está o resumo de Saldo em Conta Corrente dos seus clientes:</p>
                <ul>
                    <li><strong>Total de clientes:</strong> {total_clientes}</li>
                    <li><strong>Saldo total em Conta:</strong> {formatar_brasileiro(saldo_cc_total)}</li>
                </ul>
                <p>O relatório detalhado com a lista de clientes está em anexo.</p>
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
                msg["Subject"] = f"📩 Saldo em Conta Corrente – {data_hoje}"

                msg.attach(MIMEText(resumo_html, "html"))
                anexo = MIMEApplication(output.read(), Name=nome_arquivo)
                anexo["Content-Disposition"] = f'attachment; filename="{nome_arquivo}"'
                msg.attach(anexo)

                # 📧 ENVIAR EMAIL - DESABILITADO PARA TESTE
                # try:
                #     with smtplib.SMTP("smtp.gmail.com", 587) as smtp:
                #         smtp.starttls()
                #         smtp.login(email_remetente, senha_app)
                #         smtp.send_message(msg)
                #     enviados_email += 1
                #     st.success(f"📨 E-mail enviado para {assessor} ({email_destino})")
                # except Exception as e:
                #     st.error(f"❌ Erro ao enviar e-mail para {assessor}: {e}")
                
                st.info(f"📧 Envio de e-mail desabilitado (modo teste WhatsApp)")

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

            # ✅ Enviar relatório consolidado para Rafael (EMAIL) - DESABILITADO PARA TESTE
            # try:
            #     # 🧮 Resumo consolidado geral
            #     saldo_cc_total = df_final["Saldo CC"].sum()
            #     total_clientes = len(df_final)

            #     resumo_geral_html = f"""
            #     <p>Olá Rafael,</p>
            #     <p>Segue o relatório consolidado com todos os dados enviados aos assessores:</p>
            #     <ul>
            #         <li><strong>Total de clientes:</strong> {total_clientes}</li>
            #         <li><strong>Saldo total em Conta:</strong> {formatar_brasileiro(saldo_cc_total)}</li>
            #     </ul>
            #     <p>Relatório detalhado em anexo.</p>
            #     """

            #     output_consolidado = io.BytesIO()
            #     df_final.drop(columns=["Email Assessor"]).to_excel(output_consolidado, index=False)
            #     output_consolidado.seek(0)

            #     # 📎 Nome do arquivo consolidado com data
            #     nome_arquivo_consolidado = f"Saldo_em_Conta_Consolidado_{data_hoje}.xlsx"

            #     msg_resumo = MIMEMultipart()
            #     msg_resumo["From"] = email_remetente
            #     msg_resumo["To"] = "rafael@convexainvestimentos.com"
            #     msg_resumo["Subject"] = f"📊 Relatório Consolidado – {data_hoje}"

            #     msg_resumo.attach(MIMEText(resumo_geral_html, "html"))
            #     anexo_resumo = MIMEApplication(output_consolidado.read(), Name=nome_arquivo_consolidado)
            #     anexo_resumo["Content-Disposition"] = f'attachment; filename="{nome_arquivo_consolidado}"'
            #     msg_resumo.attach(anexo_resumo)

            #     with smtplib.SMTP("smtp.gmail.com", 587) as smtp:
            #         smtp.starttls()
            #         smtp.login(email_remetente, senha_app)
            #         smtp.send_message(msg_resumo)

            #     st.success("📨 Relatório consolidado enviado para rafael@convexainvestimentos.com.")

            # except Exception as e:
            #     st.error(f"❌ Erro ao enviar relatório consolidado: {e}")
            
            st.info("📧 Envio de relatório consolidado desabilitado (modo teste WhatsApp)")

            # 📊 Resumo final
            st.info(f"✅ {enviados_email} e-mails enviados com sucesso.")
            st.info(f"✅ {enviados_whatsapp} mensagens WhatsApp enviadas com sucesso.")

# Executar o aplicativo
if __name__ == "__main__":
    executar()
