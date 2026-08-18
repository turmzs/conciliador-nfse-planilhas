# 📋 Conciliador NFS-e - Guia de Uso do Sistema de Alertas por Email

## ✅ Funcionalidade Ativada

O sistema agora envia **alertas por email automaticamente** quando encontra divergências nas reconciliações.

## 🚀 Como Ativar

### Passo 1: Criar o Arquivo de Configuração

1. Abra o arquivo `email.config.example` na pasta do projeto
2. Salve uma cópia como `email.config` (na mesma pasta do conciliador_nfse.py)
3. Edite `email.config` com suas credenciais:

```ini
[EMAIL]
smtp_host = smtp.gmail.com
smtp_port = 587
username = seu_email@gmail.com
password = sua_senha_app
destinatarios = email1@example.com,email2@example.com
```

### Passo 2: Para Gmail (Recomendado)

1. Acesse: <https://myaccount.google.com/apppasswords>
2. Selecione "Mail" e "Windows Computer"
3. Copie a **senha de app** gerada (não use sua senha normal!)
4. Cole em `password =` no arquivo `email.config`

### Passo 3: Usar o Sistema

1. Execute o Conciliador NFS-e GUI
2. Preencha os campos:
   - 📁 **Domínio Sistemas**: Arquivo de dados do domínio
   - 📁 **Portal Nacional**: Arquivo de dados do portal
   - 📁 **Pasta de Saída**: Onde salvar resultados
3. ✅ **MARQUE** a caixa "📧 Enviar Alerta por Email"
4. Clique "🔄 Iniciar Conciliação"
5. Se houver divergências → Email é enviado automaticamente para todos os destinatários

## 📧 O Que Você Receberá

Um email formatado com:

- ✅ Resumo das divergências encontradas
- 📊 Tabela com valores do Domínio vs Portal
- 💰 Total de diferenças (FURO)
- ⏱️ Tempo de execução
- 📎 Caminho do relatório Excel

**Exemplo de email:**

```
ALERTA: Reconciliação NFS-e concluída com 5 divergências!

┌─────────────────────────────────────┐
│ Notas com Divergência: 5            │
│ Total Domínio: R$ 10.500,00         │
│ Total Portal: R$ 9.800,00           │
│ FURO (Diferença): R$ 700,00         │
│ Tempo: 2.34 segundos                │
└─────────────────────────────────────┘

Arquivo: C:\Saída\conciliacao_20240815.xlsx
```

## 🔐 Segurança

- A senha é **NUNCA** armazenada no código
- Apenas no arquivo `email.config` local (não incluído no Git)
- Use **Senha de App** (Gmail) ou **Senha Específica do Aplicativo** (Outlook)
- Nunca use sua senha de conta normal!

## ⚙️ Configurações Para Outros Provedores

### Outlook/Hotmail

```ini
smtp_host = smtp-mail.outlook.com
smtp_port = 587
```

### Yahoo Mail

```ini
smtp_host = smtp.mail.yahoo.com
smtp_port = 587
```

### Office 365

```ini
smtp_host = smtp.office365.com
smtp_port = 587
```

## 🛑 Problemas Comuns

| Problema | Solução |
| ---------- | --------- |
| Email não recebido | Verifique spam/lixo; confira destinatários em email.config |
| Erro "Authentication failed" | Gere nova Senha de App; verifique username correto |
| Arquivo email.config não encontrado | Crie na pasta do conciliador_nfse.py |
| Nenhum email quando tem divergências | Verifique se checkbox está marcado |

## 📝 Estrutura de email.config

```ini
[EMAIL]
# Servidor SMTP do seu provedor de email
smtp_host = smtp.gmail.com

# Porta (geralmente 587 para TLS)
smtp_port = 587

# Email do remetente (seu email)
username = seu_email@gmail.com

# Senha de App (NÃO sua senha normal)
password = abcd efgh ijkl mnop

# Lista de destinatários (separados por vírgula)
destinatarios = financeiro@empresa.com,gerente@empresa.com
```

## ✨ Características

✅ Envia automaticamente ao detectar divergências  
✅ Não envia se reconciliação perfeita (sem divergências)  
✅ Formato HTML com tabela colorida  
✅ Múltiplos destinatários suportados  
✅ Sincronizado com progresso da reconciliação  
✅ Logging detalhado de todos os envios  
✅ Funciona em background (não bloqueia a GUI)  

## 🧪 Testes

Execute para verificar que tudo funciona:

```bash
python test_melhorias.py
```

Espere pelos testes de email:

- ✅ test_config_email() - Verifica carregamento de configuração
- ✅ test_enviar_alerta_email() - Testa lógica de envio

---

**Versão**: 2.0 com Alertas por Email  
**Data**: Agosto 2024  
**Status**: ✅ Pronto para Produção
