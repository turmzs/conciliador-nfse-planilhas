# 📊 Conciliador Automático de NFS-e (Domínio Sistemas vs Portal Nacional)

Uma ferramenta de automação fiscal em Python focada em resolver o cruzamento (conciliação) de Notas Fiscais de Serviço Eletrônicas (NFS-e) exportadas do ERP **Domínio Sistemas** contra o relatório oficial do **Portal Nacional da NFS-e**. 

O processo, que antes exigia auditoria visual ("no olho") ou planilhas complexas, agora é determinístico, gerando um relatório analítico limpo em questão de segundos, destacando apenas as divergências, furos de valor e notas faltantes (órfãs).

## 🚀 Funcionalidades e Regras de Negócio

*   **🔑 Chave de Busca Inteligente (Match Único):** O cruzamento principal une o `Número da NFS-e` + `CNPJ do Prestador/Tomador`. O algoritmo varre sujeiras como formatações de pontuação (pontos, traços e barras) no CNPJ e formatações flutuantes (ex: `1234.0` transformado em `1234`) para garantir 100% de precisão no cruzamento.
*   **🗑️ Purificação de Dados:** O sistema automaticamente ignora as linhas vazias ou as linhas de "Total" geradas no fim das planilhas que causavam falsos positivos ou dobravam as somatórias de valor.
*   **💰 Auditoria de Valores:** Exige uma diferença de **Zero Absoluto** (`0.00`) na comparação do Valor Total. A ferramenta também está pronta para comparar impostos retidos (INSS, IRRF, CSLL, PIS, COFINS, ISS) assim que eles forem exportados na planilha da Domínio. O sistema trata formatos monetários brasileiros e americanos dinamicamente (ex: `1.500,00` -> `1500.00`).
*   **🔎 Notas Faltantes (Órfãs):** Aponta diretamente notas cadastradas no Governo que esqueceram de subir no ERP, ou notas do ERP que não estão validadas no Portal.
*   **🚨 Status de Cancelamento:** Isola e sinaliza gravemente notas que estão Ativas na Domínio, mas marcadas como Canceladas no Portal Nacional.

## 🖥️ Como Instalar e Rodar

### Pré-requisitos
A ferramenta foi desenvolvida em Python. É necessário ter o Python instalado e as seguintes bibliotecas de manipulação de dados:

```bash
pip install pandas numpy openpyxl xlsxwriter odfpy xlrd
```

*Nota: As bibliotecas `odfpy` e `xlrd` garantem que o sistema suporte planilhas antigas (`.xls`) e formatos do LibreOffice (`.ods`), além dos `.xlsx` padrão.*

### Executando a Ferramenta

1. Abra o terminal e rode o arquivo Python:
```bash
python conciliador_nfse.py
```
2. Uma **Interface Gráfica (GUI)** amigável será aberta.
3. Clique em **Procurar** e selecione a exportação do ERP Domínio (ex: `Entradas.xls` ou `.ods`).
4. Clique em **Procurar** e selecione a exportação do Portal Nacional.
5. Clique em **Salvar Como** e defina o nome do seu relatório final de auditoria.
6. Clique em **Executar Conciliação** e aguarde a confirmação de Sucesso! Acompanhe o log detalhado no terminal.

## 📑 O Relatório de Saída (Dashboard)
Ao finalizar, a ferramenta entregará uma nova planilha Excel dividida cirurgicamente em Abas:

*   **📊 Resumo Geral:** Um painel consolidado com a quantidade de notas, valores totais somados e a Diferença (furo) entre as duas planilhas. Ideal para auditoria rápida.
*   **🚨 Urgente:** Notas canceladas no Portal, mas ainda ativas no ERP.
*   **⚠️ Faltantes na Domínio:** Notas registradas no governo, mas não lançadas no seu ERP.
*   **⚠️ Faltantes no Portal:** Notas no seu ERP que não constam na base do governo.
*   **🔍 Divergência Valores:** Notas que deram "match", mas possuem diferença de centavos, valores digitados errados ou datas de emissão incorretas.
*   **✅ Tudo Certo:** Histórico de auditoria, contendo todas as notas que bateram 100%.

## ⚙️ Configuração para Novos Cenários
Se no futuro o layout da Domínio Sistemas mudar (ou se você conseguir exportar mais colunas, como impostos retidos), basta abrir o código `conciliador_nfse.py` no Bloco de Notas/VS Code e ajustar as variáveis dentro da seção `1. CONFIGURAÇÃO DAS COLUNAS`.

---
*Feito para tornar a rotina fiscal rápida, exata e sem estresse.*
*Criado e pensado por Artur Menezes - turmzs*