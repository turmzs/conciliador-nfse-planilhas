import pandas as pd
import numpy as np
import tkinter as tk
from tkinter import filedialog, messagebox
import os

def conciliar_nfse(caminho_dominio, caminho_portal, caminho_saida):
    """
    Função principal para realizar a conciliação entre o relatório da Domínio Sistemas e o Portal Nacional.
    """
    
    # ==============================================================================================
    # 1. LEITURA DOS DADOS E TRATAMENTO DE EXTENSÃO
    # ==============================================================================================
    def read_file(filepath):
        filepath_str = str(filepath).lower()
        if filepath_str.endswith('.ods'):
            engine = 'odf'
        elif filepath_str.endswith('.xls'):
            engine = 'xlrd'
        elif filepath_str.endswith('.xlsx'):
            engine = 'openpyxl'
        else:
            engine = None # Deixa o pandas tentar adivinhar
            
        try:
            return pd.read_excel(filepath, engine=engine)
        except ValueError as e:
            if "Excel file format cannot be determined" in str(e):
                try:
                    return pd.read_csv(filepath, sep=';', encoding='latin-1')
                except:
                    pass
            raise e

    print("\n=======================================================")
    print("[DEBUG] INICIANDO CONCILIAÇÃO")
    print("=======================================================")
    print(f"[*] Lendo Domínio: {caminho_dominio}")
    df_dom = read_file(caminho_dominio)
    print(f"[*] Lendo Portal: {caminho_portal}")
    df_portal = read_file(caminho_portal)

    # ==============================================================================================
    # 2. DESCOBERTA DINÂMICA DE COLUNAS (UNIVERSAL)
    # ==============================================================================================
    def get_column(df, possible_names, required=False):
        """Encontra a primeira coluna correspondente que tenha dados reais. Pula colunas vazias."""
        first_match = None
        for name in possible_names:
            for col in df.columns:
                if str(col).strip().lower() == str(name).lower():
                    non_null = int(df[col].notna().sum())
                    if non_null > 0:
                        print(f"[DEBUG] Coluna mapeada: '{col}' ({non_null} linhas com dados)")
                        return col
                    if first_match is None:
                        first_match = col
        
        if first_match:
            print(f"[DEBUG] Coluna mapeada (sem dados): '{first_match}'")
            return first_match
        
        if required:
            raise KeyError(f"Coluna obrigatória não encontrada.\nProcurado por: {possible_names}\nDisponíveis no arquivo: {list(df.columns)}")
        return None

    # Domínio - Mapeamento flexível (cobre tomados E prestados)
    COL_NUM_NF_DOM = get_column(df_dom, ['NUMERAÇÃO NOTA', 'documento_inicial', 'numero_nota', 'nota', 'numero'], required=True)
    COL_CNPJ_DOM = get_column(df_dom, ['cnpj_fornecedor', 'cnpj_cliente', 'cnpj_empresa', 'cnpj', 'cpf/cnpj'], required=True)
    COL_DOM_DATA = get_column(df_dom, ['DATA EMISSAO', 'data_emissao', 'data_servico', 'data_entrada'])
    COL_DOM_VALOR_LIQ = get_column(df_dom, ['VALOR DA NOTA', 'valor_contabil', 'valor_liquido', 'valor'], required=True)
    COL_DOM_STATUS = get_column(df_dom, ['situacao', 'status'])
    COL_DOM_VALOR_BRUTO = None
    COL_DOM_INSS = get_column(df_dom, ['valor_inss', 'retencao_inss'])
    COL_DOM_IRRF = get_column(df_dom, ['valor_irrf', 'retencao_irrf'])
    COL_DOM_CSLL = get_column(df_dom, ['valor_csll', 'retencao_csll'])
    COL_DOM_PIS = get_column(df_dom, ['valor_pis', 'retencao_pis'])
    COL_DOM_COFINS = get_column(df_dom, ['valor_cofins', 'retencao_cofins'])
    COL_DOM_ISS = get_column(df_dom, ['valor_iss', 'retencao_iss', 'valor_imposto', 'iss'])

    # Portal Nacional - Mapeamento flexível (cobre tomados E prestados)
    COL_NUM_NF_PORTAL = get_column(df_portal, ['Número NFS-e', 'Nmero NFS-e', 'Numero'], required=True)
    COL_CNPJ_PORTAL = get_column(df_portal, ['CNPJ/CPF Prestador', 'CNPJ Prestador', 'CNPJ/CPF Tomador'], required=True)
    COL_PORTAL_DATA = get_column(df_portal, ['Data Geração', 'Data Gerao', 'Competência', 'Competncia'])
    COL_PORTAL_VALOR_LIQ = get_column(df_portal, ['Valor do Serviço (R$)', 'Valor do Servio (R$)', 'Valor'], required=True)
    COL_PORTAL_STATUS = get_column(df_portal, ['Situação NFS-e', 'Situao NFS-e', 'Situação'])
    COL_PORTAL_VALOR_BRUTO = None
    COL_PORTAL_INSS = get_column(df_portal, ['Contrib. Previd. Ret. (R$)', 'INSS'])
    COL_PORTAL_IRRF = get_column(df_portal, ['IRRF (R$)'])
    COL_PORTAL_CSLL = None
    COL_PORTAL_PIS = get_column(df_portal, ['PIS - Débito (R$)', 'PIS - Dbito (R$)'])
    COL_PORTAL_COFINS = get_column(df_portal, ['COFINS - Débito (R$)', 'COFINS - Dbito (R$)'])
    COL_PORTAL_ISS = get_column(df_portal, ['Valor do ISSQN (R$)', 'ISSQN'])

    # Termos usados para identificar nota cancelada
    STATUS_CANCELADO_PORTAL = ['CANCELADA', 'CANCELADO', 'C']
    STATUS_ATIVO_DOMINIO = ['NORMAL', 'ATIVA', 'N', 'A']
    
    # Limpar linhas nulas ou de "Total" que vêm no final dos relatórios
    df_dom = df_dom.dropna(subset=[COL_NUM_NF_DOM])
    df_portal = df_portal.dropna(subset=[COL_NUM_NF_PORTAL])
    
    df_dom = df_dom[df_dom[COL_NUM_NF_DOM].astype(str).str.contains(r'\d', regex=True)].copy()
    df_portal = df_portal[df_portal[COL_NUM_NF_PORTAL].astype(str).str.contains(r'\d', regex=True)].copy()

    print(f"[DEBUG] Total de notas reais lidas - Domínio: {len(df_dom)} | Portal: {len(df_portal)}")

    # ==============================================================================================
    # 3. TRATAMENTO DE CHAVES DE BUSCA (EVITAR FALSOS NEGATIVOS)
    # ==============================================================================================
    def limpar_chave(serie):
        # Converte para string, tira espaços, tira .0 se o pandas leu como float, e remove não-números (ex: pontuação de CNPJ)
        s = serie.astype(str).str.strip()
        s = s.str.replace(r'\.0$', '', regex=True)
        s = s.str.replace(r'\D', '', regex=True) # \D = tudo que não é dígito
        return s

    chave_num_dom = limpar_chave(df_dom[COL_NUM_NF_DOM])
    chave_cnpj_dom = limpar_chave(df_dom[COL_CNPJ_DOM])
    df_dom['Chave_Busca'] = chave_num_dom + "_" + chave_cnpj_dom

    chave_num_portal = limpar_chave(df_portal[COL_NUM_NF_PORTAL])
    chave_cnpj_portal = limpar_chave(df_portal[COL_CNPJ_PORTAL])
    df_portal['Chave_Busca'] = chave_num_portal + "_" + chave_cnpj_portal

    # Função auxiliar para limpar valores financeiros sujos (strings com vírgula)
    def parse_valor(val):
        if pd.isna(val) or val == '': return 0.0
        if isinstance(val, (int, float)): return float(val)
        val = str(val).strip()
        if not val or val.upper() in ('NAN', 'NONE', 'NULL'): return 0.0
        # Tenta resolver o formato brasileiro 1.500,00 -> 1500.00
        if ',' in val and '.' in val:
            if val.rfind(',') > val.rfind('.'):
                val = val.replace('.', '').replace(',', '.')
            else:
                val = val.replace(',', '')
        elif ',' in val:
            val = val.replace(',', '.')
        try:
            return float(val)
        except ValueError:
            return 0.0

    # ==============================================================================================
    # 4. IDENTIFICAÇÃO DE NOTAS FALTANTES (ÓRFÃS)
    # ==============================================================================================
    chaves_dom = set(df_dom['Chave_Busca'])
    chaves_portal = set(df_portal['Chave_Busca'])

    chaves_faltantes_dom = chaves_portal - chaves_dom
    chaves_faltantes_portal = chaves_dom - chaves_portal

    print(f"[DEBUG] Faltam na Domínio (Tem no Portal): {len(chaves_faltantes_dom)}")
    print(f"[DEBUG] Faltam no Portal (Tem na Domínio): {len(chaves_faltantes_portal)}")

    df_faltantes_dominio = df_portal[df_portal['Chave_Busca'].isin(chaves_faltantes_dom)].copy()
    df_faltantes_portal = df_dom[df_dom['Chave_Busca'].isin(chaves_faltantes_portal)].copy()

    # ==============================================================================================
    # 5. CRUZAMENTO DE DADOS (MATCH)
    # ==============================================================================================
    chaves_comuns = chaves_dom.intersection(chaves_portal)
    
    df_dom_match = df_dom[df_dom['Chave_Busca'].isin(chaves_comuns)].set_index('Chave_Busca')
    df_portal_match = df_portal[df_portal['Chave_Busca'].isin(chaves_comuns)].set_index('Chave_Busca')

    lista_urgente = []
    lista_divergencia = []
    lista_tudo_certo = []

    # Apenas compara os valores financeiros se ambas as planilhas tiverem as colunas configuradas (diferentes de None)
    mapeamentos_financeiros = [
        (COL_DOM_VALOR_LIQ, COL_PORTAL_VALOR_LIQ, 'Valor Líquido'),
        (COL_DOM_VALOR_BRUTO, COL_PORTAL_VALOR_BRUTO, 'Valor Bruto'),
        (COL_DOM_INSS, COL_PORTAL_INSS, 'INSS'),
        (COL_DOM_IRRF, COL_PORTAL_IRRF, 'IRRF'),
        (COL_DOM_CSLL, COL_PORTAL_CSLL, 'CSLL'),
        (COL_DOM_PIS, COL_PORTAL_PIS, 'PIS'),
        (COL_DOM_COFINS, COL_PORTAL_COFINS, 'COFINS'),
        (COL_DOM_ISS, COL_PORTAL_ISS, 'ISS'),
    ]
    colunas_comparar = [(d, p, n) for d, p, n in mapeamentos_financeiros if d and p]

    for chave in chaves_comuns:
        row_dom = df_dom_match.loc[chave]
        row_portal = df_portal_match.loc[chave]
        
        # Tratamento caso haja duplicatas de chave primária
        if isinstance(row_dom, pd.DataFrame):
            row_dom = row_dom.iloc[0]
        if isinstance(row_portal, pd.DataFrame):
            row_portal = row_portal.iloc[0]

        divergencias_encontradas = []

        # 5.1 Auditoria de Status (Cancelamentos)
        status_dom = "N/A"
        status_portal = "N/A"
        
        if COL_DOM_STATUS and COL_PORTAL_STATUS:
            status_portal = str(row_portal.get(COL_PORTAL_STATUS, '')).strip().upper()
            status_dom = str(row_dom.get(COL_DOM_STATUS, '')).strip().upper()

            if status_portal in STATUS_CANCELADO_PORTAL and status_dom in STATUS_ATIVO_DOMINIO:
                lista_urgente.append({
                    'Chave': chave,
                    'Nota': row_dom[COL_NUM_NF_DOM],
                    'CNPJ': row_dom[COL_CNPJ_DOM],
                    'Status_Portal': status_portal,
                    'Status_Dominio': status_dom,
                    'Motivo': 'Cancelada no Portal, mas Ativa na Domínio'
                })
                continue # Pula validações de valor

        # 5.2 Auditoria de Datas
        if COL_DOM_DATA and COL_PORTAL_DATA:
            data_dom = pd.to_datetime(row_dom.get(COL_DOM_DATA), dayfirst=True, errors='coerce')
            data_portal = pd.to_datetime(row_portal.get(COL_PORTAL_DATA), dayfirst=True, errors='coerce')
            
            if pd.notnull(data_dom) and pd.notnull(data_portal) and data_dom != data_portal:
                divergencias_encontradas.append(f"Data Domínio: {data_dom.date()} | Data Portal: {data_portal.date()}")

        # 5.3 Auditoria de Valores
        for col_d, col_p, nome_legivel in colunas_comparar:
            val_d = round(parse_valor(row_dom.get(col_d, 0)), 2)
            val_p = round(parse_valor(row_portal.get(col_p, 0)), 2)
            
            # Exigindo zero absoluto na diferença
            if val_d != val_p: 
                divergencias_encontradas.append(f"{nome_legivel}: Domínio {val_d} x Portal {val_p}")

        # Classificação final do match
        registro_base = {
            'Chave': chave,
            'Nota': row_dom[COL_NUM_NF_DOM],
            'CNPJ': row_dom[COL_CNPJ_DOM],
            'Status_Dominio': status_dom if COL_DOM_STATUS else "Sem Coluna de Status",
            'Status_Portal': status_portal if COL_PORTAL_STATUS else "Sem Coluna de Status"
        }

        if divergencias_encontradas:
            registro_div = registro_base.copy()
            registro_div['Detalhes_Divergencia'] = " || ".join(divergencias_encontradas)
            lista_divergencia.append(registro_div)
        else:
            lista_tudo_certo.append(registro_base)

    # Convertendo listas para DataFrames
    df_urgente = pd.DataFrame(lista_urgente)
    df_divergencia = pd.DataFrame(lista_divergencia)
    df_tudo_certo = pd.DataFrame(lista_tudo_certo)

    # ==============================================================================================
    # 6. GERAÇÃO DO RELATÓRIO DE SAÍDA E RESUMO
    # ==============================================================================================
    
    # Criar um resumo geral
    soma_valor_dom = sum(parse_valor(v) for v in df_dom[COL_DOM_VALOR_LIQ]) if COL_DOM_VALOR_LIQ else 0
    soma_valor_portal = sum(parse_valor(v) for v in df_portal[COL_PORTAL_VALOR_LIQ]) if COL_PORTAL_VALOR_LIQ else 0
    diferenca_total = soma_valor_dom - soma_valor_portal

    print(f"[DEBUG] Notas com divergência de valores: {len(df_divergencia)}")
    print(f"[DEBUG] Soma Valor Domínio: R$ {soma_valor_dom:.2f}")
    print(f"[DEBUG] Soma Valor Portal: R$ {soma_valor_portal:.2f}")
    print(f"[DEBUG] DIFERENÇA (FURO): R$ {diferenca_total:.2f}")
    print("=======================================================\n")

    resumo = pd.DataFrame([
        {'Métrica': 'Total de Notas na Domínio', 'Valor': len(df_dom)},
        {'Métrica': 'Total de Notas no Portal', 'Valor': len(df_portal)},
        {'Métrica': 'Notas Faltantes na Domínio', 'Valor': len(df_faltantes_dominio)},
        {'Métrica': 'Notas Faltantes no Portal', 'Valor': len(df_faltantes_portal)},
        {'Métrica': 'Notas com Divergência de Valor', 'Valor': len(df_divergencia)},
        {'Métrica': 'Valor Total - Domínio (R$)', 'Valor': round(soma_valor_dom, 2)},
        {'Métrica': 'Valor Total - Portal (R$)', 'Valor': round(soma_valor_portal, 2)},
        {'Métrica': 'DIFERENÇA DE VALOR (Furo)', 'Valor': round(diferenca_total, 2)}
    ])

    with pd.ExcelWriter(caminho_saida, engine='xlsxwriter') as writer:
        resumo.to_excel(writer, sheet_name='📊 Resumo Geral', index=False)
        
        if not df_urgente.empty:
            df_urgente.to_excel(writer, sheet_name='🚨 Urgente', index=False)
        else:
            pd.DataFrame({'Mensagem': ['Nenhuma nota crítica encontrada']}).to_excel(writer, sheet_name='🚨 Urgente', index=False)

        if not df_faltantes_dominio.empty:
            df_faltantes_dominio.to_excel(writer, sheet_name='⚠️ Faltantes na Domínio', index=False)
        
        if not df_faltantes_portal.empty:
            df_faltantes_portal.to_excel(writer, sheet_name='⚠️ Faltantes no Portal', index=False)
            
        if not df_divergencia.empty:
            df_divergencia.to_excel(writer, sheet_name='🔍 Divergência Valores', index=False)
            
        if not df_tudo_certo.empty:
            df_tudo_certo.to_excel(writer, sheet_name='✅ Tudo Certo', index=False)


# ==============================================================================================
# INTERFACE GRÁFICA (GUI)
# ==============================================================================================
class ConciliadorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Conciliador Automático de NFS-e")
        self.root.geometry("550x320")
        self.root.resizable(False, False)
        
        self.caminho_dominio = tk.StringVar()
        self.caminho_portal = tk.StringVar()
        self.caminho_saida = tk.StringVar()
        
        tk.Label(root, text="Conciliação Domínio Sistemas vs Portal NFS-e", font=("Arial", 14, "bold")).pack(pady=10)
        
        # Seleção Domínio
        tk.Label(root, text="Planilha Domínio Sistemas (.xlsx, .xls, .ods):", font=("Arial", 10)).pack(anchor="w", padx=20)
        frame_dom = tk.Frame(root)
        frame_dom.pack(fill="x", padx=20, pady=5)
        tk.Entry(frame_dom, textvariable=self.caminho_dominio, state='readonly').pack(side=tk.LEFT, fill="x", expand=True, padx=(0, 10))
        tk.Button(frame_dom, text="Procurar", command=self.selecionar_dominio, width=10).pack(side=tk.RIGHT)
        
        # Seleção Portal
        tk.Label(root, text="Planilha Portal Nacional (.xlsx, .xls, .ods):", font=("Arial", 10)).pack(anchor="w", padx=20)
        frame_portal = tk.Frame(root)
        frame_portal.pack(fill="x", padx=20, pady=5)
        tk.Entry(frame_portal, textvariable=self.caminho_portal, state='readonly').pack(side=tk.LEFT, fill="x", expand=True, padx=(0, 10))
        tk.Button(frame_portal, text="Procurar", command=self.selecionar_portal, width=10).pack(side=tk.RIGHT)
        
        # Seleção Saída
        tk.Label(root, text="Onde salvar o Relatório de Conciliação (.xlsx):", font=("Arial", 10)).pack(anchor="w", padx=20)
        frame_saida = tk.Frame(root)
        frame_saida.pack(fill="x", padx=20, pady=5)
        tk.Entry(frame_saida, textvariable=self.caminho_saida, state='readonly').pack(side=tk.LEFT, fill="x", expand=True, padx=(0, 10))
        tk.Button(frame_saida, text="Salvar Como", command=self.selecionar_saida, width=10).pack(side=tk.RIGHT)
        
        tk.Button(root, text="▶ Executar Conciliação", bg="#4CAF50", fg="white", font=("Arial", 12, "bold"), command=self.executar).pack(pady=20)

    def selecionar_dominio(self):
        arquivo = filedialog.askopenfilename(title="Selecione a planilha da Domínio", filetypes=[("Arquivos Excel/ODS", "*.xlsx *.xls *.ods")])
        if arquivo:
            self.caminho_dominio.set(arquivo)

    def selecionar_portal(self):
        arquivo = filedialog.askopenfilename(title="Selecione a planilha do Portal", filetypes=[("Arquivos Excel/ODS", "*.xlsx *.xls *.ods")])
        if arquivo:
            self.caminho_portal.set(arquivo)

    def selecionar_saida(self):
        arquivo = filedialog.asksaveasfilename(title="Salvar relatório final", defaultextension=".xlsx", filetypes=[("Arquivo Excel", "*.xlsx")])
        if arquivo:
            self.caminho_saida.set(arquivo)

    def executar(self):
        if not self.caminho_dominio.get() or not self.caminho_portal.get() or not self.caminho_saida.get():
            messagebox.showwarning("Aviso", "Por favor, preencha todos os campos.")
            return
            
        try:
            conciliar_nfse(self.caminho_dominio.get(), self.caminho_portal.get(), self.caminho_saida.get())
            messagebox.showinfo("Sucesso", f"Conciliação concluída!\n\nRelatório salvo em:\n{self.caminho_saida.get()}")
        except KeyError as e:
            messagebox.showerror("Erro de Coluna", f"Erro: {str(e)}\n\nO layout da planilha mudou ou as colunas esperadas não estão presentes.")
        except Exception as e:
            messagebox.showerror("Erro Inesperado", f"Ocorreu um erro durante a conciliação:\n{str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = ConciliadorGUI(root)
    root.mainloop()
