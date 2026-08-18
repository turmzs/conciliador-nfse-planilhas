import pandas as pd
import numpy as np
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import threading
import logging
from typing import Optional, Dict, List, Tuple
import re
import unicodedata
import hashlib
import json
from datetime import datetime
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from configparser import ConfigParser

# Configurar logging
logging.basicConfig(
    level=logging.INFO,
    format='[%(levelname)s] %(message)s'
)
logger = logging.getLogger(__name__)

def remover_acentos(texto: str) -> str:
    """Remove acentuações de uma string para comparação robusta."""
    if not texto:
        return ''
    nfd = unicodedata.normalize('NFD', texto)
    return ''.join(char for char in nfd if unicodedata.category(char) != 'Mn')

def gerar_hash_arquivo(caminho: str) -> str:
    """Gera hash SHA256 de um arquivo para validação."""
    sha256_hash = hashlib.sha256()
    try:
        with open(caminho, 'rb') as f:
            for byte_block in iter(lambda: f.read(4096), b''):
                sha256_hash.update(byte_block)
        return sha256_hash.hexdigest()
    except Exception as e:
        logger.warning(f"Não foi possível gerar hash do arquivo {caminho}: {e}")
        return "erro"

def validar_dados(df: pd.DataFrame, col_num_nota: str, col_cnpj: str, col_valor: str) -> Dict[str, any]:
    """Valida dados da planilha e retorna relatório de problemas encontrados.
    
    Valida:
    - Linhas duplicadas (mesmo CNPJ + numero_nota)
    - CNPJ inválido (dígitos verificadores errados)
    - Valores negativos ou zero
    - CNPJs vazios obrigatoriamente
    
    Returns:
        Dict com 'ok' (bool), 'avisos' (list), 'erros' (list), 'total_problemas' (int)
    """
    avisos = []
    erros = []
    
    # 1. CNPJs vazios
    cnpjs_vazios = df[col_cnpj].isna().sum() + df[col_cnpj].astype(str).str.strip().eq('').sum()
    if cnpjs_vazios > 0:
        erros.append(f"CNPJ vazio/nulo em {cnpjs_vazios} linhas (obrigatório)")
    
    # 2. Números de nota vazios
    notas_vazias = df[col_num_nota].isna().sum() + df[col_num_nota].astype(str).str.strip().eq('').sum()
    if notas_vazias > 0:
        erros.append(f"Número de nota vazio/nulo em {notas_vazias} linhas (obrigatório)")
    
    # 3. Linhas duplicadas (CNPJ + numero_nota)
    def limpar_chave_temp(serie):
        s = serie.astype(str).str.strip()
        s = s.str.replace(r'\.0$', '', regex=True)
        s = s.str.replace(r'\D', '', regex=True)
        return s
    
    chave_dup = limpar_chave_temp(df[col_num_nota]) + "_" + limpar_chave_temp(df[col_cnpj])
    duplicatas = chave_dup[chave_dup.duplicated(keep=False)]
    if len(duplicatas) > 0:
        n_duplic = len(duplicatas[duplicatas.duplicated(keep='first')])
        avisos.append(f"Encontradas {n_duplic} linhas duplicadas (mesmo CNPJ + numero_nota)")
    
    # 4. Valores negativos ou zero (AVISOS, não erros)
    valores_invalidos = 0
    if col_valor:
        try:
            # Precisa de uma função auxiliar para converter
            from functools import partial
            vals = pd.Series([parse_valor(v) for v in df[col_valor]])
            valores_zero = (vals <= 0).sum()
            if valores_zero > 0:
                avisos.append(f"{valores_zero} linhas com valor <= 0 (zero ou negativo)")
                valores_invalidos = valores_zero
        except Exception as e:
            avisos.append(f"Erro ao validar valores: {str(e)}")
    
    # 5. CNPJ inválido (valida dígitos verificadores)
    def validar_cnpj(cnpj_str: str) -> bool:
        """Valida CNPJ usando dígitos verificadores."""
        cnpj = ''.join(filter(str.isdigit, str(cnpj_str)))
        if len(cnpj) != 14:
            return False
        try:
            n = int(cnpj)
            # Validar dígitos verificadores (simplificado)
            # Primeiro dígito verificador
            calc = sum(int(cnpj[i]) * (5 - (i % 8)) for i in range(12))
            dv1 = 11 - (calc % 11)
            dv1 = 0 if dv1 > 9 else dv1
            if int(cnpj[12]) != dv1:
                return False
            # Segundo dígito verificador
            calc = sum(int(cnpj[i]) * (6 - (i % 8)) for i in range(13))
            dv2 = 11 - (calc % 11)
            dv2 = 0 if dv2 > 9 else dv2
            if int(cnpj[13]) != dv2:
                return False
            return True
        except:
            return False
    
    cnpjs_invalidos = 0
    for cnpj in df[col_cnpj].dropna():
        if str(cnpj).strip() and not validar_cnpj(str(cnpj)):
            cnpjs_invalidos += 1
    
    if cnpjs_invalidos > 0:
        avisos.append(f"{cnpjs_invalidos} CNPJ(s) com dígitos verificadores inválidos")
    
    ok = len(erros) == 0
    total_problemas = len(erros) + len(avisos)
    
    resultado = {
        'ok': ok,
        'avisos': avisos,
        'erros': erros,
        'total_problemas': total_problemas,
        'linhas_duplicadas': len(duplicatas) if 'duplicatas' in locals() else 0,
        'valores_invalidos': valores_invalidos,
        'cnpjs_invalidos': cnpjs_invalidos
    }
    
    return resultado

def salvar_historico(caminho_saida: str, caminho_dominio: str, caminho_portal: str, 
                     tempo_execucao: float, num_notas_dom: int, num_notas_portal: int,
                     num_divergencias: int, num_faltantes_dom: int, num_faltantes_portal: int,
                     soma_dom: float, soma_portal: float) -> None:
    """Salva metadados da conciliação em arquivo de histórico JSON.
    
    Cria pasta /histórico/ e salva:
    - arquivo JSON com timestamp, totais, divergências, hashes
    """
    pasta_historico = os.path.join(os.path.dirname(caminho_saida), 'histórico')
    os.makedirs(pasta_historico, exist_ok=True)
    
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    nome_arquivo = f"conciliacao_{timestamp}.json"
    caminho_json = os.path.join(pasta_historico, nome_arquivo)
    
    metadata = {
        'timestamp': datetime.now().isoformat(),
        'arquivo_dominio': os.path.basename(caminho_dominio),
        'arquivo_portal': os.path.basename(caminho_portal),
        'hash_dominio': gerar_hash_arquivo(caminho_dominio),
        'hash_portal': gerar_hash_arquivo(caminho_portal),
        'tempo_execucao_segundos': round(tempo_execucao, 2),
        'notas': {
            'total_dominio': num_notas_dom,
            'total_portal': num_notas_portal,
            'faltantes_dominio': num_faltantes_dom,
            'faltantes_portal': num_faltantes_portal,
            'com_divergencias': num_divergencias
        },
        'valores': {
            'soma_dominio': round(soma_dom, 2),
            'soma_portal': round(soma_portal, 2),
            'diferenca': round(soma_dom - soma_portal, 2)
        }
    }
    
    try:
        with open(caminho_json, 'w', encoding='utf-8') as f:
            json.dump(metadata, f, indent=2, ensure_ascii=False)
        logger.info(f"Histórico salvo em: {caminho_json}")
    except Exception as e:
        logger.warning(f"Erro ao salvar histórico: {e}")

def exportar_divergencias_csv(df_divergencias: pd.DataFrame, caminho_base: str) -> str:
    r"""Exporta divergências em CSV simples para investigação rápida.
    
    Arquivo salvo em: {pasta}\divergencias_{timestamp}.csv
    Colunas: Numero_Nota | Valor_Dominio | Valor_Portal | Diferenca
    """
    if df_divergencias.empty:
        logger.info("Nenhuma divergência para exportar")
        return None
    
    pasta_saida = os.path.dirname(caminho_base)
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    nome_arquivo = f"divergencias_{timestamp}.csv"
    caminho_csv = os.path.join(pasta_saida, nome_arquivo)
    
    # Simplificar dataframe: manter apenas colunas relevantes
    df_export = df_divergencias[['Nota', 'CNPJ', 'Detalhes_Divergencia']].copy() if 'Detalhes_Divergencia' in df_divergencias.columns else df_divergencias.copy()
    
    try:
        df_export.to_csv(caminho_csv, index=False, encoding='utf-8', sep=';')
        logger.info(f"Divergências exportadas para: {caminho_csv}")
        return caminho_csv
    except Exception as e:
        logger.warning(f"Erro ao exportar divergências para CSV: {e}")
        return None

def parse_valor(val) -> float:
    r"""Converte um valor monetário em float com detecção robusta de formato BR/US.
    
    Detecta o formato por regex:
      - Padrão BR: \d+(\.\d{3})*(,\d{1,2})? (ex: 1.500,00 ou 1.234.567,89)
      - Padrão US: \d+([,]\d{3})*(\.\d{1,2})? (ex: 1,500.00 ou 1,234,567.89)
    
    Aceita símbolos como 'R$', espaços e parênteses para negativos.
    """
    if pd.isna(val) or val == '':
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    
    val_original = str(val).strip()
    if not val_original or val_original.upper() in ('NAN', 'NONE', 'NULL'):
        return 0.0
    
    # Detecta negativo em formato contábil: (1.500,00)
    negativo = val_original.startswith('(') and val_original.endswith(')')
    if negativo:
        val_original = val_original[1:-1]
    
    # Remove símbolos como 'R$' e espaços
    val_limpo = re.sub(r'[R$\s]', '', val_original)
    
    # Detecta se há multiplicidade (mais de uma ocorrência) de ponto vs vírgula
    # Usando regex para detectar padrão BR vs US mais robustamente
    # Padrão BR: \d+(\.\d{3})*(,\d{1,2})?
    if re.match(r'^-?\d{1,3}(\.\d{3})*(,\d{1,2})?$', val_limpo):
        # Formato BR: remove pontos (milhares) e troca vírgula por ponto
        numero = val_limpo.replace('.', '').replace(',', '.')
    # Padrão US: \d+([,]\d{3})*(\.\d{1,2})?
    elif re.match(r'^-?\d{1,3}(,\d{3})*(\.\d{1,2})?$', val_limpo):
        # Formato US: remove vírgulas (milhares)
        numero = val_limpo.replace(',', '')
    else:
        # Fallback: último separador define o formato (compatível com versão anterior)
        last_dot = val_limpo.rfind('.')
        last_com = val_limpo.rfind(',')
        if last_dot == -1 and last_com == -1:
            numero = val_limpo
        elif last_com > last_dot:
            # Formato BR
            numero = val_limpo.replace('.', '').replace(',', '.')
        else:
            # Formato US
            numero = val_limpo.replace(',', '')
    
    try:
        resultado = float(numero)
    except ValueError:
        logger.warning(f"Não foi possível converter '{val_original}' para float")
        return 0.0
    
    return -resultado if negativo else resultado

def carregar_config_email() -> Dict[str, str]:
    """Carrega configuração de email do arquivo email.config.
    
    Arquivo esperado em: ./email.config
    Formato:
        [EMAIL]
        smtp_host = smtp.gmail.com
        smtp_port = 587
        username = seu_email@gmail.com
        password = sua_senha_app
        destinatarios = email1@example.com,email2@example.com
    
    Returns:
        Dict com configurações ou dict vazio se arquivo não existir
    """
    config_file = 'email.config'
    config = {}
    
    if not os.path.exists(config_file):
        logger.debug(f"Arquivo de configuração {config_file} não encontrado. Alertas por email desabilitados.")
        return config
    
    try:
        parser = ConfigParser()
        parser.read(config_file, encoding='utf-8')
        
        if parser.has_section('EMAIL'):
            config = {
                'smtp_host': parser.get('EMAIL', 'smtp_host', fallback=None),
                'smtp_port': parser.getint('EMAIL', 'smtp_port', fallback=587),
                'username': parser.get('EMAIL', 'username', fallback=None),
                'password': parser.get('EMAIL', 'password', fallback=None),
                'destinatarios': parser.get('EMAIL', 'destinatarios', fallback=None)
            }
            logger.debug(f"Configuração de email carregada. Destinatários: {config['destinatarios']}")
    except Exception as e:
        logger.warning(f"Erro ao carregar configuração de email: {e}")
    
    return config

def enviar_alerta_email(config: Dict, num_divergencias: int, soma_dom: float, soma_portal: float,
                        diferenca: float, tempo_execucao: float, arquivo_saida: str) -> bool:
    r"""Envia email de alerta quando há divergências detectadas.
    
    Args:
        config: Dict com smtp_host, smtp_port, username, password, destinatarios
        num_divergencias: Número de notas com divergências
        soma_dom: Soma dos valores da Domínio
        soma_portal: Soma dos valores do Portal
        diferenca: Diferença total (soma_dom - soma_portal)
        tempo_execucao: Tempo total em segundos
        arquivo_saida: Caminho do arquivo Excel gerado
    
    Returns:
        True se email foi enviado, False caso contrário
    """
    # Validar configuração
    if not config or not all([config.get('smtp_host'), config.get('username'), 
                              config.get('password'), config.get('destinatarios')]):
        logger.debug("Configuração de email incompleta. Alertas desabilitados.")
        return False
    
    if num_divergencias == 0:
        logger.debug("Nenhuma divergência encontrada. Email de alerta não será enviado.")
        return False
    
    try:
        # Preparar lista de destinatários
        destinatarios = [e.strip() for e in config['destinatarios'].split(',')]
        
        # Criar mensagem
        msg = MIMEMultipart('alternative')
        msg['Subject'] = f"🚨 ALERTA: {num_divergencias} divergência(s) encontrada(s) na Conciliação NFS-e"
        msg['From'] = config['username']
        msg['To'] = ', '.join(destinatarios)
        
        # Corpo do email em HTML
        html = f"""
        <html>
            <head>
                <style>
                    body {{ font-family: Arial, sans-serif; background-color: #f5f5f5; }}
                    .container {{ background-color: white; padding: 20px; border-radius: 8px; margin: 20px; }}
                    .alert {{ background-color: #fff3cd; border-left: 4px solid #ffc107; padding: 15px; margin: 10px 0; }}
                    .error {{ background-color: #f8d7da; border-left: 4px solid #dc3545; padding: 15px; margin: 10px 0; }}
                    .success {{ background-color: #d4edda; border-left: 4px solid #28a745; padding: 15px; margin: 10px 0; }}
                    table {{ width: 100%; border-collapse: collapse; margin: 15px 0; }}
                    th {{ background-color: #007bff; color: white; padding: 10px; text-align: left; }}
                    td {{ padding: 10px; border-bottom: 1px solid #ddd; }}
                    .metric-label {{ font-weight: bold; color: #333; }}
                    .metric-value {{ color: #0056b3; font-weight: bold; }}
                    .negative {{ color: #dc3545; }}
                    .positive {{ color: #28a745; }}
                    .footer {{ margin-top: 20px; font-size: 12px; color: #999; border-top: 1px solid #ddd; padding-top: 10px; }}
                </style>
            </head>
            <body>
                <div class="container">
                    <h2>⚠️ Alerta de Conciliação NFS-e</h2>
                    
                    <div class="error">
                        <strong>Divergências Detectadas!</strong>
                        <p>Foram encontradas <strong>{num_divergencias}</strong> nota(s) com divergências de valores.</p>
                    </div>
                    
                    <h3>📊 Resumo da Conciliação</h3>
                    <table>
                        <tr>
                            <td class="metric-label">Soma Domínio Sistemas:</td>
                            <td class="metric-value">R$ {soma_dom:,.2f}</td>
                        </tr>
                        <tr>
                            <td class="metric-label">Soma Portal Nacional:</td>
                            <td class="metric-value">R$ {soma_portal:,.2f}</td>
                        </tr>
                        <tr>
                            <td class="metric-label">Diferença (Furo):</td>
                            <td class="metric-value negative" if diferenca < 0 else class="metric-value positive">R$ {diferenca:,.2f}</td>
                        </tr>
                        <tr>
                            <td class="metric-label">Tempo de Execução:</td>
                            <td class="metric-value">{tempo_execucao:.2f}s</td>
                        </tr>
                    </table>
                    
                    <div class="success">
                        <p>✅ Relatório completo disponível em: <br/><code>{arquivo_saida}</code></p>
                        <p>Abra o arquivo Excel para investigar as divergências em detalhes.</p>
                    </div>
                    
                    <div class="footer">
                        <p>Este é um alerta automático do Conciliador Automático de NFS-e.</p>
                        <p>Data: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}</p>
                    </div>
                </div>
            </body>
        </html>
        """
        
        # Versão em texto simples
        texto = f"""
ALERTA DE CONCILIAÇÃO NFS-e
{'='*50}

DIVERGÊNCIAS DETECTADAS!
Foram encontradas {num_divergencias} nota(s) com divergências.

RESUMO:
  • Soma Domínio Sistemas: R$ {soma_dom:,.2f}
  • Soma Portal Nacional: R$ {soma_portal:,.2f}
  • Diferença (Furo): R$ {diferenca:,.2f}
  • Tempo de Execução: {tempo_execucao:.2f}s

Relatório salvo em: {arquivo_saida}

Data: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}
"""
        
        part1 = MIMEText(texto, 'plain')
        part2 = MIMEText(html, 'html')
        msg.attach(part1)
        msg.attach(part2)
        
        # Enviar email
        logger.info(f"Enviando alerta por email para {len(destinatarios)} destinatário(s)...")
        with smtplib.SMTP(config['smtp_host'], config['smtp_port']) as server:
            server.starttls()  # Usar conexão TLS
            server.login(config['username'], config['password'])
            server.send_message(msg)
        
        logger.info(f"✅ Alerta enviado com sucesso para: {config['destinatarios']}")
        return True
        
    except Exception as e:
        logger.error(f"Erro ao enviar alerta por email: {e}")
        return False

def conciliar_nfse(caminho_dominio: str, caminho_portal: str, caminho_saida: str, progress_callback: Optional[callable] = None, enviar_email: bool = False) -> None:
    """
    Função principal para realizar a conciliação entre o relatório da Domínio Sistemas e o Portal Nacional.
    
    Args:
        caminho_dominio: Caminho do arquivo da Domínio Sistemas
        caminho_portal: Caminho do arquivo do Portal Nacional
        caminho_saida: Caminho de saída do relatório Excel
        progress_callback: Função callback opcional para atualizar progresso (etapa: str, valor: int, máximo: int)
        enviar_email: Se True, envia alerta por email se houver divergências (requer email.config)
    """
    
    def _progress(etapa: str, valor: int = 0, maximo: int = 100):
        """Envia atualização de progresso se callback foi fornecido."""
        if progress_callback:
            progress_callback(etapa, valor, maximo)
        logger.info(f"{etapa}...")
    
    # ==============================================================================================
    # 1. LEITURA DOS DADOS E TRATAMENTO DE EXTENSÃO
    # ==============================================================================================
    # Registrar tempo de início para histórico
    tempo_inicio = datetime.now()
    
    def read_file(filepath: str) -> pd.DataFrame:
        """Lê arquivo Excel ou CSV com múltiplas tentativas de engine/encoding."""
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
            # Se o pandas não identificou o formato, tenta como CSV com várias combinações.
            if "Excel file format cannot be determined" not in str(e):
                raise e
            logger.warning(f"Formato Excel não reconhecido, tentando como CSV: {filepath}")
            tentativas = [
                {'sep': ';', 'encoding': 'latin-1'},
                {'sep': ',', 'encoding': 'utf-8'},
                {'sep': ',', 'encoding': 'latin-1'},
                {'sep': ';', 'encoding': 'utf-8'},
                {'sep': '\t', 'encoding': 'utf-8'},
                {'sep': '\t', 'encoding': 'latin-1'},
            ]
            ultimo_erro = None
            for kw in tentativas:
                try:
                    df = pd.read_csv(filepath, **kw)
                    logger.info(f"CSV lido com sucesso usando: sep={kw['sep']!r}, encoding={kw['encoding']!r}")
                    return df
                except Exception as csv_err:
                    ultimo_erro = csv_err
                    continue
            # Nenhuma combinação funcionou — levanta erro descritivo com todas as tentativas.
            raise ValueError(
                f"Não foi possível ler o arquivo como Excel nem como CSV.\n"
                f"Arquivo: {filepath}\n"
                f"Erro original do Excel: {e}\n"
                f"Último erro do CSV: {ultimo_erro}"
            )

    logger.info("=" * 55)
    logger.info("INICIANDO CONCILIAÇÃO")
    logger.info("=" * 55)
    
    _progress("Lendo Domínio Sistemas", 1, 8)
    df_dom = read_file(caminho_dominio)
    
    _progress("Lendo Portal Nacional", 2, 8)
    df_portal = read_file(caminho_portal)

    # ==============================================================================================
    # 2. DESCOBERTA DINÂMICA DE COLUNAS (UNIVERSAL)
    # ==============================================================================================
    def get_column(df: pd.DataFrame, possible_names: List[str], required: bool = False) -> Optional[str]:
        """Encontra a primeira coluna correspondente que tenha dados reais. Pula colunas vazias.
        Compara nomes ignorando acentuações para maior robustez."""
        first_match = None
        for name in possible_names:
            name_clean = remover_acentos(str(name).strip().lower())
            for col in df.columns:
                col_clean = remover_acentos(str(col).strip().lower())
                if col_clean == name_clean:
                    non_null = int(df[col].notna().sum())
                    if non_null > 0:
                        logger.debug(f"Coluna mapeada: '{col}' ({non_null} linhas com dados)")
                        return col
                    if first_match is None:
                        first_match = col
        
        if first_match:
            logger.debug(f"Coluna mapeada (sem dados): '{first_match}'")
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
    
    # ==============================================================================================
    # 2.5 VALIDAÇÃO DE DADOS
    # ==============================================================================================
    logger.info("Validando integridade dos dados...")
    val_dom = validar_dados(df_dom, COL_NUM_NF_DOM, COL_CNPJ_DOM, COL_DOM_VALOR_LIQ)
    val_portal = validar_dados(df_portal, COL_NUM_NF_PORTAL, COL_CNPJ_PORTAL, COL_PORTAL_VALOR_LIQ)
    
    for aviso in val_dom['avisos']:
        logger.warning(f"[Domínio] {aviso}")
    for aviso in val_portal['avisos']:
        logger.warning(f"[Portal] {aviso}")
    
    if not val_dom['ok']:
        for erro in val_dom['erros']:
            logger.error(f"[Domínio] {erro}")
        raise ValueError(f"Erros de validação na Domínio: {val_dom['erros']}")
    
    if not val_portal['ok']:
        for erro in val_portal['erros']:
            logger.error(f"[Portal] {erro}")
        raise ValueError(f"Erros de validação no Portal: {val_portal['erros']}")
    
    logger.info(f"Validação OK - Domínio ({val_dom['total_problemas']} avisos), Portal ({val_portal['total_problemas']} avisos)")
    
    _progress("Limpando dados", 3, 8)
    # Limpar linhas nulas ou de "Total" que vém no final dos relatórios
    df_dom = df_dom.dropna(subset=[COL_NUM_NF_DOM])
    df_portal = df_portal.dropna(subset=[COL_NUM_NF_PORTAL])
    
    df_dom = df_dom[df_dom[COL_NUM_NF_DOM].astype(str).str.contains(r'\d', regex=True)].copy()
    df_portal = df_portal[df_portal[COL_NUM_NF_PORTAL].astype(str).str.contains(r'\d', regex=True)].copy()

    logger.info(f"Total de notas reais lidas - Domínio: {len(df_dom)} | Portal: {len(df_portal)}")

    # ==============================================================================================
    # 3. TRATAMENTO DE CHAVES DE BUSCA (EVITAR FALSOS NEGATIVOS)
    # ==============================================================================================
    def limpar_chave(serie: pd.Series) -> pd.Series:
        """Remove espaços, .0 de floats e pontuação de CNPJ."""
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

    # ==============================================================================================
    # 4. IDENTIFICAÇÃO DE NOTAS FALTANTES (ÓRFÃS)
    # ==============================================================================================
    _progress("Identificando notas faltantes", 4, 8)
    chaves_dom = set(df_dom['Chave_Busca'])
    chaves_portal = set(df_portal['Chave_Busca'])

    chaves_faltantes_dom = chaves_portal - chaves_dom
    chaves_faltantes_portal = chaves_dom - chaves_portal

    logger.info(f"Faltam na Domínio (Tem no Portal): {len(chaves_faltantes_dom)}")
    logger.info(f"Faltam no Portal (Tem na Domínio): {len(chaves_faltantes_portal)}")

    df_faltantes_dominio = df_portal[df_portal['Chave_Busca'].isin(chaves_faltantes_dom)].copy()
    df_faltantes_portal = df_dom[df_dom['Chave_Busca'].isin(chaves_faltantes_portal)].copy()

    # ==============================================================================================
    # 5. CRUZAMENTO DE DADOS (MATCH) - Versão otimizada via pd.merge (vetorizada)
    # ==============================================================================================
    _progress("Cruzando dados (merge)", 5, 8)
    # Mapeamentos financeiros (filtra apenas os pares com colunas presentes em ambos lados)
    mapeamentos_financeiros = [
        (COL_DOM_VALOR_LIQ, COL_PORTAL_VALOR_LIQ, 'Valor Liquido'),
        (COL_DOM_VALOR_BRUTO, COL_PORTAL_VALOR_BRUTO, 'Valor Bruto'),
        (COL_DOM_INSS, COL_PORTAL_INSS, 'INSS'),
        (COL_DOM_IRRF, COL_PORTAL_IRRF, 'IRRF'),
        (COL_DOM_CSLL, COL_PORTAL_CSLL, 'CSLL'),
        (COL_DOM_PIS, COL_PORTAL_PIS, 'PIS'),
        (COL_DOM_COFINS, COL_PORTAL_COFINS, 'COFINS'),
        (COL_DOM_ISS, COL_PORTAL_ISS, 'ISS'),
    ]
    colunas_comparar = [(d, p, n) for d, p, n in mapeamentos_financeiros if d and p]

    # Merge inner em Chave_Busca. Sufixos _dom/_portal para evitar colisão de nomes.
    # drop_duplicates mantém só a 1ª ocorrência em caso de chaves repetidas (compativel com .iloc[0] do loop original).
    df_dom_uniq = df_dom.drop_duplicates(subset=['Chave_Busca'], keep='first')
    df_portal_uniq = df_portal.drop_duplicates(subset=['Chave_Busca'], keep='first')
    df_match = pd.merge(df_dom_uniq, df_portal_uniq, on='Chave_Busca', suffixes=('_dom', '_portal'))
    logger.info(f"Notas com match por chave: {len(df_match)}")

    # ---------- 5.1 Urgentes (Cancelada no Portal, Ativa na Dominio) ----------
    # O merge adiciona sufixos _dom/_portal apenas em colunas com nomes colidindo.
    # Quando as colunas tem nomes diferentes (caso comum), mantêm o nome original.
    def _col(nome_original, lado):
        """Retorna o nome real da coluna no df_match após o merge."""
        sufixo = '_dom' if lado == 'dom' else '_portal'
        return f'{nome_original}{sufixo}' if f'{nome_original}{sufixo}' in df_match.columns else nome_original

    col_status_dom = _col(COL_DOM_STATUS, 'dom') if COL_DOM_STATUS else None
    col_status_portal = _col(COL_PORTAL_STATUS, 'portal') if COL_PORTAL_STATUS else None
    col_num_dom = _col(COL_NUM_NF_DOM, 'dom')
    col_cnpj_dom = _col(COL_CNPJ_DOM, 'dom')

    if COL_DOM_STATUS and COL_PORTAL_STATUS:
        status_dom_serie = df_match[col_status_dom].astype(str).str.strip().str.upper()
        status_portal_serie = df_match[col_status_portal].astype(str).str.strip().str.upper()
        mask_urgente = status_portal_serie.isin(STATUS_CANCELADO_PORTAL) & status_dom_serie.isin(STATUS_ATIVO_DOMINIO)
        df_urgente = df_match.loc[mask_urgente, ['Chave_Busca', col_num_dom, col_cnpj_dom,
                                                 col_status_portal, col_status_dom]].copy()
        if not df_urgente.empty:
            df_urgente.columns = ['Chave', 'Nota', 'CNPJ', 'Status_Portal', 'Status_Dominio']
            df_urgente['Motivo'] = 'Cancelada no Portal, mas Ativa na Dominio'
        df_match_nao_urgente = df_match.loc[~mask_urgente].copy()
    else:
        df_urgente = pd.DataFrame(columns=['Chave', 'Nota', 'CNPJ', 'Status_Portal', 'Status_Dominio', 'Motivo'])
        df_match_nao_urgente = df_match

    # ---------- 5.2 Comparações financeiras (vetorizadas) ----------
    # Pre-converte cada coluna financeira uma unica vez e compara vetorialmente.
    _progress("Comparando valores", 6, 8)
    divergencias_por_linha = [[] for _ in range(len(df_match_nao_urgente))]
    for col_d, col_p, nome in colunas_comparar:
        cd = _col(col_d, 'dom')
        cp = _col(col_p, 'portal')
        vals_dom = df_match_nao_urgente[cd].map(parse_valor).round(2)
        vals_portal = df_match_nao_urgente[cp].map(parse_valor).round(2)
        mascara_diff = (vals_dom != vals_portal).values
        for i in np.where(mascara_diff)[0]:
            divergencias_por_linha[i].append(f"{nome}: Dominio {vals_dom.iat[i]} x Portal {vals_portal.iat[i]}")

    # ---------- 5.3 Comparação de datas (vetorizada) ----------
    if COL_DOM_DATA and COL_PORTAL_DATA:
        cd = _col(COL_DOM_DATA, 'dom')
        cp = _col(COL_PORTAL_DATA, 'portal')
        datas_dom = pd.to_datetime(df_match_nao_urgente[cd], dayfirst=True, errors='coerce')
        datas_portal = pd.to_datetime(df_match_nao_urgente[cp], dayfirst=True, errors='coerce')
        mascara_data = datas_dom.notna().values & datas_portal.notna().values & (datas_dom != datas_portal).values
        for i in np.where(mascara_data)[0]:
            divergencias_por_linha[i].append(f"Data Dominio: {datas_dom.iat[i].date()} | Data Portal: {datas_portal.iat[i].date()}")

    # ---------- 5.4 Classificação final ----------
    divergencias_texto = [' || '.join(lst) for lst in divergencias_por_linha]
    tem_divergencia = np.array([bool(s) for s in divergencias_texto])

    base_cols = {
        'Chave': df_match_nao_urgente['Chave_Busca'].values,
        'Nota': df_match_nao_urgente[col_num_dom].values,
        'CNPJ': df_match_nao_urgente[col_cnpj_dom].values,
        'Status_Dominio': (df_match_nao_urgente[col_status_dom].astype(str).str.strip().str.upper().values
                           if COL_DOM_STATUS else np.array(['Sem Coluna de Status'] * len(df_match_nao_urgente))),
        'Status_Portal': (df_match_nao_urgente[col_status_portal].astype(str).str.strip().str.upper().values
                          if COL_PORTAL_STATUS else np.array(['Sem Coluna de Status'] * len(df_match_nao_urgente))),
    }

    df_classificado = pd.DataFrame(base_cols)
    df_classificado['Detalhes_Divergencia'] = divergencias_texto
    df_divergencia = df_classificado.loc[tem_divergencia].drop(columns=['Detalhes_Divergencia']).join(
        df_classificado.loc[tem_divergencia, ['Detalhes_Divergencia']]
    ).reset_index(drop=True)
    df_tudo_certo = df_classificado.loc[~tem_divergencia].drop(columns=['Detalhes_Divergencia']).reset_index(drop=True)

    # ==============================================================================================
    # 6. GERAÇÃO DO RELATÓRIO DE SAÍDA E RESUMO
    # ==============================================================================================
    _progress("Gerando relatório", 7, 8)
    
    # Criar um resumo geral
    soma_valor_dom = sum(parse_valor(v) for v in df_dom[COL_DOM_VALOR_LIQ]) if COL_DOM_VALOR_LIQ else 0
    soma_valor_portal = sum(parse_valor(v) for v in df_portal[COL_PORTAL_VALOR_LIQ]) if COL_PORTAL_VALOR_LIQ else 0
    diferenca_total = soma_valor_dom - soma_valor_portal

    logger.info(f"Notas com divergência de valores: {len(df_divergencia)}")
    logger.info(f"Soma Valor Domínio: R$ {soma_valor_dom:.2f}")
    logger.info(f"Soma Valor Portal: R$ {soma_valor_portal:.2f}")
    logger.info(f"DIFERENÇA (FURO): R$ {diferenca_total:.2f}")
    logger.info("=" * 55)

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

    _progress("Salvando Excel", 8, 8)
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
    
    logger.info(f"✅ Conciliação concluída! Relatório salvo em: {caminho_saida}")
    
    # ==============================================================================================
    # 7. SALVAR HISTÓRICO E DIVERGÊNCIAS
    # ==============================================================================================
    # Tempo total de execução
    tempo_total = (datetime.now() - tempo_inicio).total_seconds()
    
    # Salvar histórico com metadados
    salvar_historico(
        caminho_saida=caminho_saida,
        caminho_dominio=caminho_dominio,
        caminho_portal=caminho_portal,
        tempo_execucao=tempo_total,
        num_notas_dom=len(df_dom),
        num_notas_portal=len(df_portal),
        num_divergencias=len(df_divergencia),
        num_faltantes_dom=len(df_faltantes_dominio),
        num_faltantes_portal=len(df_faltantes_portal),
        soma_dom=soma_valor_dom,
        soma_portal=soma_valor_portal
    )
    
    # Exportar divergências para CSV (se houver)
    if not df_divergencia.empty:
        caminho_csv = exportar_divergencias_csv(df_divergencia, caminho_saida)
        if caminho_csv:
            logger.info(f"💾 CSV de divergências disponível para investigação rápida")
    
    # ==============================================================================================
    # 8. ENVIAR ALERTA POR EMAIL (SE HABILITADO E HOUVER DIVERGÊNCIAS)
    # ==============================================================================================
    if enviar_email:
        config_email = carregar_config_email()
        if config_email:
            enviar_alerta_email(
                config=config_email,
                num_divergencias=len(df_divergencia),
                soma_dom=soma_valor_dom,
                soma_portal=soma_valor_portal,
                diferenca=diferenca_total,
                tempo_execucao=tempo_total,
                arquivo_saida=caminho_saida
            )
        else:
            logger.warning("Alertas por email solicitado, mas arquivo email.config não encontrado ou inválido.")


# ==============================================================================================
# INTERFACE GRÁFICA (GUI)
# ==============================================================================================
class ConciliadorGUI:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Conciliador Automático de NFS-e")
        self.root.geometry("550x420")
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

        # Botão de execução
        self.btn_executar = tk.Button(root, text="▶ Executar Conciliação", bg="#4CAF50", fg="white",
                                       font=("Arial", 12, "bold"), command=self.executar)
        self.btn_executar.pack(pady=15)
        
        # Botão para abrir histórico/divergências (disponível após execução)
        frame_uteis = tk.Frame(root)
        frame_uteis.pack(pady=10)
        self.btn_abrir_pasta = tk.Button(frame_uteis, text="📁 Abrir Histórico/Divergências", bg="#2196F3", 
                                         fg="white", font=("Arial", 10), command=self.abrir_pasta_saida)
        self.btn_abrir_pasta.pack(side=tk.LEFT, padx=5)
        self.btn_abrir_pasta.config(state='disabled')
        
        # Checkbox para habilitar alertas por email
        frame_email = tk.Frame(root)
        frame_email.pack(pady=5)
        self.enviar_email = tk.BooleanVar(value=False)
        self.check_email = tk.Checkbutton(frame_email, text="📧 Enviar alerta por email se houver divergências", 
                                          variable=self.enviar_email, font=("Arial", 9))
        self.check_email.pack(side=tk.LEFT, padx=5)
        tk.Label(frame_email, text="(requer email.config)", font=("Arial", 8), fg="#999").pack(side=tk.LEFT, padx=2)

        # Barra de progresso determinada + status (visíveis apenas durante execução)
        self.progress = ttk.Progressbar(root, mode='determinate', length=400, maximum=8)
        self.status_label = tk.Label(root, text="", font=("Arial", 9), fg="#555")

    def selecionar_dominio(self):
        arquivo = filedialog.askopenfilename(title="Selecione a planilha da Domínio",
                                             filetypes=[("Arquivos Excel/ODS", "*.xlsx *.xls *.ods")])
        if arquivo:
            self.caminho_dominio.set(arquivo)

    def selecionar_portal(self):
        arquivo = filedialog.askopenfilename(title="Selecione a planilha do Portal",
                                             filetypes=[("Arquivos Excel/ODS", "*.xlsx *.xls *.ods")])
        if arquivo:
            self.caminho_portal.set(arquivo)

    def selecionar_saida(self):
        arquivo = filedialog.asksaveasfilename(title="Salvar relatório final", defaultextension=".xlsx",
                                               filetypes=[("Arquivo Excel", "*.xlsx")])
        if arquivo:
            self.caminho_saida.set(arquivo)

    def abrir_pasta_saida(self) -> None:
        """Abre a pasta de histórico/divergências no explorador de arquivos."""
        if not self.caminho_saida.get():
            messagebox.showwarning("Aviso", "Nenhum caminho de saída definido ainda.")
            return
        
        pasta = os.path.dirname(self.caminho_saida.get())
        if os.path.exists(pasta):
            # Abre a pasta no explorador (Windows)
            import subprocess
            subprocess.Popen(f'explorer "{pasta}"')
        else:
            messagebox.showwarning("Aviso", f"Pasta não encontrada: {pasta}")
    
    def _set_busy(self, busy: bool) -> None:
        """Habilita/desabilita controles e mostra/oculta a barra de progresso."""
        state = 'disabled' if busy else 'normal'
        self.btn_executar.config(state=state)
        if busy:
            self.progress.pack(pady=(0, 5))
            self.status_label.pack(pady=(0, 10))
            self.progress['value'] = 0  # Reseta a barra
            self.status_label.config(text="Inicializando...")
        else:
            self.progress.pack_forget()
            self.status_label.pack_forget()
            self.progress['value'] = 0

    def _progress_callback(self, etapa: str, valor: int, maximo: int) -> None:
        """Callback para atualizar progresso da conciliação."""
        self.progress['value'] = valor
        self.status_label.config(text=etapa)
        self.root.update_idletasks()

    def _run_in_thread(self, caminho_dom: str, caminho_port: str, caminho_out: str, enviar_email: bool = False) -> None:
        """Worker: roda conciliar_nfse e agenda o callback na thread principal."""
        resultado = {'erro': None}
        try:
            conciliar_nfse(caminho_dom, caminho_port, caminho_out, progress_callback=self._progress_callback, enviar_email=enviar_email)
        except Exception as e:
            logger.exception("Erro durante conciliação")
            resultado['erro'] = e
        # Volta para a thread do Tk com after()
        self.root.after(0, self._on_done, resultado)

    def _on_done(self, resultado: Dict) -> None:
        """Callback executado na thread do Tk após o worker terminar."""
        self._set_busy(False)
        caminho = self.caminho_saida.get()
        if resultado['erro'] is None:
            self.btn_abrir_pasta.config(state='normal')  # Habilita botão de histórico
            messagebox.showinfo("Sucesso", f"Conciliação concluída!\n\nRelatório salvo em:\n{caminho}\n\nClique em 'Abrir Histórico/Divergências' para acessar os arquivos de suporte.")
        else:
            e = resultado['erro']
            if isinstance(e, KeyError):
                messagebox.showerror("Erro de Coluna",
                    f"Erro: {str(e)}\n\nO layout da planilha mudou ou as colunas esperadas não estão presentes.")
            else:
                messagebox.showerror("Erro Inesperado", f"Ocorreu um erro durante a conciliação:\n{str(e)}")

    def executar(self) -> None:
        """Valida campos e inicia conciliação em thread separada."""
        if not self.caminho_dominio.get() or not self.caminho_portal.get() or not self.caminho_saida.get():
            messagebox.showwarning("Aviso", "Por favor, preencha todos os campos.")
            return

        self._set_busy(True)
        t = threading.Thread(target=self._run_in_thread,
                             args=(self.caminho_dominio.get(),
                                   self.caminho_portal.get(),
                                   self.caminho_saida.get(),
                                   self.enviar_email.get()),
                             daemon=True)
        t.start()

if __name__ == "__main__":
    root = tk.Tk()
    app = ConciliadorGUI(root)
    root.mainloop()
