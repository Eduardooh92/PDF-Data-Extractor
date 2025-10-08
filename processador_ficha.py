import fitz  # PyMuPDF
import openpyxl
import re
import os
import glob
import logging
import configparser
import shutil
from logging.handlers import RotatingFileHandler

# --- CONFIGURAÇÃO INICIAL ---
# Lê as configurações do arquivo .ini, que é o novo cérebro do sistema.
config = configparser.ConfigParser()
config.read('config.ini')

try:
    PDF_INPUT_FOLDER = config.get('Paths', 'InputFolder')
    OUTPUT_DIR = config.get('Paths', 'OutputFolder')
    PROCESSED_DIR = config.get('Paths', 'ProcessedFolder')
    ERROR_DIR = config.get('Paths', 'ErrorFolder')
    EXCEL_TEMPLATE_FILE = config.get('Paths', 'ExcelTemplate')
    LOG_FILE = config.get('Settings', 'LogFile')
except configparser.NoSectionError as e:
    print(f"Erro Crítico: Seção faltando no arquivo config.ini. Detalhes: {e}")
    exit(1)
except configparser.NoOptionError as e:
    print(f"Erro Crítico: Opção faltando no arquivo config.ini. Detalhes: {e}")
    exit(1)

# Cria as pastas de controle se elas não existirem.
os.makedirs(PROCESSED_DIR, exist_ok=True)
os.makedirs(ERROR_DIR, exist_ok=True)

# --- CONFIGURAÇÃO AVANÇADA DE LOGS ---
# Log em arquivo com rotação (ex: 5MB por arquivo, mantém 3 arquivos de backup)
log_handler = RotatingFileHandler(LOG_FILE, maxBytes=5*1024*1024, backupCount=3, encoding='utf-8')
log_formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
log_handler.setFormatter(log_formatter)

# Adiciona o handler ao logger raiz
logger = logging.getLogger()
logger.setLevel(logging.INFO)
logger.addHandler(log_handler)
# Adiciona um handler para também mostrar o log no console
console_handler = logging.StreamHandler()
console_handler.setFormatter(log_formatter)
logger.addHandler(console_handler)


# Mapeamento de escrita.
CELL_MAPPING = {
    'razao_social': 'I12',
    'nome_fantasia': 'I14',
    'endereco': 'I16',
    'bairro': 'I18',
    'cep_parte1': 'AG18',
    'cep_parte2': 'AN18',
    'uf': 'AU18',
    'cidade': 'I20',
    'cnpj_PART1': 'AF22',
    'cnpj_PARTE2': 'AU22',
    'insc_estadual': 'I24',
    'segmento': 'AI14',
    'representante': 'AI24'
}

# --- FUNÇÕES DE LÓGICA (sem alterações, apenas usando 'logger' em vez de 'logging') ---

def extract_text_from_pdf(pdf_path):
    if not pdf_path or not os.path.exists(pdf_path):
        return None
    try:
        logger.info(f"Iniciando extração de texto do arquivo: {os.path.basename(pdf_path)}")
        doc = fitz.open(pdf_path)
        text = "".join(page.get_text() for page in doc)
        doc.close()
        logger.info(f"Extração de texto de {os.path.basename(pdf_path)} concluída.")
        return text
    except Exception as e:
        logger.error(f"Falha ao ler o arquivo PDF '{pdf_path}'. Erro: {e}")
        return None

def parse_cnpj_data(text):
    # (Esta e as outras funções de parse e apply_rules permanecem idênticas,
    # apenas troque 'logging.' por 'logger.' se quiser consistência)
    logger.info("Analisando texto do PDF de CNPJ...")
    data = {}
    patterns = {
        'cnpj': r'NÚMERO DE INSCRIÇÃO\s*\n(.*?)\n',
        'razao_social': r'NOME EMPRESARIAL\s*\n(.*?)\n',
        'nome_fantasia': r'TÍTULO DO ESTABELECIMENTO \(NOME DE FANTASIA\)\s*\n(.*?)\n',
        'logradouro': r'LOGRADOURO\s*\n(.*?)\n',
        'numero': r'NÚMERO\s*\n\s*(\d+)\s*\n',
        'complemento': r'COMPLEMENTO\s*\n(.*?)\n',
        'bairro': r'BAIRRO/DISTRITO\s*\n(.*?)\n',
        'cep': r'CEP\s*\n(.*?)\n',
        'cidade': r'MUNICÍPIO\s*\n(.*?)\n',
        'uf': r'UF\s*\n(.*?)\n',
        'segmento': r'CÓDIGO E DESCRIÇÃO DA ATIVIDADE ECONÔMICA PRINCIPAL\s*\n(.*?)\n'
    }

    for key, pattern in patterns.items():
        match = re.search(pattern, text, re.IGNORECASE)
        data[key] = match.group(1).strip() if match else ''
    if data.get('razao_social'):
        logger.info(f"Dados do CNPJ extraídos com sucesso para: {data['razao_social']}")
    else:
        logger.warning("Não foi possível extrair a Razão Social do PDF de CNPJ.")
    return data

def parse_ie_data(text):
    logger.info("Analisando texto do PDF de Inscrição Estadual...")
    data = {}
    pattern = r'INSCRIÇÃO\s*[:\-]?\s*([\d.\-]+)'
    match = re.search(pattern, text, re.IGNORECASE)
    data['insc_estadual'] = match.group(1).strip() if match else ''
    if data.get('insc_estadual'):
        logger.info(f"Inscrição Estadual encontrada: {data['insc_estadual']}")
    else:
        logger.warning("Nenhuma Inscrição Estadual encontrada no PDF correspondente.")
    return data

def apply_business_rules(data):
    logger.info("Aplicando regras de negócio e formatação...")
    processed_data = data.copy()
    logradouro = data.get('logradouro', '')
    numero = data.get('numero', '')
    complemento = data.get('complemento', '')
    partes_endereco = []
    if logradouro:
        partes_endereco.append(logradouro.strip())
    if numero:
        partes_endereco.append(f"Nº {numero.strip()}")
    if complemento and '********' not in complemento:
        partes_endereco.append(complemento.strip())
    processed_data['endereco'] = " ; ".join(partes_endereco)
    
    cep_completo = data.get('cep', '')
    if cep_completo and '-' in cep_completo:
        cep_limpo = re.sub(r'[^\d-]', '', cep_completo)
        partes = cep_limpo.split('-')
        if len(partes) == 2:
            processed_data['cep_parte1'] = partes[0]
            processed_data['cep_parte2'] = partes[1]
    
    cnpj_completo = data.get('cnpj', '')    
    if cnpj_completo and '-' in cnpj_completo:
        partes = cnpj_completo.split('-')
        if len(partes) == 2:
            processed_data['cnpj_PART1'] = partes[0]
            processed_data['cnpj_PARTE2'] = partes[1]
            
    processed_data['representante'] = 'A MARQUES'
    return processed_data

def fill_excel_template(template_path, output_path, data):
    try:
        logger.info(f"Abrindo template do Excel: {template_path}")
        workbook = openpyxl.load_workbook(template_path)
        sheet = workbook.active
        for key, cell_coord in CELL_MAPPING.items():
            if key in data and data[key]:
                try:
                    logger.info(f"Escrevendo '{data[key]}' na célula {cell_coord}.")
                    sheet[cell_coord] = data[key]
                except AttributeError as e:
                    if "'MergedCell' object attribute 'value' is read-only" in str(e):
                        logger.warning(f"Célula {cell_coord} é parte de um bloco mesclado. Ignorando.")
                    else:
                        raise e
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        logger.info(f"Salvando arquivo final em: {output_path}")
        workbook.save(output_path)
        logger.info("Arquivo Excel salvo com sucesso.")
        return True
    except FileNotFoundError:
        logger.error(f"Template do Excel '{template_path}' não encontrado.")
        return False
    except PermissionError:
        logger.error(f"Falha ao salvar o Excel em '{output_path}'. O arquivo pode estar aberto ou você não tem permissão.")
        return False
    except Exception as e:
        logger.error(f"Falha inesperada ao escrever no arquivo Excel. Erro: {e}")
        return False

def move_file(src_path, dest_folder):
    """Move um arquivo para a pasta de destino, garantindo que não haja sobrescrita."""
    try:
        dest_path = os.path.join(dest_folder, os.path.basename(src_path))
        # Adiciona um sufixo numérico se o arquivo já existir no destino
        if os.path.exists(dest_path):
            base, ext = os.path.splitext(dest_path)
            i = 1
            while os.path.exists(f"{base}_{i}{ext}"):
                i += 1
            dest_path = f"{base}_{i}{ext}"
        shutil.move(src_path, dest_path)
        logger.info(f"Arquivo {os.path.basename(src_path)} movido para {dest_folder}")
    except Exception as e:
        logger.error(f"Não foi possível mover o arquivo {os.path.basename(src_path)} para {dest_folder}. Erro: {e}")

# --- BLOCO DE EXECUÇÃO PRINCIPAL (PRODUÇÃO) ---
if __name__ == "__main__":
    logger.info("======================================================")
    logger.info("Iniciando processador de clientes em modo de PRODUÇÃO.")
    
    all_pdfs_in_folder = glob.glob(os.path.join(PDF_INPUT_FOLDER, '*.pdf'))
    
    if not all_pdfs_in_folder:
        logger.info("Nenhum arquivo PDF encontrado na pasta de entrada. Encerrando.")
        exit(0)

    # Agrupando arquivos por cliente (assumindo que todos na pasta são de um cliente)
    # Para múltiplos clientes, a lógica precisaria de um agrupador.
    
    final_data = {}
    cnpj_data_found = False
    ie_data_found = False
    files_to_process = list(all_pdfs_in_folder) # Cópia para iterar
    
    for pdf_path in files_to_process:
        try:
            text = extract_text_from_pdf(pdf_path)
            if not text:
                logger.warning(f"Texto não extraído de {os.path.basename(pdf_path)}. Movendo para erros.")
                move_file(pdf_path, ERROR_DIR)
                continue

            text_upper = text.upper()
            if 'NOME EMPRESARIAL' in text_upper and 'NÚMERO DE INSCRIÇÃO' in text_upper:
                logger.info(f"Arquivo {os.path.basename(pdf_path)} identificado como CNPJ.")
                final_data.update(parse_cnpj_data(text))
                cnpj_data_found = True
            elif 'INSCRIÇÃO' in text_upper and re.search(r'INSCRIÇÃO\s*[:\-]?\s*([\d.\-]+)', text, re.IGNORECASE):
                logger.info(f"Arquivo {os.path.basename(pdf_path)} identificado como IE.")
                final_data.update(parse_ie_data(text))
                ie_data_found = True
            else:
                logger.warning(f"Conteúdo de {os.path.basename(pdf_path)} não reconhecido. Movendo para erros.")
                move_file(pdf_path, ERROR_DIR)
        except Exception as e:
            logger.critical(f"Erro catastrófico ao processar {os.path.basename(pdf_path)}: {e}")
            move_file(pdf_path, ERROR_DIR)
    
    if not cnpj_data_found:
        logger.critical("Nenhum arquivo com conteúdo de CNPJ foi processado com sucesso. Planilha não será gerada.")
        # Move os arquivos restantes (se houver) que não foram movidos individualmente
        for pdf in glob.glob(os.path.join(PDF_INPUT_FOLDER, '*.pdf')):
             move_file(pdf, ERROR_DIR)
    else:
        if not ie_data_found:
            logger.warning("Nenhum arquivo de IE identificado. Aplicando regra 'Isento'.")
            final_data['insc_estadual'] = 'Isento'

        processed_data = apply_business_rules(final_data)
        
        if not processed_data.get('razao_social'):
            logger.error("Razão Social não extraída. O arquivo Excel não será gerado. Movendo arquivos para pasta de erro.")
            for pdf in glob.glob(os.path.join(PDF_INPUT_FOLDER, '*.pdf')):
                move_file(pdf, ERROR_DIR)
        else:
            razao_social_clean = re.sub(r'[\\/*?:"<>|]', "", processed_data.get('razao_social'))
            output_filename = f"FICHA CADASTRAL - {razao_social_clean}.xlsx"
            output_filepath = os.path.join(OUTPUT_DIR, output_filename)
            
            success = fill_excel_template(EXCEL_TEMPLATE_FILE, output_filepath, processed_data)
            
            # Movimentação final dos arquivos baseada no sucesso da operação
            if success:
                logger.info("Operação bem-sucedida. Movendo arquivos de origem para a pasta de processados.")
                for pdf in glob.glob(os.path.join(PDF_INPUT_FOLDER, '*.pdf')):
                    move_file(pdf, PROCESSED_DIR)
            else:
                logger.error("Falha ao gerar a planilha Excel. Movendo arquivos de origem para a pasta de erro.")
                for pdf in glob.glob(os.path.join(PDF_INPUT_FOLDER, '*.pdf')):
                    move_file(pdf, ERROR_DIR)
    
    logger.info("Processo finalizado.")
    logger.info("======================================================")
    input("\nPressione Enter para sair...")