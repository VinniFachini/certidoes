import pandas as pd
import os
import pikepdf
import datetime
import re
import logging
# import win32api
from datetime import date
from dateutil.relativedelta import relativedelta
from dateutil.tz import tzutc, tzoffset


logging.basicConfig(level=logging.DEBUG, filename="./test.log",
                    filemode="a", format="%(asctime)s - %(levelname)s: %(message)s")

archive_input = str(input('Digite o nome do arquivo de entrada: ').upper() + '.xlsx')
if archive_input.startswith('.'):
    archive_input = 'INPUT.xlsx'

try:
    df = pd.read_excel(archive_input)
    logging.debug('Banco de Dados carregado com sucesso!')
except Exception:
    logging.critical(f'Houve um problema na importação do banco de dados! Arquivo "{archive_input}" não encontrado')
    exit();

df = pd.read_excel(f'INPUT.xlsx')

def unformat_cnpj(cnpj): return str(cnpj).replace(
    '.', '').replace('/', '').replace('-', '')


def transform_date(date_str):
    global pdf_date_pattern
    pdf_date_pattern = re.compile(''.join([r"(D:)?", r"(?P<year>\d\d\d\d)", r"(?P<month>\d\d)", r"(?P<day>\d\d)", r"(?P<hour>\d\d)",
                                  r"(?P<minute>\d\d)", r"(?P<second>\d\d)", r"(?P<tz_offset>[+-zZ])?", r"(?P<tz_hour>\d\d)?", r"'?(?P<tz_minute>\d\d)?'?"]))
    match = re.match(pdf_date_pattern, date_str)
    if match:
        date_info = match.groupdict()
        for k, v in date_info.items():
            if v is None:
                pass
            elif k == 'tz_offset':
                date_info[k] = v.lower()
            else:
                date_info[k] = int(v)
        if date_info['tz_offset'] in ('z', None):
            date_info['tzinfo'] = tzutc()
        else:
            multiplier = 1 if date_info['tz_offset'] == '+' else -1
            date_info['tzinfo'] = tzoffset(
                None, multiplier*(3600 * date_info['tz_hour'] + 60 * date_info['tz_minute']))
        for k in ('tz_offset', 'tz_hour', 'tz_minute'):
            del date_info[k]
        return datetime.datetime(**date_info)


for indexes, row in df.iterrows():
    files = os.listdir('./CERTIDOES')
    treated = unformat_cnpj(row['cnpj'])
    if f"CND-{treated}.pdf" in files and f"CRF-{treated}.pdf" in files:
        logging.info(
            f'A CRF do CNPJ {row["cnpj"]} Está presente no diretório (1/2)')
        logging.info(
            f'A CND do CNPJ {row["cnpj"]} Está presente no diretório (2/2)')
        file_data01 = pikepdf.Pdf.open(f'./CERTIDOES/CND-{treated}.pdf')
        file_data02 = pikepdf.Pdf.open(f'./CERTIDOES/CRF-{treated}.pdf')
        docinfo01 = file_data01.docinfo
        docinfo02 = file_data02.docinfo
        for k, v in docinfo01.items():
            if k == "/CreationDate":
                if not(transform_date(str(v)).date() > (date.today() + relativedelta(days=179))):
                    logging.info(
                        f'A CND do CNPJ {row["cnpj"]} ainda está válida!')
                    for k, v in docinfo02.items():
                        if k == "/CreationDate":
                            if not(transform_date(str(v)).date() > (date.today() + relativedelta(days=29))):
                                logging.info(
                                    f'A CRF do CNPJ {row["cnpj"]} ainda está válida!')
                                for i in range(row['quantidade']):
                                    logging.info(
                                        f'Imprimindo a CND do CNPJ {row["cnpj"]} ({i+1}/{row["quantidade"]})')
                                    # win32api.ShellExecute(0,'print', f"CND-{unformat_cnpj(row['cnpj'])}.pdf", None, './CERTIDOES', 0)
                                    logging.info(
                                        f'Imprimindo a CRF do CNPJ {row["cnpj"]} ({i+1}/{row["quantidade"]})')
                                    # win32api.ShellExecute(0,'print', f"CRF-{unformat_cnpj(row['cnpj'])}.pdf", None, './CERTIDOES', 0)
                            else:
                                logging.error(
                                    f'A CRF do CNPJ {row["cnpj"]} não está válida. É necessária a reemissão!')
                else:
                    logging.error(
                        f'A CND do CNPJ {row["cnpj"]} não está válida. É necessária a reemissão!')

    elif f"CND-{treated}.pdf" in files and f"CRF-{treated}.pdf" not in files:
        logging.info(
            f'Apenas a CND do CNPJ {row["cnpj"]} Está presente no diretório')
        file_data = pikepdf.Pdf.open(f'./CERTIDOES/CND-{treated}.pdf')
        docinfo = file_data.docinfo
        for k, v in docinfo.items():
            if k == "/CreationDate":
                if not(transform_date(str(v)).date() > (date.today() + relativedelta(days=179))):
                    logging.info(
                        f'A CND do CNPJ {row["cnpj"]} ainda está válida!')
                    for i in range(row['quantidade']):
                        logging.info(
                            f'Imprimindo a CND do CNPJ {row["cnpj"]} ({i+1}/{row["quantidade"]})')
                        # win32api.ShellExecute(0,'print', f"CND-{unformat_cnpj(row['cnpj'])}.pdf", None, './CERTIDOES', 0)
                else:
                    logging.error(
                        f'A CND do CNPJ {row["cnpj"]} não está válida. É necessária a reemissão!')

    elif f"CND-{treated}.pdf" not in files and f"CRF-{treated}.pdf" in files:
        logging.info(
            f'Apenas a CRF do CNPJ {row["cnpj"]} Está presente no diretório')
        file_data = pikepdf.Pdf.open(f'./CERTIDOES/CRF-{treated}.pdf')
        docinfo = file_data.docinfo
        for k, v in docinfo.items():
            if k == "/CreationDate":
                if not(transform_date(str(v)).date() > (date.today() + relativedelta(days=29))):
                    logging.info(
                        f'A CRF do CNPJ {row["cnpj"]} ainda está válida!')
                    for i in range(row['quantidade']):
                        logging.info(
                            f'Imprimindo a CRF do CNPJ {row["cnpj"]} ({i+1}/{row["quantidade"]})')
                        # win32api.ShellExecute(0,'print', f"CRF-{unformat_cnpj(row['cnpj'])}.pdf", None, './CERTIDOES', 0)
                else:
                    logging.error(
                        f'A CRF do CNPJ {row["cnpj"]} não está válida. É necessária a reemissão!')

    elif f"CND-{treated}.pdf" not in files and f"CRF-{treated}.pdf" not in files:
        logging.error(
            f'A CRF do CNPJ {row["cnpj"]} Não está presente no diretório (1/2)')
        logging.error(
            f'A CND do CNPJ {row["cnpj"]} Não está presente no diretório (2/2)')

logging.debug('Término de execução')
