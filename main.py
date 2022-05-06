import pandas as pd
import win32api, os, pikepdf, datetime, re
from datetime import date
from dateutil.relativedelta import relativedelta
from dateutil.tz import tzutc, tzoffset

df = pd.read_excel(f'INPUT.xlsx')

unformat_cnpj = lambda cnpj: str(cnpj).replace('.','').replace('/','').replace('-','')

def format_cnpj(cnpj):
    if len(str(cnpj)) != 14: return 
    formated = f"{cnpj[0]}{cnpj[1]}.{cnpj[2]}{cnpj[3]}{cnpj[4]}.{cnpj[5]}{cnpj[6]}{cnpj[7]}/{cnpj[8]}{cnpj[9]}{cnpj[10]}{cnpj[11]}-{cnpj[12]}{cnpj[13]}"
    return formated

def transform_date(date_str):
    global pdf_date_pattern
    pdf_date_pattern = re.compile(''.join([
        r"(D:)?",
        r"(?P<year>\d\d\d\d)",
        r"(?P<month>\d\d)",
        r"(?P<day>\d\d)",
        r"(?P<hour>\d\d)",
        r"(?P<minute>\d\d)",
        r"(?P<second>\d\d)",
        r"(?P<tz_offset>[+-zZ])?",
        r"(?P<tz_hour>\d\d)?",
        r"'?(?P<tz_minute>\d\d)?'?"]))
    match = re.match(pdf_date_pattern, date_str)
    if match:
        date_info = match.groupdict()
        for k, v in date_info.items():
            if v is None: pass
            elif k == 'tz_offset': date_info[k] = v.lower()
            else: date_info[k] = int(v)
        if date_info['tz_offset'] in ('z', None): date_info['tzinfo'] = tzutc()
        else:
            multiplier = 1 if date_info['tz_offset'] == '+' else -1
            date_info['tzinfo'] = tzoffset(None, multiplier*(3600 * date_info['tz_hour'] + 60 * date_info['tz_minute']))
        for k in ('tz_offset', 'tz_hour', 'tz_minute'): del date_info[k]
        return datetime.datetime(**date_info)

def crf_validity(cnpj):
    for file in os.listdir('./CERTIDOES'):
        if file == f"CRF-{unformat_cnpj(cnpj)}.pdf":
            file_data = pikepdf.Pdf.open(f'./CERTIDOES/{file}')
            docinfo = file_data.docinfo
            for k,v in docinfo.items():
                if k == "/CreationDate":
                    if transform_date(str(v)).date() > (date.today() + relativedelta(days=29)): return False
                    else: return True

def crf_present(cnpj):
    files = os.listdir('./CERTIDOES')
    a = unformat_cnpj(cnpj)
    if f"CRF-{a}.pdf" in files: return True
    else: return False

def check_crf(cnpj):
    if crf_present(cnpj):
        print(f'[CRF] - A CRF do CNPJ {cnpj} Está presente no diretório')
        if crf_validity(cnpj):
            print(f'[CRF] - A CRF do CNPJ {cnpj} Ainda é válida.')
            return True
        else: print(f'[CRF] - A CRF do CNPJ {cnpj} Não é mais válida. É Necessário a reemissão'); return False
    else: print(f'[CRF] - A CRF do CNPJ {cnpj} Não está presente no diretório'); return False

def cnd_validity(cnpj):
    for file in os.listdir('./CERTIDOES'):
        if file == f"CND-{unformat_cnpj(cnpj)}.pdf":
            file_data = pikepdf.Pdf.open(f'./CERTIDOES/{file}')
            docinfo = file_data.docinfo
            for k,v in docinfo.items():
                if k == "/CreationDate":
                    if transform_date(str(v)).date() > (date.today() + relativedelta(days=29)): return False
                    else: return True

def cnd_present(cnpj):
    files = os.listdir('./CERTIDOES')
    treated = unformat_cnpj(cnpj)
    if f"CND-{treated}.pdf" in files: return True
    else: return False

def check_cnd(cnpj):
    if cnd_present(cnpj):
        print(f'[CND] - A CND do CNPJ {cnpj} Está presente no diretório')
        if cnd_validity(cnpj):
            print(f'[CND] - A CND do CNPJ {cnpj} Ainda é válida.')
            return True
        else: print(f'[CND] - A CND do CNPJ {cnpj} Não é mais válida. É Necessário a reemissão'); return False
    else: print(f'[CND] - A CND do CNPJ {cnpj} Não está presente no diretório'); return False

for indexes,row in df.iterrows():
    if check_cnd(row['cnpj']) and check_crf(row['cnpj']):
        for i in range(row['quantidade']):
            win32api.ShellExecute(0,'print', f"CND-{unformat_cnpj(row['cnpj'])}.pdf", None, './CERTIDOES', 0)
            win32api.ShellExecute(0,'print', f"CRF-{unformat_cnpj(row['cnpj'])}.pdf", None, './CERTIDOES', 0)
        df = df.drop(index=indexes, axis=0)
    elif check_cnd(row['cnpj']) and not check_crf(row['cnpj']):
        for i in range(row['quantidade']):
            win32api.ShellExecute(0,'print', f"CND-{unformat_cnpj(row['cnpj'])}.pdf", None, './CERTIDOES', 0)
        df = df.drop(index=indexes, axis=0)
    elif not check_cnd(row['cnpj']) and check_crf(row['cnpj']):
        for i in range(row['quantidade']):
            win32api.ShellExecute(0,'print', f"CRF-{unformat_cnpj(row['cnpj'])}.pdf", None, './CERTIDOES', 0)
        df = df.drop(index=indexes, axis=0)
    df['cnpj'] = format_cnpj(row['cnpj'])
df.to_excel(f"INPUT.xlsx", index=False, startcol=0)