    """Converte a planilha JEO 2026 em dados.json para o painel GitHub Pages."""
import json
import re
from pathlib import Path
from datetime import date, datetime
from openpyxl import load_workbook

ARQUIVO_JSON = 'dados.json'
ABAS_INICIATIVAS = {
    '2000 - SE': 'SE', '216H - SE': 'SE', '4641- SE': 'SE',
    '21FC - SENAPIR': 'SENAPIR', '21FE - SQPT': 'SQPT',
    '21FF - SEPAR': 'SEPAR', '21HN - SEPAR': 'SEPAR',
    '21FG - CNPIR': 'CNPIR', '21HN - OUV': 'OUV',
}
UNIDADES = ('SEPAR', 'OUV', 'SENAPIR', 'SQPT', 'CNPIR', 'SE')

def ss(value):
    if value is None:
        return ''
    if isinstance(value, (date, datetime)):
        return value.strftime('%d/%m/%Y')
    return str(value).strip()

def nn(value):
    if value is None or isinstance(value, bool):
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    value = ss(value).replace('R$', '').replace(' ', '')
    if not value or value.startswith('#') or value in ('-', '—'):
        return 0.0
    if ',' in value:
        value = value.replace('.', '').replace(',', '.')
    try:
        return float(value)
    except ValueError:
        return 0.0

def at(row, index):
    return row[index] if 0 <= index < len(row) else None

def cabecalhos(row):
    return {re.sub(r'\s+', ' ', ss(v)).upper(): i for i, v in enumerate(row) if ss(v)}

def pegar(row, header, prefix):
    for title, index in header.items():
        if title.startswith(prefix):
            return at(row, index)
    return None

def localizar_excel(folder):
    candidates = sorted(folder.glob('*.xlsx'))
    valid = [p for p in candidates if not p.name.startswith('~$') and 'JEO' in p.name.upper()]
    if len(valid) == 1:
        return valid[0]
    if not valid:
        raise FileNotFoundError('Nenhuma planilha Excel JEO (.xlsx) encontrada no repositório.')
    raise ValueError('Há mais de uma planilha JEO no repositório: ' + ', '.join(p.name for p in valid))

def converter(caminho):
    wb = load_workbook(caminho, data_only=True, read_only=True)
    try:
        if 'Consolidado 2026' not in wb.sheetnames:
            raise ValueError('Aba Consolidado 2026 não encontrada.')
        consolidado = list(wb['Consolidado 2026'].values)
        total_row = next((r for r in consolidado if ss(at(r, 0)).upper() == 'TOTAL GERAL'), None)
        if total_row is None:
            raise ValueError('Linha TOTAL GERAL não encontrada no Consolidado 2026.')
        labels = ('ploa', 'loa', 'dot', 'bloq', 'bloq_sof', 'livre', 'prog', 'desc', 'aguard', 'emp', 'liq', 'pago', 'saldo', 'pct_prog', 'pct_emp', 'pct_liq', 'pct_pago')
        columns = (2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18)
        tot = {key: nn(at(total_row, col)) for key, col in zip(labels, columns)}
        tot.update({'programado': tot['prog'], 'empenhado': tot['emp'], 'liquidado': tot['liq']})
        lme_sheet = wb['LIMITE DE MOVIMENTAÇÃO EMPENHO'] if 'LIMITE DE MOVIMENTAÇÃO EMPENHO' in wb.sheetnames else None
        if lme_sheet:
            lme_values = next(lme_sheet.iter_rows(min_row=2, max_row=2, values_only=True))
            lme = {'limite': nn(at(lme_values, 0)), 'aguard': nn(at(lme_values, 1)), 'emp': nn(at(lme_values, 2)), 'saldo': nn(at(lme_values, 6))}
        else:
            r = consolidado[3]
            lme = {'limite': nn(at(r, 1)), 'aguard': nn(at(r, 2)), 'emp': nn(at(r, 3)), 'saldo': nn(at(r, 7))}
        upd = next((ss(v) for r in consolidado for v in r if re.match(r'^Atualiza', ss(v), re.I)), '')
        units = []
        for row in consolidado:
            unidade = ss(at(row, 0)).upper()
            unidade = {'OUVIDORIA': 'OUV', 'SECRETARIA EXECUTIVA': 'SE'}.get(unidade, unidade)
            if unidade in UNIDADES:
                units.append(dict(u=unidade, name=unidade, prog=nn(at(row, 8)), emp=nn(at(row, 11)), liq=nn(at(row, 12)), pago=nn(at(row, 13)), dot=nn(at(row, 4)), livre=nn(at(row, 7)), saldo=nn(at(row, 14))))
        # As linhas consolidadas nem sempre incluem todas as unidades separadamente.
        present = {u['u'] for u in units}
        for unit in UNIDADES:
            if unit not in present:
                units.append(dict(u=unit, name=unit, prog=0, emp=0, liq=0, pago=0, dot=0, livre=0, saldo=0))
        ini = []
        for sheet_name, unit in ABAS_INICIATIVAS.items():
            if sheet_name not in wb.sheetnames:
                continue
            header = None
            for row in wb[sheet_name].values:
                if ss(at(row, 0)).upper() == 'UNIDADE':
                    header = cabecalhos(row)
                    continue
                if header is None:
                    continue
                get = lambda prefix: pegar(row, header, prefix)
                name = ss(get('INICIATIVAS'))
                prog = nn(get('PROGRAMADO'))
                emp = nn(get('EMPENHADO'))
                if not (prog or emp) or name.upper().startswith('TOTAL'):
                    continue
                obs_key = next((k for k in header if k.startswith('OBSERVAÇÃO')), None)
                ini.append(dict(
                    u=unit, acao=sheet_name.split(' ')[0], po=ss(get('PO')),
                    name=name or '(sem descrição) — ' + (ss(get('PARCEIRO')) or ss(get('DIRETORIA')) or '—'),
                    sei=ss(get('SEI')), dir=ss(get('DIRETORIA')),
                    resp=ss(get('RESPONSÁVEL')), inst=ss(get('INSTRUMENTO')),
                    siafi=ss(get('SIAFI')), parc=ss(get('PARCEIRO')),
                    ano=ss(get('ANO DE')), prog=prog, emp=emp,
                    liq=nn(get('LIQUIDADO')), pago=nn(get('PAGO')),
                    st=ss(get('STATUS')) or '—', prev=ss(get('PREVISÃO')),
                    obs=ss(at(row, header[obs_key])) if obs_key else '', ppa=ss(get('ENTREGA'))
                ))
        if not ini:
            raise ValueError('Nenhuma iniciativa válida encontrada. JSON anterior preservado.')
        return {'tot': tot, 'lme': lme, 'upd': upd, 'units': units, 'ini': ini}
    finally:
        wb.close()

if __name__ == '__main__':
    origem = Path(__file__).resolve().parent
