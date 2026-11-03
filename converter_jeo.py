           if 'LIMITE DE MOVIMENTAÇÃO EMPENHO' in wb.sheetnames:
            lme_sheet = wb['LIMITE DE MOVIMENTAÇÃO EMPENHO']
            lme_values = next(
                lme_sheet.iter_rows(min_row=2, max_row=2, values_only=True),
                None)
            if lme_values is None:
                raise ValueError('Aba LIMITE DE MOVIMENTAÇÃO EMPENHO sem linha 2.')
            lme = {
                'limite': nn(at(lme_values, 0)),
                'aguard': nn(at(lme_values, 1)),
                'emp': nn(at(lme_values, 2)),
                'saldo': nn(at(lme_values, 6))
            }
        else:
            if len(consolidado) < 4:
                raise ValueError('Consolidado 2026 não possui a linha 4.')
            r = consolidado[3]
            lme = {
                'limite': nn(at(r, 1)),
                'aguard': nn(at(r, 2)),
                'emp': nn(at(r, 3)),
                'saldo': nn(at(r, 7))
            }

        upd = next(
            (ss(v) for r in consolidado for v in r
             if re.match(r'^Atualiza', ss(v), re.I)), '')

        units = []
        for row in consolidado:
            unidade = ss(at(row, 0)).upper()
            unidade = {
                'OUVIDORIA': 'OUV',
                'SECRETARIA EXECUTIVA': 'SE'
            }.get(unidade, unidade)
            if unidade in UNIDADES:
                units.append(dict(
                    u=unidade, name=unidade,
                    prog=nn(at(row, 8)), emp=nn(at(row, 11)),
                    liq=nn(at(row, 12)), pago=nn(at(row, 13)),
                    dot=nn(at(row, 4)), livre=nn(at(row, 7)),
                    saldo=nn(at(row, 14))
                ))

        present = {unit['u'] for unit in units}
        for unit in UNIDADES:
            if unit not in present:
                units.append(dict(
                    u=unit, name=unit, prog=0, emp=0, liq=0,
                    pago=0, dot=0, livre=0, saldo=0
                ))

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

                def get(prefix):
                    return pegar(row, header, prefix)

                name = ss(get('INICIATIVAS'))
                prog = nn(get('PROGRAMADO'))
                emp = nn(get('EMPENHADO'))
                if not (prog or emp) or name.upper().startswith('TOTAL'):
                    continue

                obs_key = next(
                    (k for k in header if k.startswith('OBSERVAÇÃO')), None)
                ini.append(dict(
                    u=unit, acao=sheet_name.split(' ')[0],
                    po=ss(get('PO')),
                    name=name or '(sem descrição) — ' +
                    (ss(get('PARCEIRO')) or ss(get('DIRETORIA')) or '—'),
                    sei=ss(get('SEI')), dir=ss(get('DIRETORIA')),
                    resp=ss(get('RESPONSÁVEL')),
                    inst=ss(get('INSTRUMENTO')),
                    siafi=ss(get('SIAFI')), parc=ss(get('PARCEIRO')),
                    ano=ss(get('ANO DE')), prog=prog, emp=emp,
                    liq=nn(get('LIQUIDADO')), pago=nn(get('PAGO')),
                    st=ss(get('STATUS')) or '—',
                    prev=ss(get('PREVISÃO')),
                    obs=ss(at(row, header[obs_key])) if obs_key else '',
                    ppa=ss(get('ENTREGA'))
                ))

        if not ini:
            raise ValueError(
                'Nenhuma iniciativa válida encontrada. JSON anterior preservado.')
        if len(units) != 6:
            raise ValueError('Conversão incompleta: unidades não identificadas.')
        return {'tot': tot, 'lme': lme, 'upd': upd,
                'units': units, 'ini': ini}
    finally:
        wb.close()


if __name__ == '__main__':
    origem = Path(__file__).resolve().parent
    resultado = converter(localizar_excel(origem))
    destino = origem / ARQUIVO_JSON
    conteudo = json.dumps(
        resultado, ensure_ascii=False, indent=2, default=str) + '\n'
    destino.write_text(conteudo, encoding='utf-8')
    print(f"Gerado {destino.name}: {len(resultado['ini'])} iniciativas; "
          f"{len(resultado['units'])} unidades; atualização: {resultado['upd']}")
