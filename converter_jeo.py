       ini = []
    for sheet_name, u in sheets:
        if sheet_name not in wb.sheetnames:
            continue
        h = None
        for r in sh(sheet_name):
            if at(r, 0) == "UNIDADE":
                h = {ss(v).upper(): i for i, v in enumerate(r) if v}
                continue
            if h is None:
                continue

            def g(prefix):
                for k, idx in h.items():
                    if k.startswith(prefix):
                        return at(r, idx)
                return None

            name, prog, emp = ss(g("INICIATIVAS")), nn(g("PROGRAMADO")), nn(g("EMPENHADO"))
            if not (prog or emp) or name.upper().startswith("TOTAL"):
                continue
            ok = next((k for k in h if k.startswith("OBSERVAÇÃO")), None)
            ini.append(dict(
                u=u, acao=sheet_name.split(" ")[0], po=ss(g("PO")),
                name=name or "(sem descrição) — " + (ss(g("PARCEIRO")) or ss(g("DIRETORIA")) or "—"),
                sei=ss(g("SEI")), dir=ss(g("DIRETORIA")),
                resp=ss(g("RESPONSÁVEL")), inst=ss(g("INSTRUMENTO")),
                siafi=ss(g("SIAFI")), parc=ss(g("PARCEIRO")),
                ano=ss(g("ANO DE")), prog=prog, emp=emp,
                liq=nn(g("LIQUIDADO")), pago=nn(g("PAGO")),
                st=ss(g("STATUS")) or "—", prev=ss(g("PREVISÃO")),
                obs=ss(at(r, h[ok])) if ok else "", ppa=ss(g("ENTREGA"))
            ))

    upd = next((ss(at(r, 11)) for r in c
                if re.match(r"^Atualiza", ss(at(r, 11)), re.I)), "")
    wb.close()
    return dict(
        tot=tot,
        lme=dict(limite=nn(at(l, 0)), aguard=nn(at(l, 1)),
                 emp=nn(at(l, 2)), saldo=nn(at(l, 6))),
        upd=upd, units=units, ini=ini
    )

if __name__ == "__main__":
    origem = Path(__file__).resolve().parent
    resultado = converter(origem / ARQUIVO_EXCEL)
    destino = origem / ARQUIVO_JSON
    if not resultado["ini"] or len(resultado["units"]) != 6:
        raise ValueError("Conversão incompleta: dados.json não será substituído")
    destino.write_text(json.dumps(resultado, ensure_ascii=False, indent=2, default=str) + "\n",
                       encoding="utf-8")
    print(f"Gerado {destino.name}: {len(resultado['ini'])} iniciativas; "
          f"{len(resultado['units'])} unidades; atualização: {resultado['upd']}")
