import openpyxl, re, unicodedata

PATH = "data/producao2026_v2.xlsx"
wb = openpyxl.load_workbook(PATH, data_only=True)

MONTHS = ["JANEIRO","FEVEREIRO","MARÇO","ABRIL","MAIO","JUNHO","JULHO",
          "AGOSTO","SETEMBRO","OUTUBRO","NOVEMBRO","DEZEMBRO"]

def norm(s):
    if s is None: return ""
    s = str(s).strip()
    return s

def is_number(v):
    if v is None: return False
    if isinstance(v, (int, float)): return True
    try:
        float(str(v).replace(",", "."))
        return True
    except:
        return False

print("ABAS:", wb.sheetnames)

for name in wb.sheetnames:
    ws = wb[name]
    print("\n" + "="*95)
    print(f"ABA: {name}  (max_row={ws.max_row}, max_col={ws.max_column})")
    print("="*95)

    # find header row (the one containing 'PROCEDIMENTO')
    header_row = None
    for r in range(1, 12):
        for c in range(1, ws.max_column+1):
            v = norm(ws.cell(r,c).value).upper()
            if v.startswith("PROCEDIMENTO"):
                header_row = r
                break
        if header_row: break
    print(f"[header] linha de cabecalho detectada: R{header_row}")

    # collect headers
    headers = []
    for c in range(1, ws.max_column+1):
        v = norm(ws.cell(header_row, c).value)
        if v != "":
            headers.append((c, v))
    print("[header] colunas:", [h[1] for h in headers])

    # ---- CHECK 1: TAB header naming pattern ----
    tab_issues = []
    for c, h in headers:
        hu = h.upper()
        if "TAB" in hu:
            # expected exactly MONTH/TAB
            if not re.match(r"^(JANEIRO|FEVEREIRO|MARÇO|ABRIL|MAIO|JUNHO|JULHO|AGOSTO|SETEMBRO|OUTUBRO|NOVEMBRO|DEZEMBRO)/TAB$", hu):
                tab_issues.append(h)
    if tab_issues:
        print("  [!] cabeçalhos TAB fora do padrão 'MÊS/TAB':", tab_issues)
    else:
        print("  [ok] cabeçalhos TAB no padrão (ou sem TAB)")

    # ---- CHECK 2: non-numeric values in data cells ----
    # data rows = from header_row+1 until a row whose first cell starts with 'Total'
    data_start = header_row + 1
    total_row = None
    last_proc_row = ws.max_row
    for r in range(data_start, ws.max_row+1):
        first = norm(ws.cell(r,1).value).upper()
        if first.startswith("TOTAL PROCEDIMENTOS") or first.startswith("TOTAL DE PROCEDIMENTOS"):
            total_row = r
            last_proc_row = r-1
            break

    print(f"  [info] linhas de procedimento: R{data_start}..R{last_proc_row}; linha total: R{total_row}")

    # month -> column map (non-tab month columns)
    monthcols = {}
    tabcols = {}
    for c, h in headers:
        hu = h.upper()
        if hu in MONTHS:
            monthcols[hu] = c
        elif "/TAB" in hu:
            m = hu.split("/")[0]
            tabcols[m] = c

    # non-numeric check across proc rows
    nonnum = []
    for r in range(data_start, last_proc_row+1):
        label = norm(ws.cell(r,1).value)
        if label == "": continue
        for c, h in headers:
            if h.upper() == "TOTAL" or h.upper().startswith("PROCEDIMENTO"): continue
            v = ws.cell(r,c).value
            if v is None: continue
            if not is_number(v):
                nonnum.append((r, h, repr(v), label[:40]))
    if nonnum:
        print("  [!] valores NÃO numéricos em células de dados:")
        for r,h,v,lbl in nonnum:
            print(f"       R{r} col '{h}' = {v}  (proc: {lbl})")
    else:
        print("  [ok] nenhum valor não-numérico nas células de dados")

    # ---- CHECK 3: recalculated TOTAL (row sum) vs typed TOTAL ----
    # typed total column
    total_col = None
    for c, h in headers:
        if h.upper() == "TOTAL":
            total_col = c
    if total_col:
        bad_rowtotals = []
        for r in range(data_start, last_proc_row+1):
            label = norm(ws.cell(r,1).value)
            if label == "": continue
            s = 0.0
            for m, c in monthcols.items():
                v = ws.cell(r,c).value
                if is_number(v): s += float(str(v).replace(",","."))
            typed = ws.cell(r, total_col).value
            if is_number(typed):
                if abs(float(typed) - s) > 0.5:
                    bad_rowtotals.append((r, label[:45], s, float(typed)))
        if bad_rowtotals:
            print("  [!] TOTAL digitado difere da soma dos meses (linha a linha):")
            for r,lbl,calc,typed in bad_rowtotals:
                print(f"       R{r} '{lbl}': soma meses={calc:.0f}  vs  TOTAL digitado={typed:.0f}")
        else:
            print("  [ok] TOTAIS por linha batem com a soma dos meses")

    # ---- CHECK 4: empty month-total cells on the total row ----
    if total_row:
        empties = []
        for m, c in monthcols.items():
            v = ws.cell(total_row, c).value
            if v is None or norm(v)=="":
                # only flag if that month has data above
                has_data = any(is_number(ws.cell(r,c).value) and float(ws.cell(r,c).value)!=0
                               for r in range(data_start,last_proc_row+1))
                empties.append((m, has_data))
        flagged = [(m,hd) for m,hd in empties if hd]
        if flagged:
            print("  [!] 'Total procedimentos' VAZIO em meses que têm dados nas linhas:")
            for m,hd in flagged:
                print(f"       mês {m}: total vazio, mas há produção nas linhas")

    # ---- CHECK 5: TAB > Realizado (per month, per proc) ----
    tabgt = []
    for r in range(data_start, last_proc_row+1):
        label = norm(ws.cell(r,1).value)
        if label=="": continue
        for m, tc in tabcols.items():
            mc = monthcols.get(m)
            if mc is None: continue
            real = ws.cell(r, mc).value
            tab = ws.cell(r, tc).value
            if is_number(real) and is_number(tab):
                if float(tab) > float(real):
                    tabgt.append((m, label[:40], float(real), float(tab)))
    if tabgt:
        print(f"  [i] TAB > Realizado em {len(tabgt)} ocorrência(s) (DADO, não erro de formato) - exemplos:")
        for m,lbl,real,tab in tabgt[:6]:
            print(f"       {m}: '{lbl}' realizado={real:.0f} tab={tab:.0f}")

    # ---- CHECK 6: isolated zeros (0 flanked by big numbers) in main procedures ----
    for r in range(data_start, last_proc_row+1):
        label = norm(ws.cell(r,1).value)
        if "ATENDIMENTO INDIVIDUAL" not in label.upper(): continue
        seq = []
        for m in MONTHS:
            c = monthcols.get(m)
            if c is None: continue
            v = ws.cell(r,c).value
            if is_number(v): seq.append((m, float(v)))
        vals = [v for _,v in seq]
        if vals:
            mx = max(vals)
            for i,(m,v) in enumerate(seq):
                if v==0 and mx>50:
                    print(f"  [i] possível subnotificação: '{label[:40]}' mês {m} = 0 (máx do ano={mx:.0f})")
