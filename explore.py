import json

data=json.load(open("data/dados_extraidos.json",encoding="utf-8"))
MONTHS=["JANEIRO","FEVEREIRO","MARÇO","ABRIL","MAIO","JUNHO","JULHO","AGOSTO","SETEMBRO","OUTUBRO","NOVEMBRO","DEZEMBRO"]
def mshort(m): return m[:3].title()

def rsum(p): return sum(p["realizado"].values())
def tsum(p): return sum(p["tab"].values())

print("="*95)
print("ANÁLISE EXPLORATÓRIA - REDE CAPS PETROLINA 2026")
print("="*95)

# ---------- 1. Período real por unidade ----------
print("\n[1] PERÍODO REAL (meses com produção > 0) e meses com TAB")
for u in data["units"]:
    real_months=[m for m in MONTHS if any(p["realizado"].get(m,0)>0 for p in u["procedures"])]
    tab_months=[m for m in MONTHS if any(p["tab"].get(m,0)>0 for p in u["procedures"])]
    print(f"  {u['name']:<38} real: {[mshort(m) for m in real_months]}")
    if u['has_tab']:
        print(f"  {'':38} tab : {[mshort(m) for m in tab_months]}")

# ---------- 2. Taxa de tabulação JUSTA (só meses com TAB) geral vs tabulaveis ----------
print("\n[2] TAXA DE TABULAÇÃO (apenas meses jan-jul, onde há TAB)")
TABP=["JANEIRO","FEVEREIRO","MARÇO","ABRIL","MAIO","JUNHO","JULHO"]
for u in data["units"]:
    if not u["has_tab"]: continue
    real_jj=sum(p["realizado"].get(m,0) for p in u["procedures"] for m in TABP)
    tab_jj =sum(p["tab"].get(m,0) for p in u["procedures"] for m in TABP)
    # tabulaveis = procedimentos que tem algum tab>0 no ano
    tabulaveis=[p for p in u["procedures"] if tsum(p)>0]
    real_tabulaveis=sum(p["realizado"].get(m,0) for p in tabulaveis for m in TABP)
    tab_tabulaveis =sum(p["tab"].get(m,0) for p in tabulaveis for m in TABP)
    g = 100*tab_jj/real_jj if real_jj else 0
    t = 100*tab_tabulaveis/real_tabulaveis if real_tabulaveis else 0
    print(f"  {u['name']:<38} geral={g:5.1f}%   tabulaveis={t:5.1f}%  ({len(tabulaveis)}/{len(u['procedures'])} proc tabulados)")

# ---------- 3. Procedimentos que NUNCA sao tabulados (sempre 0) ----------
print("\n[3] PROCEDIMENTOS NUNCA TABULADOS (tab sempre 0, mas realizado > 0) - por unidade CAPS")
for u in data["units"]:
    if not u["has_tab"]: continue
    nunca=[(p["code"],p["desc"][:45],rsum(p)) for p in u["procedures"] if tsum(p)==0 and rsum(p)>0]
    if nunca:
        print(f"  {u['name']}:")
        for c,d,r in sorted(nunca,key=lambda x:-x[2]):
            print(f"     {c} {d:<47} realizado={r:.0f} / tabulado=0")

# ---------- 4. Mix de procedimentos (top por volume, rede toda) ----------
print("\n[4] TOP PROCEDIMENTOS POR VOLUME (rede inteira, realizado)")
agg={}
for u in data["units"]:
    for p in u["procedures"]:
        key=f"{p['code']} {p['desc'][:40]}"
        agg[key]=agg.get(key,0)+rsum(p)
for k,v in sorted(agg.items(),key=lambda x:-x[1])[:10]:
    print(f"     {v:8.0f}  {k}")

# ---------- 5. Evolução mensal da rede (realizado) + tendencia ----------
print("\n[5] PRODUÇÃO MENSAL DA REDE (realizado, soma todas unidades)")
for m in MONTHS:
    tot=sum(p["realizado"].get(m,0) for u in data["units"] for p in u["procedures"])
    if tot>0:
        bar="#"*int(tot/200)
        print(f"     {mshort(m)}: {tot:7.0f} {bar}")

# ---------- 6. Concentração: quanto cada unidade representa ----------
print("\n[6] PARTICIPAÇÃO DE CADA UNIDADE NO TOTAL DA REDE (realizado)")
grand=sum(rsum(p) for u in data["units"] for p in u["procedures"])
for u in data["units"]:
    ur=sum(rsum(p) for p in u["procedures"])
    print(f"     {u['name']:<38} {ur:7.0f}  ({100*ur/grand:4.1f}%)")
print(f"     {'TOTAL REDE':<38} {grand:7.0f}")

# ---------- 7. TAB > Realizado ----------
print("\n[7] CASOS TAB > REALIZADO (ponto a investigar)")
cnt=0
for u in data["units"]:
    for p in u["procedures"]:
        for m in MONTHS:
            r=p["realizado"].get(m); t=p["tab"].get(m)
            if r is not None and t is not None and t>r:
                cnt+=1
print(f"     total de ocorrências: {cnt}")

# ---------- 8. Procedimentos zerados o ano todo ----------
print("\n[8] PROCEDIMENTOS SEM NENHUMA PRODUÇÃO NO ANO (realizado total = 0)")
for u in data["units"]:
    zero=[f"{p['code']}" for p in u["procedures"] if rsum(p)==0]
    if zero:
        print(f"     {u['name']:<38} {len(zero)} proc zerados: {zero}")

# ---------- 9. UAI quase inexistente ----------
print("\n[9] OBSERVAÇÃO UAI")
uai=[u for u in data["units"] if u["sheet"]=="UAI"][0]
print(f"     UAI realizado total no ano: {sum(rsum(p) for p in uai['procedures']):.0f}")
print("     (volume quase nulo - verificar subutilização vs subnotificação)")
