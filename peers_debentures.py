"""
Debentures dos peers com look-through (abre cotas de fundos). Fonte: CVM, carteira mensal (CDA), publica.
12 meses com a carteira toda aberta. Os meses recentes (confidenciais; a CVM abre ~90 dias depois) ficam de fora.
Uso:    python peers_debentures.py
Saida:  peers_debentures.xlsx ao lado deste arquivo, uma linha por mes x fundo x debenture:
        mes | fundo | cnpj | codigo | financeiro (R$) | perc_pl (% do PL);  aba "meses": % sem codigo e peers sem carteira
        Linha sem codigo na CVM: codigo achado na SND (debentures.com.br) por CNPJ do emissor + vencimento.
Dados:  dados_peers_debentures/ guarda um resumo pequeno por mes; so baixa de novo o mes que a CVM republicar.
Requer: pandas, requests, openpyxl.
"""
import io
import zipfile
from concurrent.futures import ProcessPoolExecutor
from datetime import date
from pathlib import Path
import re

import pandas as pd
import requests

PEERS = {                       # se um CNPJ tambem estiver em NOSSOS_FUNDOS, vale como nosso
    "XP Corporate Top Credito": "04.621.721/0001-70",
    "Valora Absolute": "10.326.625/0001-00",
    "XP Corporate Light": "11.046.179/0001-34",
    "Principal Claritas": "11.447.136/0001-60",
    "Sparta Top": "14.188.162/0001-00",
    "Plural Credito Corporativo": "18.316.558/0001-46",
    "Sparta Max": "26.773.148/0001-52",
    "Itau Multimercado Credito Privado": "28.840.420/0001-03",
    "AF Invest Geraes 30": "29.044.189/0001-04",
    "JGP Corporate Plus": "32.892.264/0001-93",
    "Dual Advanced": "34.803.938/0001-61",
    "Compass Yield 30": "36.318.479/0001-56",
    "Itau High Yield All": "42.263.927/0001-64",
    "Icatu Vanguarda": "42.501.967/0001-05",
    "ARX Vinson": "42.698.327/0001-29",
    "Itau Sinfonia Multimercado": "42.717.960/0001-17",
    "Occam Credito Corporativo": "47.586.648/0001-55",
    "Absolute Creta Selecao": "48.094.354/0001-79",
    "Mag High Grade Plus 30": "50.697.486/0001-37",
    "Vinland Credito Selecao": "50.980.211/0001-06",
    "Itau Sinfonia All": "54.278.393/0001-29",
    "Solis Capital Antares": "13.054.728/0001-48",
    "Capitania Premium 45": "20.146.294/0001-71",
    "AZ Quest Altro": "22.100.009/0001-07",
    "XP Corporate Plus": "23.999.611/0001-90",
    "ARX Everest": "32.102.131/0001-76",
    "Precision Advanced": "32.292.528/0001-78",
    "JGP Select": "32.892.615/0001-66",
    "AZ Quest Supra": "36.352.498/0001-07",
    "Root Capital Credito HG": "42.405.028/0001-59",
    "SPX Seahawk": "42.431.531/0001-89",
    "AF Horizonte": "44.025.131/0001-07",
    "Sparta Max 60": "44.643.192/0001-20",
    "Absolute Atenas Itau": "49.645.368/0001-04",
    "ARX RF": "41.575.611/0001-45",
    "ARX K2 IVP Inflacao Curta": "46.997.356/0001-42",
    "BTG Pactual Credito Corporativo": "42.827.247/0001-26",
    "Capitania Itau": "41.709.507/0001-04",
    "Capitania Idence Itau": "42.827.631/0001-29",
    "Icatu Vanguarda Absoluto II": "47.212.476/0001-50",
    "Icatu Vanguarda A": "34.781.249/0001-01",
    "Itau Active Fix": "41.301.133/0001-85",
    "Itau High Yield II": "42.860.483/0001-44",
    "JGP Credito Itau": "41.955.494/0001-45",
    "Kinea FI RF": "26.491.419/0001-87",
    "Porto Credito RF": "54.974.029/0001-01",
    "Schroder Idencia Itau I": "42.535.150/0001-40",
    "Sparta Inflacao IU": "43.737.649/0001-00",
    "Sparta IU": "46.997.384/0001-60",
    "SPX Seahawk Itau": "42.014.260/0001-66",
    "SulAmerica Cred ESG Itau": "45.615.787/0001-34",
    "Vinland Credito A T1 MM": "49.456.416/0001-08",
    "Absolute Delfos Itau": "49.645.692/0001-14",
    "Capitania Reit Itau": "42.934.005/0001-31",
    "Ibiuna I Credit Itau": "45.644.076/0001-98",
    "Icatu Vanguarda A Qualificado": "41.867.248/0001-31",
    "Itau BTG Pactual CorpPlus": "41.735.237/0001-06",
    "Itau Flexprev Advanced": "49.803.664/0001-88",
    "Itau Dual Prev": "58.156.912/0001-37",
    "Itau Sinfonia MM": "42.380.882/0001-08",
    "Vinland Credito IQ": "61.735.868/0001-24",
}

MESES = 12                      # meses com a carteira aberta; os recentes (ainda confidenciais, ~90 dias) ficam de fora
AQUI = Path(__file__).resolve().parent
DADOS = AQUI / "dados_peers_debentures"
OUT = AQUI / "peers_debentures.xlsx"
URL = "https://dados.cvm.gov.br/dados/FI/DOC/CDA/DADOS/"
SND = "https://www.debentures.com.br/exploreosnd/consultaadados/emissoesdedebentures/caracteristicas_e.asp?tip_deb=publicas&ativo=&op_exc="
VERSAO = 2                      # formato do resumo mensal (mudou = refaz os resumos)
REN = {"CNPJ_FUNDO": "CNPJ_FUNDO_CLASSE", "CNPJ_FUNDO_COTA": "CNPJ_FUNDO_CLASSE_COTA"}     # nomes antigos da CVM


def le(z, nome, cols):
    return pd.read_csv(z.open(nome), sep=";", encoding="latin1", dtype=str, quoting=3,
                       usecols=lambda c: c in cols or c in REN).rename(columns=REN)


def resumo(ym, versao):
    """Baixa o zip do mes e guarda so o necessario de TODOS os fundos: PL, cotas de fundos e debentures."""
    f = DADOS / f"{ym}_{versao}_v{VERSAO}.pkl"
    if f.exists():
        return f
    r = requests.get(f"{URL}cda_fi_{ym}.zip", timeout=900)
    r.raise_for_status()
    z = zipfile.ZipFile(io.BytesIO(r.content))
    junta = lambda chaves, cols: pd.concat([le(z, n, cols).assign(conf="CONFID" in n) for n in z.namelist()
                                            if any(k in n for k in chaves)], ignore_index=True)
    pl = junta(["_PL_"], {"CNPJ_FUNDO_CLASSE", "VL_PATRIM_LIQ"}).drop_duplicates("CNPJ_FUNDO_CLASSE")
    pl = pl.set_index("CNPJ_FUNDO_CLASSE").VL_PATRIM_LIQ.astype(float)
    cota = junta(["_BLC_2_"], {"CNPJ_FUNDO_CLASSE", "CNPJ_FUNDO_CLASSE_COTA", "VL_MERC_POS_FINAL"}).dropna()
    cota = cota.rename(columns={"CNPJ_FUNDO_CLASSE": "pai", "CNPJ_FUNDO_CLASSE_COTA": "filho", "VL_MERC_POS_FINAL": "v"})
    d = junta(["_BLC_4_", "_BLC_6_", "_BLC_8_", "CONFID"], {"CNPJ_FUNDO_CLASSE", "TP_APLIC", "VL_MERC_POS_FINAL", "CD_ATIVO",
              "DS_ATIVO", "EMISSOR", "CPF_CNPJ_EMISSOR", "DT_VENC", "PR_CUPOM_POSFX"})
    d = d[d.TP_APLIC.fillna("").str.startswith("Deb")]                  # Debentures / Debentures conversiveis (nao compromissada)
    t = lambda c: d[c].fillna("").str.strip() if c in d else pd.Series("", index=d.index)
    x = t("DS_ATIVO").str.extract(r"(\d{2}/\d{2}/\d{4}) / (\d{14})")    # BLC_8: "DEBENTURES / venc / CNPJ do emissor"
    d["codigo"] = t("CD_ATIVO").str.upper().where(~d.conf, "(confidencial)")
    sem = d.codigo.eq("")                                                # sem codigo: guarda emissor, vencimento e cupom p/ achar na SND
    d["cnpj_em"] = t("CPF_CNPJ_EMISSOR").str.replace(r"\D", "", regex=True).where(sem, "")
    d["cnpj_em"] = d.cnpj_em.where(d.cnpj_em != "", x[1].fillna("").where(sem, ""))
    d["venc"] = pd.to_datetime(t("DT_VENC"), errors="coerce").fillna(pd.to_datetime(x[0], format="%d/%m/%Y", errors="coerce"))
    d["venc"] = d.venc.dt.strftime("%Y-%m-%d").fillna("").where(sem, "")
    d["cupom"] = pd.to_numeric(t("PR_CUPOM_POSFX").str.replace(",", "."), errors="coerce").round(4).fillna(-1).where(sem, -1)
    d["emissor"] = t("EMISSOR").where(t("EMISSOR") != "", t("DS_ATIVO")).where(sem, "")
    d = d.rename(columns={"CNPJ_FUNDO_CLASSE": "fundo", "VL_MERC_POS_FINAL": "v"})
    for y in (cota, d):
        y["v"] = pd.to_numeric(y.v, errors="coerce").fillna(0)
    d = d.groupby(["fundo", "codigo", "cnpj_em", "venc", "cupom", "emissor"], as_index=False).v.sum()
    for velho in DADOS.glob(f"{ym}_*.pkl"):                              # CVM republicou (ex.: abriu as confidenciais): troca o resumo
        velho.unlink()
    pd.to_pickle((pl, cota, d), f)
    return f


def snd():
    """Cadastro de todas as debentures da SND (ativas e vencidas): codigo, CNPJ do emissor, vencimento, taxa. 1x por dia."""
    f = DADOS / f"snd_{date.today()}.pkl"
    if f.exists():
        return pd.read_pickle(f)
    partes = []
    try:
        for exc in ("False", "True"):
            t = requests.get(SND + exc, headers={"User-Agent": "Mozilla/5.0"}, timeout=300).content.decode("latin1").splitlines()
            i = next(j for j, l in enumerate(t) if l.startswith("Codigo do Ativo"))
            partes.append(pd.read_csv(io.StringIO("\n".join(t[i:])), sep="\t", dtype=str, quoting=3))
    except (requests.RequestException, StopIteration):
        print("aviso: SND (debentures.com.br) fora do ar; linhas sem codigo ficam 'sem codigo: ...'")
        return pd.DataFrame(columns=["codigo", "cnpj_em", "venc", "juros"])
    a = pd.concat(partes, ignore_index=True)
    a.columns = a.columns.str.strip()
    a = a.apply(lambda s: s.str.strip())
    s = pd.concat([pd.DataFrame({"codigo": a["Codigo do Ativo"], "cnpj_em": a.CNPJ.str.replace(r"\D", "", regex=True),
                                 "venc": pd.to_datetime(a[c], format="%d/%m/%Y", errors="coerce").dt.strftime("%Y-%m-%d"),
                                 "juros": pd.to_numeric(a["Juros Criterio Novo - Taxa"].str.replace(",", "."), errors="coerce")})
                   for c in ("Data de Vencimento", "Data de Saida / Novo Vencimento")]).dropna(subset=["venc"]).drop_duplicates()
    for velho in DADOS.glob("snd_*.pkl"):
        velho.unlink()
    s.to_pickle(f)
    return s


def nomeia(deb, s):
    """Linha sem codigo -> codigo da SND pelo CNPJ do emissor + vencimento (empate: mesmo cupom). Varias series: 'A/B'."""
    sem = deb.codigo.eq("")
    m = deb.loc[sem, ["cnpj_em", "venc", "cupom"]].drop_duplicates().merge(s, on=["cnpj_em", "venc"])
    m["bate"] = (m.juros - m.cupom).abs().lt(0.01)
    m = m[m.bate | ~m.groupby(["cnpj_em", "venc", "cupom"]).bate.transform("any")]
    nome = m.groupby(["cnpj_em", "venc", "cupom"], as_index=False).codigo.agg(lambda c: "/".join(sorted(set(c))))
    achou = deb.loc[sem, ["cnpj_em", "venc", "cupom"]].merge(nome, how="left").codigo.values
    achou = pd.Series(achou, index=deb.index[sem])
    deb.loc[sem, "codigo"] = achou.fillna("sem codigo: " + (deb.emissor + " " + deb.venc).str.strip()[sem])
    return deb.groupby(["fundo", "codigo"], as_index=False).v.sum()


def explode(pl, cota, deb, peers):
    """f = fracao de cada veiculo que pertence ao peer (1 no proprio peer; cota/PL do fundo investido nos seguintes).
    Cotas de FIDC/FII/FIAGRO nao abrem (esses fundos nao mandam CDA) - e nao sao debenture."""
    front = pd.DataFrame({"peer": peers, "veic": peers, "f": 1.0})
    achou = []
    for _ in range(10):                                                  # 10 niveis de FIC dentro de FIC e mais que suficiente
        if front.empty:
            break
        achou.append(front.merge(deb, left_on="veic", right_on="fundo").assign(fin=lambda x: x.v * x.f))
        nx = front.merge(cota, left_on="veic", right_on="pai")
        nx["f"] = nx.f * nx.v / nx.filho.map(pl)                          # fundo investido sem carteira na CVM: some (NaN)
        front = nx[nx.f > 0].groupby(["peer", "filho"], as_index=False).f.sum().rename(columns={"filho": "veic"})
    x = pd.concat(achou, ignore_index=True).groupby(["peer", "codigo"], as_index=False).fin.sum()
    x = x[x.fin.round(2) != 0]
    return x.assign(perc_pl=x.fin / x.peer.map(pl) * 100)


def main():
    DADOS.mkdir(exist_ok=True)
    html = requests.get(URL, timeout=60).text
    pub = dict(re.findall(r"cda_fi_(\d{6})\.zip</a>\s+(\d{2}-\w{3}-\d{4} \d{2}:\d{2})", html))
    todos = sorted(pub)
    s = snd()
    cnpj = {re.sub(r"\D", "", c): n for n, c in PEERS.items()}
    fmt = lambda c: f"{c[:2]}.{c[2:5]}.{c[5:8]}/{c[8:12]}-{c[12:]}"
    feitos, meses = {}, todos[-MESES:]
    print("CVM: baixando o que falta (cada mes ~25 MB, so na 1a vez ou quando a CVM republica)...")
    while True:                     # volta no tempo ate ter MESES meses abertos (a CVM abre as confidenciais ~90 dias depois)
        falta = [m for m in meses if m not in feitos]
        with ProcessPoolExecutor(4) as ex:
            for m, f in zip(falta, ex.map(resumo, falta, [re.sub(r"\W", "", pub[m]) for m in falta])):
                pl, cota, deb = pd.read_pickle(f)
                peers = [fmt(c) for c in cnpj if fmt(c) in pl.index]
                x = explode(pl, cota, nomeia(deb, s), peers).assign(mes=f"{m[:4]}-{m[4:]}")
                conf = x.fin[x.codigo.eq("(confidencial)")].sum() / x.fin.sum() * 100
                feitos[m] = x, conf, [n for c, n in cnpj.items() if fmt(c) not in pl.index]
        abertos = [m for m in meses if feitos[m][1] < 20]
        if len(abertos) >= MESES or meses[0] == todos[0]:
            break
        meses = todos[-(len(meses) + MESES - len(abertos)):]
    df = pd.concat([feitos[m][0] for m in abertos[-MESES:]], ignore_index=True)
    df = df.assign(fundo=df.peer.str.replace(r"\D", "", regex=True).map(cnpj), cnpj=df.peer, financeiro=df.fin.round(2))
    df = df[["mes", "fundo", "cnpj", "codigo", "financeiro", "perc_pl"]].sort_values(
        ["mes", "fundo", "financeiro"], ascending=[False, True, False])
    g = df.groupby("mes")
    resumo_mes = pd.DataFrame({
        "peers_com_carteira": g.fundo.nunique(),
        "financeiro_bi": g.financeiro.sum() / 1e9,
        "pct_sem_codigo": df.financeiro.where(df.codigo.str.startswith("sem codigo"), 0).groupby(df.mes).sum() / g.financeiro.sum() * 100,
        "peers_sem_carteira": [", ".join(feitos[m.replace("-", "")][2]) for m in g.groups]}).sort_index(ascending=False).round(2)
    with pd.ExcelWriter(OUT) as w:
        df.to_excel(w, sheet_name="debentures", index=False)
        resumo_mes.to_excel(w, sheet_name="meses")
    print(resumo_mes.drop(columns="peers_sem_carteira").to_string())
    print(f"salvo {OUT.name}: {len(df):,} linhas, {df.fundo.nunique()} fundos, {df.mes.nunique()} meses abertos "
          f"({df.mes.min()} a {df.mes.max()}); fora, ainda confidenciais: {', '.join(m for m in meses if m not in abertos)}")


if __name__ == "__main__":
    main()
