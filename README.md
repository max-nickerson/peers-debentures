# peers-debentures

Debêntures dos peers, com look-through (abre cotas de fundos): os 12 meses mais recentes com a carteira toda aberta (os ~3 últimos, ainda confidenciais, ficam de fora). Dados públicos da CVM (carteira mensal CDA).

```
pip install pandas requests openpyxl
python peers_debentures.py
```

Gera `peers_debentures.xlsx` ao lado do arquivo: `mes | fundo | cnpj | codigo | financeiro | perc_pl`.

- Lista de peers: `PEERS` no topo do arquivo. Meses abertos: `MESES = 12`.
- Aba `meses`: % sem código e peers sem carteira em cada mês.
- 1ª vez baixa ~300 MB da CVM (~40 s); depois só o mês que a CVM republicar (~8 s).
- Quando a CVM abre um mês (republica o arquivo), o script baixa de novo sozinho e ele entra na janela.
- Linha sem código na CVM: o código vem da SND pelo CNPJ do emissor + vencimento (`A/B` = séries gêmeas, mesmo vencimento e cupom).
  O que nem a SND tem (emissão privada) fica `sem codigo: emissor vencimento` (<0,4% do financeiro).
