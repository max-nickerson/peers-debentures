# peers-debentures

Debêntures dos peers, com look-through (abre cotas de fundos), últimos 12 meses. Dados públicos da CVM (carteira mensal CDA).

```
pip install pandas requests openpyxl
python peers_debentures.py
```

Gera `peers_debentures.xlsx` ao lado do arquivo: `mes | fundo | cnpj | codigo | financeiro | perc_pl`.

- Lista de peers: `PEERS` no topo do arquivo. Meses: `MESES = 12`.
- 1ª vez baixa ~300 MB da CVM (~30 s); depois só o mês que a CVM republicar (~8 s).
- `(confidencial)`: debêntures que o fundo ainda não abriu (a CVM abre em até 90 dias).
- `sem codigo: emissor vencimento`: linhas que a CVM informa sem código (<1% do financeiro).
