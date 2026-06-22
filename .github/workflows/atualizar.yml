name: Atualizar Painel SEMPE

on:
  schedule:
    - cron: '0 10 * * *'   # 07:00 horário de Brasília (UTC-3)
  workflow_dispatch:         # permite rodar manualmente pelo GitHub

jobs:
  gerar-json:
    runs-on: ubuntu-latest
    steps:
      - name: Checkout
        uses: actions/checkout@v4

      - name: Instalar Python
        uses: actions/setup-python@v5
        with:
          python-version: '3.11'

      - name: Instalar dependências
        run: pip install openpyxl

      - name: Gerar dados.json
        run: python scripts/gerar_json.py

      - name: Commit e push
        run: |
          git config user.name  "SEMPE Bot"
          git config user.email "bot@sempe.sp.gov.br"
          git add docs/dados.json
          git diff --cached --quiet || git commit -m "dados: atualização automática $(date -u '+%Y-%m-%d %H:%M') UTC"
          git push
