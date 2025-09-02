
# Excel Merge Service (FastAPI + openpyxl)

Servizio che prende **template originale (.xlsx/.xlsm)** e **ugly.xlsx** e restituisce un file finale
con gli **stili/merge/validazioni/macro** del template, copiando **solo i valori**.

## Endpoints
- `GET /` → health check
- `POST /merge` (multipart form):
  - `template_file`: file .xlsx o .xlsm (template originale pre-conversione)
  - `ugly_file`: file .xlsx contenente i dati corretti
  - `config_json`: JSON con mappature:
    ```json
    {
      "single_fields_by_header": { "Numero Fattura": ["Fattura","B5"] },
      "table_mappings": {
        "Righe": {
          "sheet_target": "Fattura",
          "header_row": 1,
          "start_row_target": 12,
          "start_col_target": 2,
          "columns": {
            "Descrizione": "Descrizione",
            "Quantità": "Quantità",
            "Prezzo": "Prezzo",
            "IVA %": "IVA %",
            "Totale": "Totale"
          },
          "max_rows": null
        }
      },
      "single_fields_by_cell": { "Riepilogo!B3": ["Fattura","E9"] }
    }
    ```
"# EXCELL-COMPILAZIONE" 
