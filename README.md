# PDF invoices → Excel rows

> **For non-technical clients — start here**
>
> **What it does:** Turns a folder of invoice PDFs into one clean Excel spreadsheet (vendor, date, total, currency…).
>
> **Try it (3 steps):**
> 1. Download this repo (or clone it).
> 2. Open a terminal in the folder and run: `pip install -r requirements.txt` then `python extract_invoices.py`
> 3. Open `output/invoices.xlsx` — your rows are ready.
>
> **Business benefit:** stop retyping invoices into Sheets. A batch that takes finance 30–60 minutes becomes **under a minute**. Client work from **$80**.

![Demo](docs/demo.gif)

**Samples:** [`samples/`](./samples/) (3 fictional invoices) · **Output:** [`output/invoices.xlsx`](./output/invoices.xlsx)  
**Brand:** [Nyyo Automation](https://github.com/nyyoautomation)

[![Python](https://img.shields.io/badge/Python-3776AB?style=flat&logo=python&logoColor=white)](#)
[![pandas](https://img.shields.io/badge/openpyxl-Excel-150458?style=flat)](#)
[![PDF](https://img.shields.io/badge/PDF-extract-red?style=flat)](#)

---

## What you get

- 3 sample invoices (PDF)
- Script that extracts: file, vendor, invoice #, dates, currency, subtotal, VAT, total
- Excel + CSV output with a Summary sheet by currency

## Run

```bash
pip install -r requirements.txt
python extract_invoices.py
# optional: python extract_invoices.py --samples ./samples --xlsx ./output/invoices.xlsx
```

## What’s for sale

| | |
|---|---|
| **Product** | Batch invoice / PDF → Excel tailored to your layout |
| **From** | **$80** |
| **Includes** | Script + sample run + 1 revision |

Scanned (image) PDFs can use OCR in a client build — this public demo uses digital PDFs (no paid API).

→ [Packages & contact](https://github.com/nyyoautomation) · [LaborX — Nyyo Automation](https://laborx.com/freelancers/users/id475124)

## Hire me

1. Open an **issue** with: goal · deadline · budget  
2. Or message **[LaborX — Nyyo Automation](https://laborx.com/freelancers/users/id475124)**

Crypto: USDC / SOL / ETH / BTC.

## Limits (honest)

- Sample PDFs are **fictional**. Layouts in the wild vary — production extractors are tuned per vendor template.
- Image-only scans need OCR (available on request).
