# CCV Jewellery Melt Dashboard — Version 1

A Cash Converters themed dashboard for comparing jewellery ticket prices with estimated refinery melt values.

## Version 1 features

- Upload `.xls` or `.xlsx` stocktake reports
- Select Palmerston North, New Plymouth, or Wanganui before upload
- Detect the `NOT FOUND DURING STOCKTAKE, ON SALE ACCORDING TO SYSTEM` section automatically
- Replace the selected store's stock list on each upload
- Ignore watches completely
- Read stock code, description, ticket price, shelf dates, metal and `GMS` weight
- Calculate days on sale from original shelf date
- Calculate gold and sterling silver melt values
- Automatically refresh NZD spot prices once daily, with saved-price fallback
- Manual gold and silver scenario pricing
- Adjustable red, orange, yellow and green thresholds
- Store, metal, priority and text filters
- Needs Review list for platinum, missing metal and missing weight

## Melt formulas

Gold:

`((97% × NZD gold spot per gram ÷ 24) × carat × full weight) × 1.15`

Sterling silver:

`NZD silver spot per gram × 92.5% × 97% × full weight × 1.15`

## Run locally

1. Install Node.js 22 or newer.
2. Open a terminal in the project folder.
3. Run `npm install`.
4. Run `npm run dev`.
5. Open the address shown in the terminal.

## Data storage

Version 1 stores uploaded stock, price settings and thresholds in the browser on the device being used. Uploading a new report replaces only the selected store's list.

## Private information

No `.env` file, credentials, API keys or uploaded stock reports are included in this project.
