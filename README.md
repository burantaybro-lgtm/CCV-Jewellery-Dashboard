# CCV Jewellery Melt Dashboard — Version 3

A Cash Converters themed dashboard for comparing jewellery ticket prices with estimated refinery melt values.

## Version 3 features

- Upload `.xls` or `.xlsx` stocktake reports
- Select Palmerston North, New Plymouth, or Wanganui before upload
- Detect the `NOT FOUND DURING STOCKTAKE, ON SALE ACCORDING TO SYSTEM` section automatically
- Replace the selected store's stock list on each upload
- Ignore watches completely
- Read stock code, description, ticket price, shelf dates, metal and `GMS` weight
- Recover ticket prices stored at the start of the report's first description line
- Show the total ticket price for the items included by the current search and filters
- Calculate days on sale from original shelf date
- Calculate gold and sterling silver melt values
- Automatically refresh NZD spot prices once daily, with saved-price fallback
- Manual gold and silver scenario pricing
- Estimated net sale return after GST and selling costs
- Fixed New Zealand GST at 15%
- Editable selling-fee percentage, defaulting to 15%
- Melt-risk colours based on melt value versus net sale return
- Adjustable red, orange, yellow and green thresholds
- Store, metal, priority and text filters
- Needs Review list for platinum, missing metal and missing weight

## Melt formulas

Gold:

`((97% × NZD gold spot per gram ÷ 24) × carat × full weight) × 1.15`

Sterling silver:

`NZD silver spot per gram × 92.5% × 97% × full weight × 1.15`

Net sale return:

`(ticket price ÷ 1.15) − (ticket price × editable selling fee)`

Melt risk:

`melt value ÷ net sale return × 100`

## Run locally

1. Install Node.js 22 or newer.
2. Open a terminal in the project folder.
3. Run `npm install`.
4. Run `npm run dev`.
5. Open the address shown in the terminal.

## Data storage

Version 3 stores uploaded stock, price settings, selling fee and thresholds in the browser on the device being used. Uploading a new report replaces only the selected store's list.

## Private information

No `.env` file, credentials, API keys or uploaded stock reports are included in this project.
