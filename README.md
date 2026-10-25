# GST Reconciliation Tool

**By Ebin Michael** · [Live tool](https://gst.ebinmichael.in/) · [Portfolio](https://ebinmichael.in/)

A browser-based utility for comparing a GSTR-2B workbook with a purchase-register workbook. It groups invoice records into matched, mismatched, purchase-register-only, and GSTR-2B-only results, then exports the comparison to Excel.

This repository is a cleaned-up copy of the public website source. It separates the HTML, CSS, and JavaScript, uses the current SheetJS standalone CDN URL, and renders spreadsheet values as text so workbook contents are not interpreted as HTML. These source changes are not yet deployed to the Hostinger site.

## Features

- Reads `.xlsx` and `.xls` workbooks in the browser
- Provides downloadable input templates in `Template/`
- Compares supplier GSTIN and normalized invoice number
- Flags invoice-value and taxable-value differences greater than ₹1
- Displays four reconciliation result groups
- Exports the results to `GST_Reconciliation_Report.xlsx`

## Run locally

This is a static website. Open `index.html` in a modern browser, or serve this folder with any static web server. The spreadsheet library is loaded from the SheetJS CDN, so an internet connection is needed unless the dependency is downloaded and served locally.

1. Download this repository.
2. Open `index.html` in a modern browser.
3. Use the sample templates in `Template/` to prepare test workbooks.
4. Select both workbooks, accept the on-page disclaimer, and run reconciliation.

## Data handling

In the published HTML source reviewed for this repository, workbook files are read with the browser `FileReader` API and no upload endpoint or `fetch`/`XMLHttpRequest` call was present. This describes the reviewed client-side source; review the code and hosting configuration before relying on it for sensitive files. Use synthetic or anonymized workbooks for demonstrations. Do not commit client records, tax files, credentials, or private business information.

## Important limitations

- Results are preliminary comparisons and must be verified against official records.
- Matching is based on supplier GSTIN and normalized invoice number.
- The current implementation flags invoice-value and taxable-value differences; it does not flag date or tax-component differences as mismatches.
- The matching logic should be reviewed with representative sample workbooks before being used for operational decisions.

This repository does not claim to be an official GST service or a substitute for professional review.

## License

No open-source license has been selected yet. Until a license is added, others should not assume they have permission to reuse or redistribute this code.

## Publish this project on GitHub

See [PUBLISH_TO_GITHUB.md](PUBLISH_TO_GITHUB.md) for beginner-friendly steps using GitHub's website.
