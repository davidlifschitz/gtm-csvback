# CsvBack

Apple saved the CSV as a Numbers file. Drop it here. Get a real `.csv` back. Files never leave the browser.

Ask this answers: [The worst UX of all time is Apple downloading csv files as numbers files so you have to then convert the file back to a csv](https://x.com/cbwritescopy/status/2090457086479995203)

- No account
- No upload
- Cap: 8 MB per file, 8 files
- Reads mislabeled CSV, Excel, ODS, and older Numbers zips

## Local

Open `index.html` in a browser, or:

```bash
python3 -m http.server 4173
```

New-format Numbers documents are protobuf inside a zip. If the table cannot be read, export once as Excel from Numbers and drop the `.xlsx`.

## GTM

Reply to people who re-export a CSV from Numbers by hand before a mail tool will take it. Copy is in the page footer.
