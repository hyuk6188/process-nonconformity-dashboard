# Google Sheets sync setup

Google Spreadsheet:

https://docs.google.com/spreadsheets/d/1iKBXYvK-wJ43oVpANLpT5R8wqXvc2VbGjrGqCp6zk7o/edit

Current Apps Script web app:

https://script.google.com/a/macros/fursys.com/s/AKfycbyM6GxfYzyIBqLd3zPiPLfXBUtJ3xHJ9vHwhKUksHwTlt7D7MNJfWypKpji0J0Cwenq/exec

The dashboard is already configured to use this URL. Users must be signed in to a Google account that can access the web app.

## 1. Create the Apps Script web app

1. Open the spreadsheet above.
2. Go to `Extensions > Apps Script`.
3. Replace the default `Code.gs` content with `google-sheets-sync-apps-script.js`.
4. Click `Deploy > New deployment`.
5. Select type `Web app`.
6. Set `Execute as` to `Me`.
7. Set `Who has access` to `Anyone`.
8. Deploy and copy the web app URL.

## 2. Connect the dashboard

In both `index.html` and `공정부적합_대시보드.html`, replace:

```js
const SYNC_WEB_APP_URL = '';
```

with:

```js
const SYNC_WEB_APP_URL = 'YOUR_WEB_APP_URL';
```

After committing and pushing, the GitHub Pages dashboard will save new records, Action List items, and notes into Google Sheets. Other PCs will load the same Sheet data on page load and refresh from Sheets every 60 seconds.

## ?? ??? ??? ??

GitHub Pages?? `?? ?? ??` ?? `??? ??`? ??? `Google ?? ??` ??? ????. ? ??? ?? Google ?? ???? ??? ? ????? ?????? ?? ???? ???? Google Sheets ???? ?????.
