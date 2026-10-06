# Staffing rota app

Netlify front end for the café staffing rota. The Google Sheet is the current database.

## Google Apps Script setup

1. Open the Google Sheet that contains **Colleague**, **Rostering** and **Availability**.
2. Extensions → Apps Script.
3. Create/replace `Code.gs` with the repository file `apps-script/Code.gs`.
4. If the script is bound to that spreadsheet, leave `spreadsheetId` blank. If it is a standalone Apps Script project, put the Google Sheet ID into `CONFIG.spreadsheetId`.
5. Deploy → New deployment → Web app.
6. Execute as: **Me**.
7. Choose the access setting appropriate for the devices that will use the rota app.
8. Copy the deployed `/exec` URL.

The script automatically creates these database tabs if missing:
- RegularShifts
- Absences
- Rotas
- RotaAssignments

### Current spreadsheet mapping

**Colleague:** Name, Contract, Till number, Morning Cook, Evening Cook, Till, Plate up, Wash up, Diningroom, Open up, Close down.

**Availability:** existing two-row day header layout, with colleague name in column A and From/Until pairs for Monday–Sunday. The existing `Ftriday` typo does not stop the API because the positions are mapped explicitly.

**Rostering:** existing required shift template. Blank shift times mean that shift is not required that day.

### Rules implemented server-side for saved assignments

- Till Number is the colleague ID.
- Required shift definitions come from Rostering.
- More than 6 elapsed hours = 15 minute unpaid break.
- Paid hours = elapsed hours minus the break.
- Budget cost = PaidHours × HourlyRate.
- Holiday and Sickness records are stored in Absences; blank sickness EndDate can represent ongoing absence.

## API

GET `?action=bootstrap` returns colleagues, availability, required shifts, regular shifts and absences.

POST JSON:
- `{ action: "saveAbsence", ... }`
- `{ action: "saveRota", ... }`

The Netlify front end still needs its API URL switched from the temporary JSON fixture to the Apps Script `/exec` URL after deployment.
