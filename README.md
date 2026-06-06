# Domex Data Entry

Local data-entry app for creating Domex pickup rows and exporting them to the same Excel structure as `DataSheet.xlsx`.

The app is mobile-friendly and includes PWA metadata, so mobile browsers can open it in a standalone app-like view when added to the home screen.

## Run

```powershell
npm install
npm run dev
```

Open the local URL shown by Vite, usually:

```text
http://127.0.0.1:5173
```

## Excel Format

Exports one sheet named `Pickups` with these columns:

```text
TrackingNumber, Reference, PackageDescription, ReceiverName, ReceiverAddress,
ReceiverCity, ReceiverContactNo, NoOfPcs, Kilo, Gram, Amount, Exchange, Remark
```

Rows are saved in the browser on the same computer. Use **Export** to download the `.xlsx` file.

`Reference`, `PackageDescription`, `Exchange`, and `Remark` are hidden from the entry form. `Reference`, `Exchange`, and `Remark` export as `0`.

Use **Settings** to save the default `PackageDescription` value. That saved value is written to the Excel `PackageDescription` column for every exported row.

New entries start with a tracking step. Use **Scan Barcode** to read the tracking number with the camera when the browser supports barcode scanning, or type the tracking number manually in the same field. Use **Next** to continue to receiver details and **Back** to return to the tracking step.

Receiver City is selected from the city list in `domex_branches.json`. Type to search, then select a city from the result list. Manual city values outside that list are not accepted. All visible entry fields are required before a row can be added.

## Firebase Login and Upload

The app includes the Firebase config for `dome-d-entry` and uses Firebase Realtime Database paths:

```text
users
submissions
```

Admin login:

```text
username: madu
password: 2006
```

Admin can add and delete user accounts. Users log in with the username/password created by admin, enter pickup rows locally, then press **Upload**. Uploaded rows appear in the admin notification panel. Admin can export each upload to the same old Excel `Pickups` sheet format.

User entries are auto-saved to Firebase Realtime Database under `drafts/{username}` whenever rows are added, edited, deleted, imported, or settings refresh the rows. After upload, entries stay saved until the user presses **Clear Entry** in the upload success popup.

Note: this is a client-side username/password implementation for local operations. For production security, protect Firebase Realtime Database with proper Firebase Auth, Database Rules, or Cloud Functions.
