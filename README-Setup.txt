# SAHS Days 2026 QR Attendance System

## Included files
- `Code.gs`: Google Apps Script server-side code
- `Index.html`: phone-camera scanning web interface

## Before setup
Create a new Google Spreadsheet named `SAHS Days 2026 — QR Attendance System`.

Open **Extensions → Apps Script** from that spreadsheet. Replace the contents of `Code.gs` with the provided file, then create an HTML file named exactly `Index` and paste in `Index.html`.

## Initial setup
1. In the Apps Script editor, select `setupSheets` and click Run.
2. Approve the requested spreadsheet permissions.
3. Return to the spreadsheet. The seven tabs and headers should now exist.
4. Paste participant data into `Participants`, starting in row 2:
   `ID Number | Full Name | Program | Section | Team`
5. Paste working committee data into `Committee`, starting in row 2:
   `ID Number | Full Name | Program | Section | Committee | Position`
6. Format ID Number columns as **Plain text** before pasting IDs, especially if IDs can contain leading zeroes.
7. Confirm that the master IDs are unique within each list. The same person may be in both lists; the app uses the selected group to look up their record.
8. The spreadsheet timezone is set to `Asia/Manila` by setup.

## Deploy
1. In Apps Script, click **Deploy → New deployment**.
2. Select type **Web app**.
3. Execute as: **Me** (the spreadsheet owner).
4. Who has access: choose the narrowest option that still allows all authorized scanners to use it. If using a broad access option, treat the URL as sensitive and do not publish it publicly.
5. Deploy, authorize if prompted, and open the web app URL on each scanning phone.
6. Allow camera access. Set the correct date, group, checkpoint, and a unique station ID on every device.

Use the deployed `/exec` URL for event scanning. Test with sample records before the event.

## Important behavior
- Participant checkpoints: Morning In, Morning Out, Afternoon In, Afternoon Out.
- Committee checkpoints: Committee In and Committee Out.
- Duplicate successful scans for the same ID/date/checkpoint are rejected.
- The first Committee In scan of the day is used for the food-stub arrival record. Arrival at or before 6:45 AM is marked `ELIGIBLE`; later arrivals are marked `LATE`.
- Food-stub claiming is tracked in the sheet separately. This starter version does not yet include a claim button in the web interface; staff can record `YES`, claim timestamp, and distributor in `Food Stub Monitoring`, or this can be added in a next iteration.
- Logs are the source of truth. Run `refreshParticipantSummary()` from the Apps Script editor after a test or event day to rebuild the `Participant Summary` tab from the participant list and recorded logs. Avoid running this during peak scanning.
- QR values are expected to be the ID number only. Manual ID entry is provided as a fallback.

## Operational notes
- The script uses a script lock to serialize scan writes and reduce duplicate races across simultaneous devices. At high volume, a single shared Apps Script project can still have throughput limits, so conduct a load test and have a manual fallback.
- A web app executing as the owner can write to the spreadsheet on behalf of users. Restrict access to authorized staff.
- Test the camera scan flow on the exact phone/browser combinations planned for the event. Camera access requires a secure context (the Apps Script deployed web app uses HTTPS).
- Keep a protected backup of the spreadsheet and avoid sorting/deleting rows in raw log tabs during the event.
