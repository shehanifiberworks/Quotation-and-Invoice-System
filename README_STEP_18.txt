SFG Website — Step 18 Cloud Sync fix

IMPORTANT: Website AND Google Apps Script must both be updated.

1. Back up the existing Google Apps Script Code.gs text to a .txt file.
2. Open the same existing Apps Script project that is linked to your deployed Web App.
3. Replace its Code.gs text with the separate SFG_Cloud_Sync_Step_18_Code.gs file supplied in the chat.
4. Save (Ctrl+S), then Deploy -> Manage deployments -> Edit existing Web App -> Version: New version -> Deploy.
5. Existing SFG_SHEET_ID and SFG_ACCESS_KEY Script Properties stay in place. Do not create a new Google Sheet, deployment or access key.
6. Update your GitHub Pages website with the new index.html AND cloud/cloud-client.js. You can replace the remaining unchanged files too. Ensure GitHub path is correct.
7. Reload the site (hard refresh), then Settings -> Google Sheets Cloud -> Test Connection -> Sync Saved Files.
8. Sync will download documents in small batches. Existing cloud files are not deleted.

If your web app is public, keep its access key secret; never commit the key into website source. The published website ZIP intentionally does not contain backend Code.gs.
