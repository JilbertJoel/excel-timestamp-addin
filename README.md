# Tracker Timestamp Excel Add-in

This Excel Web add-in watches the **Tracker** worksheet. When a user first enters a value in column **D**, it writes a permanent Dominican Republic timestamp in column **E**. Existing timestamps are never replaced.

## GitHub Pages setup

1. Upload all files and the `assets` folder to the root of this repository.
2. Open **Settings > Pages**.
3. Under **Build and deployment**, select **Deploy from a branch**.
4. Select branch **main**, folder **/(root)**, and choose **Save**.
5. Wait until `https://jilbertjoel.github.io/excel-timestamp-addin/taskpane.html` opens successfully.

## Excel Web setup

1. Open the workbook in Excel Web.
2. Go to **Home > Add-ins > More Settings > Upload My Add-in**.
3. Upload `manifest.xml` from your computer.
4. Open **Tracker Timestamp > Start timestamps** from the Home ribbon.
5. Wait until the task pane shows **Active**.
6. Enter a test value in column D on a blank row. Column E should update almost immediately.

Each editor must install and open the add-in once. For organization-wide use, deploy the same manifest centrally through Microsoft 365 Integrated Apps.

Keep the existing Power Automate flow enabled during testing. Since it only fills blank cells in column E, it can serve as a fallback without replacing timestamps created by this add-in.
