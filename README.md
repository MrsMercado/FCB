# Camp Excel Dashboard

This is a small browser app for camp Excel files.

It lets you:

- open the app after entering the password `FCBadmin`
- upload an `.xlsx`, `.xls`, or `.csv` file
- choose the worksheet tab to use
- summarize shirt, shorts, and sock sizes
- enter current shirt, shorts, and socks stock and see the difference
- remember stock numbers in the same browser
- print the inventory or save it as PDF
- download the inventory as CSV
- sort kids by date of birth
- create groups of 12
- sort each finished group by player last name
- create lettered groups, such as Group A, Group B, and Group C
- assign each group to Camp 1, Camp 2, or Camp 3
- add FCB coach and Aux. coach names per group
- print or download CSV rosters by the selected camp
- download a new Excel workbook with a dashboard and group sheets

## Run Locally

Open `index.html` in a browser, or run a small local web server:

```bash
python3 -m http.server 4173
```

Then open:

```text
http://localhost:4173
```

## Put On GitHub

1. Create a new GitHub repository.
2. Upload these files and folders: `index.html`, `styles.css`, `app.js`, `vendor`, `README.md`, and `render.yaml`.
3. Commit the files.

## Deploy On Render

1. In Render, choose **New +** then **Static Site**.
2. Connect the GitHub repository.
3. Use these settings:
   - Build Command: leave blank
   - Publish Directory: `.`
4. Click **Create Static Site**.

The app runs fully in the browser. Uploaded Excel files are not sent to a server.
