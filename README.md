# W&L Athletics Van Reservations

A self-service reservation page for the Athletics vehicle fleet: Van 97, Van 140, Van 203, Van 206, Van 207, Mini Van 208, and Mini Van 209.

- Works on phones and desktops
- List view and calendar view, both showing which vans are reserved, the date range, and the team
- Reservations are by full day (24 hours, midnight to midnight)
- First come, first served: once a van is reserved for a date, nobody else can book it
- Coaches sign in with their W&L Microsoft 365 (Outlook) account
- No approval step and no manual calendar entry

## How it fits together

| Piece | What it does |
|---|---|
| GitHub Pages (this repo) | Hosts the page coaches open |
| W&L Microsoft sign-in | Confirms who is booking |
| SharePoint list "Van Reservations" | Stores every reservation (your history and reports) |
| Power Automate flow 1: List | Sends active reservations to the page |
| Power Automate flow 2: Book | Checks for conflicts, then saves the reservation |

## Files

| File | Purpose |
|---|---|
| `index.html` | Page layout |
| `styles.css` | Look and feel, mobile layout, dark mode, print styles |
| `app.js` | Views, booking logic, sign-in |
| `config.js` | The only file you need to edit: flow URLs, sign-in IDs, vans, teams |

Until the flow URLs are filled in, the page runs in **demo mode** with sample reservations saved only in your browser, so you can try it right away.

---

## Step 1. Put the code on GitHub

1. Create a new repository (for example `van-reservations`) and upload all four files to the top level.
2. Go to **Settings > Pages**. Under **Build and deployment**, choose **Deploy from a branch**, branch `main`, folder `/ (root)`, and save.
3. After a minute the page is live at `https://YOUR-GITHUB-NAME.github.io/van-reservations/`. Write this address down; you need it in Step 4.

Tip: create the repository under a department or shared GitHub account if possible, so it stays accessible after staffing changes.

## Step 2. Create the SharePoint list

On your Microsoft 365 group site, create a list named **Van Reservations**. Create each column with exactly these names (no spaces), because the flows refer to them by name.

| Column | Type | Notes |
|---|---|---|
| Title | (built in) | The flow fills in something like "Van 208 - Football" |
| Van | Single line of text | 97, 140, 203, 206, 207, 208, or 209 |
| StartDate | Single line of text | Stored as YYYY-MM-DD |
| EndDate | Single line of text | Stored as YYYY-MM-DD |
| Team | Single line of text | |
| DriverName | Single line of text | |
| DriverEmail | Single line of text | Signed-in W&L account |
| ReservedBy | Single line of text | Name of the person who booked |
| Destination | Single line of text | |
| Notes | Multiple lines of text | |
| Status | Choice: Active, Cancelled | Default value: Active |

Dates are stored as text on purpose. It avoids time zone shifts and makes "full day" reservations exact.

**To cancel a reservation**, open the list and change its Status to Cancelled. The van frees up on the page within a minute.

**For printable records**, use the list's **Export > Export to Excel**, or filter the list and print it.

## Step 3. Build the two Power Automate flows

Both flows start with the trigger **When an HTTP request is received**.

> Check this first: in some Microsoft 365 setups this trigger is a premium feature. Save a test flow and run it once. If you see a license or premium warning, stop and let me know, because the design would need to change to stay at no cost.

For both triggers, set **Who can trigger the flow?** to **Anyone** (see "Locking the booking flow" below for the stricter option).

On every **Response** action, add this header so the GitHub page is allowed to read the reply:

| Key | Value |
|---|---|
| Access-Control-Allow-Origin | `*` |
| Content-Type | `application/json` |

### Flow 1: Van Reservations - List

1. **Trigger:** When an HTTP request is received. Method: POST. No schema needed.
2. **SharePoint: Get items.** Site: your group site. List: Van Reservations.
   - Filter Query: `Status eq 'Active'`
   - Top Count: `5000`
3. **Data Operation: Select.** From: `value` from Get items. Map:

   | Key | Value (expression) |
   |---|---|
   | id | `item()?['ID']` |
   | van | `item()?['Van']` |
   | startDate | `item()?['StartDate']` |
   | endDate | `item()?['EndDate']` |
   | team | `item()?['Team']` |
   | driver | `item()?['DriverName']` |
   | destination | `item()?['Destination']` |

   Email addresses and notes are left out on purpose, since anyone with the page can see this list.
4. **Response.** Status code 200, the headers above, Body:
   ```
   { "reservations": @{body('Select')} }
   ```
5. Save. Copy the **HTTP URL** from the trigger.

### Flow 2: Van Reservations - Book

1. **Trigger:** When an HTTP request is received. Method: POST. Request Body JSON Schema:
   ```json
   {
     "type": "object",
     "properties": {
       "van": { "type": "string" },
       "vanName": { "type": "string" },
       "startDate": { "type": "string" },
       "endDate": { "type": "string" },
       "team": { "type": "string" },
       "driver": { "type": "string" },
       "email": { "type": "string" },
       "reservedBy": { "type": "string" },
       "destination": { "type": "string" },
       "notes": { "type": "string" }
     },
     "required": ["van", "startDate", "endDate", "team", "driver", "email"]
   }
   ```
2. **Make it first come, first served.** Open the trigger's **Settings**, turn on **Concurrency control**, and set **Degree of parallelism** to **1**. Requests are then handled one at a time, so two coaches clicking at the same moment can never both get the same van.
3. **Condition: Is the request valid?** Switch to the advanced/expression editor and use:
   ```
   @or(
     less(triggerBody()?['endDate'], triggerBody()?['startDate']),
     less(triggerBody()?['startDate'], formatDateTime(convertFromUtc(utcNow(), 'Eastern Standard Time'), 'yyyy-MM-dd')),
     not(contains(createArray('97','140','203','206','207','208','209'), triggerBody()?['van'])),
     not(endsWith(toLower(triggerBody()?['email']), '@wlu.edu'))
   )
   ```
   - **If yes:** Response, status 400, body `{ "message": "The reservation details were not accepted. Please check the dates and try again." }`, then **Terminate** (Succeeded).
4. **SharePoint: Get items.** List: Van Reservations. Filter Query:
   ```
   Van eq '@{triggerBody()?['van']}' and Status eq 'Active'
   ```
5. **Data Operation: Filter array.** From: `value` from step 4. Advanced mode:
   ```
   @and(
     lessOrEquals(item()?['StartDate'], triggerBody()?['endDate']),
     greaterOrEquals(item()?['EndDate'], triggerBody()?['startDate'])
   )
   ```
6. **Condition:** `length(body('Filter_array'))` is greater than `0`.
   - **If yes (conflict):** Response, status 409, body `{ "message": "That van was just reserved for some of those dates. Please choose another van or different dates." }`
   - **If no:**
     1. **SharePoint: Create item** in Van Reservations:
        - Title: `Van @{triggerBody()?['van']} - @{triggerBody()?['team']}`
        - Van, StartDate, EndDate, Team: from the trigger
        - DriverName: `driver`, DriverEmail: `email`, ReservedBy: `reservedBy`
        - Destination, Notes: from the trigger
        - Status: Active
     2. Optional: **Send an email from a shared mailbox (V2)** from athletics@wlu.edu to `email`, confirming the van and dates.
     3. **Response**, status 200, body:
        ```
        { "id": "@{body('Create_item')?['ID']}", "message": "Reserved" }
        ```
7. Save. Copy the **HTTP URL** from the trigger.

Make the athletics@wlu.edu group or a second staff member a **co-owner** of both flows and the list, so they keep running after staffing changes.

## Step 4. Set up W&L sign-in

Sign-in needs a free "app registration" in Microsoft Entra ID (W&L's Microsoft 365 directory). Many schools let staff create these; some require IT. If you can't create one, forward the box below to IT.

1. Go to **entra.microsoft.com > Applications > App registrations > New registration**.
2. Name: `Athletics Van Reservations`.
3. Supported account types: **Accounts in this organizational directory only**.
4. Redirect URI: platform **Single-page application (SPA)**, address = your GitHub Pages address from Step 1, including the trailing slash.
5. Register. On the Overview page, copy the **Application (client) ID** and **Directory (tenant) ID**.
6. API permissions should list Microsoft Graph **User.Read** (added by default). Nothing else is needed.

> **Request for IT (copy and send):** Please create a single-page application registration in Entra ID named "Athletics Van Reservations," single tenant, with SPA redirect URI `https://YOUR-GITHUB-NAME.github.io/van-reservations/` and delegated Microsoft Graph User.Read. It is used only to sign W&L staff into the Athletics van reservation page. Please send me the Application (client) ID and Directory (tenant) ID.

## Step 5. Fill in config.js

```js
listUrl: "PASTE FLOW 1 HTTP URL",
bookUrl: "PASTE FLOW 2 HTTP URL",
auth: {
  clientId: "PASTE APPLICATION (CLIENT) ID",
  tenantId: "PASTE DIRECTORY (TENANT) ID",
  redirectUri: "",
  flowScope: ""
},
```

Commit the change. The demo banner disappears and the page goes live.

## Step 6. Test before sharing

1. Open the page on your phone and on a computer. Sign in with your W&L account.
2. Reserve a van for tomorrow. Confirm it appears in the SharePoint list, in List view, and in Calendar view with the team and date range.
3. Try to reserve the same van for an overlapping date. It should not be offered.
4. Change the test reservation's Status to Cancelled in SharePoint and refresh the page. The van should show as available again.

---

## Locking the booking flow (optional, stronger security)

With "Anyone" selected, the flow URLs are visible in the GitHub code, so someone who found them could send a fake booking. Sign-in on the page plus the @wlu.edu check in the flow keeps this low risk for an internal tool. For full protection:

1. In Flow 2's trigger, change **Who can trigger the flow?** to **Any user in my tenant**.
2. In the app registration, add the delegated API permission for the Power Automate service (listed as "Microsoft Flow Service" or "Power Automate," permission **User**). This may need admin consent from IT.
3. In `config.js`, set `flowScope` to `https://service.flow.microsoft.com//.default`. Confirm the exact value with IT, since Microsoft has adjusted this over time.

The page then sends the signed-in coach's token with every booking, and the flow rejects anything without one.

## Troubleshooting

- **"Could not load reservations"**: open the browser's developer console. A CORS error means a Response action is missing the `Access-Control-Allow-Origin` header. A 401 or 403 means the trigger is not set to "Anyone."
- **Sign-in loops back to the page with no account**: the redirect URI in Entra must match the GitHub Pages address exactly, including `https` and the trailing slash.
- **A van never shows as booked**: check that the Van column holds just the number (for example `208`, not `Mini Van 208`).
- **Adding or retiring a van**: edit the `vans` list in `config.js` and the `createArray(...)` list in Flow 2, step 3.
