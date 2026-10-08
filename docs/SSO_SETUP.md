# Turn on automatic sign-in (Microsoft 365 single sign-on)

With single sign-on on, the add-in reads each person's **name and work email** from the Microsoft account they're already signed in to Office with (the same account they use to log in to Windows at MHS). Nobody types a name or email.

- **Until this is set up:** the add-in works as before and asks once for name and email.
- **Still chosen by each person:** role and assigned clinics. Microsoft 365 doesn't know them, and both are optional.
- **What identity doesn't do:** it doesn't decide who can see data. The workbook's SharePoint permissions still control that.

## 1. Register the app in Microsoft Entra
You need app-registration rights, or ask IT. This takes about 10 minutes.

1. Go to **entra.microsoft.com → Applications → App registrations → New registration**.
   - Name: `MHS Billing Tickets`
   - Supported account types: **Accounts in this organizational directory only**
   - Redirect URI: leave blank
   - Click **Register**, then copy the **Application (client) ID**.
2. Under **Expose an API**:
   1. Set the **Application ID URI** to `api://brendannyates.github.io/<client-id>`. The host must match where the add-in is hosted; if the repo moves to the MHS org, use that host.
   2. Click **Add a scope**:
      - Scope name: `access_as_user`
      - Who can consent: **Admins and users**
      - Fill in the display names and descriptions, e.g. "Read your profile for MHS Billing Tickets".
   3. Click **Add a client application** twice, once for each Office client ID below. Tick the `access_as_user` scope each time.
      - `ea5a67f6-b6f3-4338-b240-c655ddc3cc8e`: Microsoft Office, all clients
      - `d3590ed6-52b3-4102-aeff-aad2292ab01c`: Office desktop
3. Under **API permissions**, keep the default **Microsoft Graph → User.Read**, then click **Grant admin consent for Mindful Health Solutions**. This step needs an admin, and it means nobody sees a consent prompt.
4. Under **Token configuration**, nothing is required. The add-in uses the standard `name` and `preferred_username` claims.

## 2. Put the ID in the add-in
Send the client ID to Claude, or run this in the repo:
```bash
node scripts/set-sso.mjs <client-id>
npm run build && npm test
git commit -am "Enable SSO" && git push
```

## 3. Roll it out
- This changes the manifest, so **everyone re-uploads `manifest.xml`**, or IT updates it once under **M365 admin center → Integrated apps**.
- On the next open, people's name and email fill in automatically. Settings → *Your profile* shows "from your Microsoft sign-in".

## If it doesn't work
| Symptom | Fix |
|---|---|
| Still asks for name and email | The manifest without the ID is still cached. Re-upload it, or on desktop clear `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\` and reopen. |
| A consent prompt appears | Admin consent (step 1.3) wasn't granted. |
| Sign-in error 13xxx in some clients | Check that the Application ID URI host matches the manifest's `<SourceLocation>` host exactly, and that both Office client IDs are pre-authorized. |
