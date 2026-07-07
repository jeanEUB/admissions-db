# admissions-db

## Microsoft 365 sign-in

This app now includes a Microsoft 365 login screen backed by Microsoft Entra ID and MSAL.

Page structure:

1. `index.html` is the dedicated Microsoft sign-in page.
2. `admissions_db.html` hosts the full admissions application after authentication.

To enable it:

1. Register a single-page application in Microsoft Entra ID.
2. Add your login page URL as a redirect URI, for example `http://localhost:5500/index.html`.
3. Copy the application client ID into `auth-config.js`.
4. If you use a tenant-specific app, replace `tenantId: 'common'` with your tenant ID.
5. Set `loginPageUrl` and `appPageUrl` in `auth-config.js` so MSAL navigation uses the right pages.
6. Serve the app over HTTP or HTTPS. The Microsoft login flow will not work correctly from `file://` URLs.

The app stays on the login page until Microsoft sign-in succeeds, then it reveals the admissions dashboard.