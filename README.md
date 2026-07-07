# admissions-db

## Microsoft 365 sign-in

This app now includes a Microsoft 365 login screen backed by Microsoft Entra ID and MSAL.

To enable it:

1. Register a single-page application in Microsoft Entra ID.
2. Add your local app URL as a redirect URI, for example `http://localhost:5500/index.html`.
3. Copy the application client ID into `auth-config.js`.
4. If you use a tenant-specific app, replace `tenantId: 'common'` with your tenant ID.
5. Serve the app over HTTP or HTTPS. The Microsoft login flow will not work correctly from `file://` URLs.

The app stays on the login page until Microsoft sign-in succeeds, then it reveals the admissions dashboard.