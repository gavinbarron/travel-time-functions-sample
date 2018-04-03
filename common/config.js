const cfg = {
    CLIENT_STATE: 'c1i3nt5t4t3',
    BOT_HOSTNAME: process.env.bothost || '********.ngrok.io',
    AZUREAD_APP_ID: 'fda1b8f2-d40b-4aa2-ab40-ca2ac951af63' || process.env.azAppId ||'guid-from-apps-dev-microsoft-com',
    AZUREAD_APP_PASSWORD: 'vnwiXILAO8kalED0105-^%|' || process.env.azAppPassword || 'password-from-apps-dev-microsoft-com',
    AZUREAD_APP_REALM: 'common',
    SCOPES: ['Calendars.ReadWrite', 'User.Read', 'offline_access', 'https://graph.microsoft.com/MailboxSettings.Read']
}
module.exports = cfg;