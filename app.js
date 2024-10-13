require('dotenv').config();
const express = require('express');
const path = require('path');
const app = express();
const msal = require('@azure/msal-node');
const session = require('express-session');

app.use(express.static(path.join(__dirname, 'public')));

// Set EJS as the templating engine
app.set('view engine', 'ejs');

// Middleware for parsing JSON bodies
app.use(express.json());

// Middleware for session management
app.use(
    session({
        secret: 's3cret',
        resave: false,
        saveUninitialized: false,
    })
);

// Middleware for logging requests (example of custom middleware)
app.use((req, res, next) => {
    console.log(`${req.method} request for '${req.url}'`);
    next();
});

// MSAL configuration
const msalConfig = {
    auth: {
        clientId: process.env.CLIENT_ID, // Your Azure AD Application client ID
        authority: process.env.AUTHORITY, // Your Azure AD tenant ID
        clientSecret: process.env.CLIENT_SECRET, // Your Azure AD Application client
        redirectUri: "http://localhost:3456/redirect", // Redirect URI (must match Azure AD configuration)
    },
};

const scopes = [
    'https://analysis.windows.net/powerbi/api/Report.Read.All',
    'https://analysis.windows.net/powerbi/api/Dataset.Read.All',
    'https://analysis.windows.net/powerbi/api/App.Read.All',
    'https://analysis.windows.net/powerbi/api/Item.Execute.All',
]

const graphScopes = ["https://graph.microsoft.com/.default"];

async function acquireGraphToken() {
    const accounts = await cca.getTokenCache().getAllAccounts();
    console.log(accounts);
    if (accounts.length > 0) {
        const silentRequest = {
            account: accounts[0],
            scopes: graphScopes,
        };

        return cca.acquireTokenSilent(silentRequest);
    }
}

// Create MSAL application object
const cca = new msal.ConfidentialClientApplication(msalConfig);

// Base route
app.get('/', (req, res) => {
    if (req.session.user) {
        res.redirect('/dashboard');
    } else {
        res.render('index', {});
    }
});

// Login route (initiates sign-in)
app.get('/login', (req, res) => {
    const authUrl = cca.getAuthCodeUrl({
        scopes,
        redirectUri: msalConfig.auth.redirectUri,
    });

    // Redirect to Azure AD login page
    authUrl.then((url) => {
        res.redirect(url);
    }).catch((error) => {
        console.log(JSON.stringify(error));
        res.status(500).send('Error during authentication');
    });
});

// Redirect route (handles the response from Azure AD)
app.get('/redirect', (req, res) => {
    const tokenRequest = {
        code: req.query.code,
        scopes,
        redirectUri: msalConfig.auth.redirectUri,
    };

    cca.acquireTokenByCode(tokenRequest)
        .then((response) => {
            // console.log('\nResponse: \n:', response);
            req.session.user = response;
            console.log('token:', response.accessToken);
            res.redirect('/dashboard');
        })
        .catch((error) => {
            console.log(JSON.stringify(error));
            res.status(500).send('Error acquiring token');
        });
});

// Logout route
app.get('/logout', (req, res) => {
    const logoutUri = `https://login.microsoftonline.com/b9d2e4d7-3331-44aa-b735-189229b4c840/oauth2/v2.0/logout?post_logout_redirect_uri=http://localhost:${PORT}/`;
    req.session.destroy(() => {
        // res.redirect(logoutUri);
        res.redirect('/');
    });
});

//Power BI 
const WORKSPACE_ID = '6be6316e-d8e9-4751-84c3-34c6eebab80f'; // portal
const REPORT_ID = '1c0eaa4e-059d-434c-a1de-c1a32bbd66a0'; // AdventureWorks Sales

async function getEmbedToken(token) {
    const url = `https://api.powerbi.com/v1.0/myorg/groups/${WORKSPACE_ID}/reports/${REPORT_ID}/GenerateToken`;
    const embedTokenURL = `https://app.powerbi.com/reportEmbed?reportId=${REPORT_ID}&groupId=${WORKSPACE_ID}`;
    const headers = {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${token}`
    };
    const body = {
        "accessLevel": "View",
        "allowSaveAs": "false"
    };

    try {
        const response = await fetch(url, {
            method: 'POST',
            headers: headers,
            body: JSON.stringify(body)
        });

        const data = await response.json();
        return {
            accessToken: token,
            embedToken: data,
            embedTokenURL: embedTokenURL + `&embedToken=${data.token}`,
            workspaceId: WORKSPACE_ID,
            reportId: REPORT_ID
        }
    }
    catch (error) {
        console.error('Error acquiring embed token:', error);
    }

}
app.get('/embed-token', async (req, res) => {

    // console.log('user_token', req.session.user);

    const data = await getEmbedToken(req.session.user.accessToken);

    res.json({ data });

    // res.sendStatus(200);


});

app.get('/trigger-pipeline', async (req, res) => {

    const URL = "https://api.fabric.microsoft.com/v1/workspaces/6be6316e-d8e9-4751-84c3-34c6eebab80f/items/b8d9d0a3-ee71-4f86-96dd-2579f3ee4b16/jobs/instances?jobType=Pipeline";

    fetch(URL, {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
            'Authorization': `Bearer ${req.session.user.accessToken}`
        }
    })
        .then(response => console.log(response))
        .catch(error => console.error('Error triggering pipeline:', error))
        .finally(() => res.sendStatus(200));

    //itemid b8d9d0a3-ee71-4f86-96dd-2579f3ee4b16
    //workspaceid 6be6316e-d8e9-4751-84c3-34c6eebab80f
    //type Pipeline 
});

app.get('/send-email', async (req, res) => {

    try {

        const URL = "https://api.powerbi.com/v1.0/myorg/reports/1c0eaa4e-059d-434c-a1de-c1a32bbd66a0/ExportTo";

        const resp = await fetch(URL, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'Authorization': `Bearer ${req.session.user.accessToken}`
            },
            body: JSON.stringify({
                "format": "PNG",
                "powerBIReportConfiguration": {
                    "pages": [
                        {
                            "pageName": "ReportSection",  // Replace with your specific page name
                            "visualName": "3a1aeaede6fc79fe5066"  // Replace with your specific visual name
                        }
                    ]
                }
            })
        })

        const { id } = await resp.json();

        const pollInterval = setInterval(async () => {

            const exportStatus = await pollExportStatus(id, req.session.user.accessToken);

            if (exportStatus.status === 'Succeeded') {
                clearInterval(pollInterval);
                console.log('Export status:', exportStatus);
                console.log('Export succeeded');
                const b64 = await downloadAndConvertToBase64(exportStatus, req.session.user.accessToken);
                if (!b64) res.sendStatus(500);
                const emailResponse = await sendEmail(req.session.user.accessToken, b64);
                console.log('Email response:', emailResponse);
                res.send('<img src="data:image/png;base64,' + b64 + '" />');
                // res.sendStatus(200);

            }

        }, 5000);


    } catch (error) {
        console.error('Error sending email:', error);
        res.sendStatus(500);
    }

});

async function sendEmail(token, b64) {

    try {

        const { accessToken } = await acquireGraphToken();

        const URL = "https://graph.microsoft.com/v1.0/me/sendMail";

        const response = await fetch(URL, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'Authorization': `Bearer ${accessToken}`
            },
            body: JSON.stringify({
                "message": {
                    "subject": "Power BI Report",
                    "body": {
                        "contentType": "HTML",
                        "content": "<p>Power BI Report</p>"
                    },
                    "toRecipients": [
                        {
                            "emailAddress": {
                                "address": "ecpar@trends.com.ph",
                            }
                        }
                    ],
                    "attachments": [
                        {
                            "@odata.type": "#microsoft.graph.fileAttachment",
                            "name": "report.png",
                            "contentBytes": b64
                        }
                    ]
                }
            })
        });

        return await response.json();

    } catch (error) {
        console.error('Error sending email:', error);
        return null;

    }
}

async function downloadAndConvertToBase64(status, token) {

    // https://api.powerbi.com/v1.0/myorg/reports/1c0eaa4e-059d-434c-a1de-c1a32bbd66a0/exports/MS9...ZjPS4=/file
    const URL = status.resourceLocation;

    const response = await fetch(URL, {
        method: 'GET',
        headers: {
            'Content-Type': 'application/json',
            'Authorization': `Bearer ${token}`
        }
    });

    if (!response.ok) {
        return null;
    }

    const arrayBuffer = await response.arrayBuffer();

    const base64Data = Buffer.from(arrayBuffer).toString('base64');

    return base64Data;

}


async function pollExportStatus(id, token) {

    console.log('Exporting image...');

    const pollUrl = `https://api.powerbi.com/v1.0/myorg/reports/1c0eaa4e-059d-434c-a1de-c1a32bbd66a0/exports/${id}`

    const pollRes = await fetch(pollUrl, {
        method: 'GET',
        headers: {
            'Content-Type': 'application/json',
            'Authorization': `Bearer ${token}`
        }
    });

    return await pollRes.json();

}

// Dashboard route
app.get('/dashboard', (req, res) => {
    // console.log(req.session);
    if (req.session.user) {
        res.render('dashboard', {
            username: req.session.user.account.username,
            name: req.session.user.name,
        });
        // res.redirect('https://app.powerbi.com/reportEmbed?reportId=1c0eaa4e-059d-434c-a1de-c1a32bbd66a0&autoAuth=true&ctid=b9d2e4d7-3331-44aa-b735-189229b4c840');
    } else {
        res.redirect('/');
    }
});

// Start the server
const PORT = process.env.PORT || 3456;
app.listen(PORT, () => {
    console.log(`Server is running on http://localhost:${PORT}`);
});
