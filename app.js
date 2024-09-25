const express = require('express');
const path = require('path');
const app = express();
const msal = require('@azure/msal-node');
const session = require('express-session');
const helmet = require("helmet");

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
    console.log('req headers', req.headers);
    console.log(`${req.method} request for '${req.url}'`);
    next();
});

// MSAL configuration
const msalConfig = {
    auth: {
        clientId: '',              // Your Azure AD Application (client) ID
        authority: 'https://login.microsoftonline.com/TENANT_ID', // Your Azure AD tenant ID
        clientSecret: '',      // Your Azure AD Application client secret
        redirectUri: 'http://localhost:3456/redirect',  // Redirect URI (must match Azure AD configuration)
    },
};

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
        scopes: ['https://analysis.windows.net/powerbi/api/.default'],
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
        scopes: ['https://analysis.windows.net/powerbi/api/.default'],
        redirectUri: msalConfig.auth.redirectUri,
    };

    cca.acquireTokenByCode(tokenRequest)
        .then((response) => {
            // console.log('\nResponse: \n:', response);
            req.session.user = response;
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

// Dashboard route
app.get('/dashboard', (req, res) => {
    console.log(req.session);
    if (req.session.user) {
        res.render('dashboard', {
            username: req.session.user.account.username,
            name: req.session.user.name,
        });
    } else {
        res.redirect('/');
    }
});

// Start the server
const PORT = process.env.PORT || 3456;
app.listen(PORT, () => {
    console.log(`Server is running on http://localhost:${PORT}`);
});
