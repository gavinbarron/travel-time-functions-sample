const https = require('https');
const Promise = require('es6-promise').Promise;
const DocumentClient = require('documentdb').DocumentClient;
const cfg = require('../common/config')
const UserRepository = require('../common/models/userRepository');

module.exports = function (context, myTimer) {
    var timeStamp = new Date().toISOString();

    if (myTimer.isPastDue) {
        context.log.warn('JavaScript is running late!');
    }
    context.log('JavaScript timer trigger function ran!', timeStamp);

    // Init DocumentDb client and repo
    var documentDbOptions = {
        host: process.env.docDbHost,
        masterKey: process.env.docDbKey,
        database: process.env.docDbName || 'botdocdb',
        collection: process.env.docDbCollection ||'botdata'
    };

    var docDbClient = new DocumentClient(documentDbOptions.host, {masterKey: documentDbOptions.masterKey});
    var userRepo = new UserRepository(docDbClient, documentDbOptions.database, documentDbOptions.collection);

    function refreshAccess(user) {
        return new Promise( (resolve, reject) => {
            const postBody = `client_id:${cfg.AZUREAD_APP_ID}`+
            `scope:${cfg.SCOPES.join(' ')}`+
            `refresh_token:${user.data.refresh_token}`+
            `redirect_uri:https://${cfg.BOT_HOSTNAME}/botauth/aadv2/callback`+
            `grant_type:refresh_token`+
            `client_secret:${cfg.AZUREAD_APP_PASSWORD}`;

            const postOpts = {
                host: 'login.microsoftonline.com',
                port: '443',
                path: '/common/oauth2/v2.0/token',
                method: 'POST',
                headers: {
                    'Content-Type': 'application/x-www-form-urlencoded',
                    'Content-Length': postBody.length
                }
            };
            let req = https.request(postOpts, function (res) {
                let responseData = ''

                res.on('data', chunk => responseData += chunk);
                res.on('end', () => {
                    if (res.statusCode !== 200) reject(JSON.parse(responseData));
                    else {
                        const refreshData = JSON.parse(responseData);

                        // Update the user object and save it back to DocDb
                        user.data.access_token = refreshData.access_token;
                        user.data.refresh_token = refreshData.refresh_token;
                        userRepo.updateItem(user, (err) => {
                            if (err) {
                                reject(err);
                            }
                            resolve(`User with subscriptionId ${user.data.subscriptionId} refreshed`);
                        });
                    }
                });
            });
            req.write(postBody);
            req.end();

            req.on('error', error => reject(error));
        });
    }

    function reflect(promise) {
        return promise.then(function (resolved) { return {result: resolved, status: 'resolved' }},
                            function (rejected) { return {result: rejected, status: 'rejected' }});
    }

    userRepo.getUsersWithRefreshTokens((err, results) => {
        const promises = new Array();
        for (const user in results) {
            promises.push(refreshAccess(user))
        }
        Promise.all(promises.map(reflect)).then((results) => {
            for (const result in results) {
                if (result.status === 'resolved') {
                    context.log.error(`Error: ${result.result}`);
                } else {
                    context.log.info(`Success: ${result.result}`);
                }
            }
            context.done();
        });
    });
}