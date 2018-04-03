const https = require('https');
const cfg = require('../common/config')

module.exports = function (context, myQueueItem) {
    context.log('Registering web hook');

    // set expiry in two days from registration
    let expires = new Date();
    expires.setDate(expires.getDate() + 2);
    const postData = {
        "changeType": "created",
        "notificationUrl": "https://graph-functions-node.azurewebsites.net/api/CalendarWebHookHandler?code=9EJaI7/UQ/W1Yvu3Lw48hFM0ZKippfJejmPVSdbVxReAO8cYb5iaYg==",
        "resource": "me/events",
        "expirationDateTime": expires.toISOString(),
        "clientState": cfg.CLIENT_STATE
    };
    const postBody = JSON.stringify(postData);

    const postOpts = {
        host: 'graph.microsoft.com',
        port: '443',
        path: '/v1.0/subscriptions',
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
            'Content-Length': postBody.length,
            'Authorization': 'Bearer ' + myQueueItem.accessToken
        }
    };
    try {
        let req = https.request(postOpts, function (res) {
            let subscriptionData = '';

            res.on('data', chunk => subscriptionData += chunk);
            res.on('end', () => {
                context.log(`http status: ${res.statusCode}`);
                if (res.statusCode === 201) context.done(JSON.parse(subscriptionData));
                else {
                    const subscription = JSON.parse(subscriptionData);
                    context.log(subscription);
                    if (!subscription.id) {
                        context.done('subscription id not present');
                        return;
                    }
                    var message = {
                        'action': 'WEBHOOK_REGISTERED',
                        'status': 'complete',
                        'subscriptionId': subscription.id,
                        'address': myQueueItem.address
                    };
                    context.log(message);
                    context.done(null, message);
                }
            });
        });
        context.log(postBody);
        req.write(postBody);
        req.end();
    } catch (e) {
        context.log(e);
        context.done(e);
    }
};