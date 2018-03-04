const https = require('https');
const cfg = require('../common/config')

module.exports = function (context, myQueueItem) {
    context.log('Registering web hook');

    // set expiry in two days from registration
    let expires = new Date();
    expires.setDate(expires.getDate() + 2);
    const postData = {
        "changeType": "created",
        "notificationUrl": "https://travlerbot-function.azurewebsites.net/api/CalendarWebHookHandler?code=E9fm1Yqe8hE0wR1bfSbJ2T/6Bl5fnJR7JY2Q2E8v6l1A4TIIJZv2Nw==",
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
        let req = https.request(postOpts, function(res){
            res.setEncoding('utf8');
            res.on('data', function (data) {
                const subscription = JSON.parse(data);
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
            });
        });
        req.write(postBody);
        req.end();
    } catch (e) {
        context.log(e);
        context.done(e);
    }
};