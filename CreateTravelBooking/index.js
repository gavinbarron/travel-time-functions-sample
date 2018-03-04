const https = require('https');

module.exports = function (context, myQueueItem) {
    context.log('Starting travel meeting booking');
    
    /**
     * Generates a POST request (of Content-type ```application/json```)
     * @param {string} path the path, relative to the host, to which this request will be sent
     * @param {string} token the access token with which the request should be authenticated
     * @param {string} data the data which will be 'POST'ed
     * @param {callback} callback
     */
    function postData (path, token, data, callback) {
        var self = this;
        const options = {
            host: 'graph.microsoft.com',
            path: path,
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                Authorization: 'Bearer ' + token,
                'Content-Length': data.length
            }
        };

        const req = https.request(options, res => {
            let subscriptionData = '';

            res.on('data', chunk => (subscriptionData += chunk));
            res.on('end', () => {
                if (res.statusCode === 201) callback(null, JSON.parse(subscriptionData));
                else callback(JSON.parse(subscriptionData), null);
            });
        });

        req.write(data);
        req.end();

        req.on('error', error => callback(error, null));
    }
    context.log(myQueueItem.start);
    context.log(myQueueItem.durationInMins)
    const start = new Date(myQueueItem.start);
    const MS_PER_MINUTE = 60000;
    const end = new Date(start.getTime() + myQueueItem.durationInMins * MS_PER_MINUTE).toISOString();

    context.log(myQueueItem.start);
    context.log(end);

    const meeting = {
        subject: "Travel time",
        start: {
            dateTime: myQueueItem.start,
            timeZone: "UTC"
        },
        end: {
            dateTime: end,
            timeZone: "UTC"
        }
    }

    postData(`/v1.0/me/events`, myQueueItem.accessToken, JSON.stringify(meeting), (error, data) => {
        if(error){
            context.done(error);
            return;
        }
        context.done();
    });
};