const https = require('https');

module.exports = function (context, myQueueItem) {
    context.log('Starting accept call');

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
            let responseData = '';
            context.log('In response');
            
            res.on('data', chunk => {
                // context.log('got data');
                (responseData += chunk);
            });
            res.on('end', () => {
                // context.log('data end');
                // context.log(res.statusCode);
                // context.log(responseData);
                // context.log(JSON.parse(responseData));
                if (res.statusCode === 201) callback(null, JSON.parse(responseData));
                else callback(JSON.parse(responseData), null);
            });
        });

        req.write(data);
        req.end();
        context.log('Made request');
        
        req.on('error', error => {
        context.log('Bad error');
            callback(error, null)
        });
    }

    const acceptData = {
        comment: '',
        sendResponse: true
    };
    const data = JSON.stringify(acceptData);
    context.log(`Accept call ${data}`);

    postData(`/v1.0/${myQueueItem.meeting}/accept`, myQueueItem.accessToken, data, (error, data) => {
        if(error){
            context.done(error);
            return;
        }
        context.done();
    });
};

