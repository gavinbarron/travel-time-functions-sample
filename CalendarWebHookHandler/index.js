
const azureStorage = require('azure-storage');
const cfg = require('../common/config')

module.exports = function (context, data) {
    context.log('Calendar Webhook Handler was triggered!');

    // Handle subscription request
    if (data.query.validationToken) {
        context.log('registering new hook');
        context.res = {
            status: 200,
            body: data.query.validationToken
        };
        context.done();
        return;
    }

    const storageConnectionString = process.env.storageConnectionString || 'DefaultEndpointsProtocol=https;AccountName=travelerbot;AccountKey=b2cmbtSrWvJ2ebxO0PD1CQ+svlaxjUY3xtC8DdzJyeFXRGVgnsGhZkqU82rvgMiXE0ybEaWlTZpEsf/65drmNA==;EndpointSuffix=core.windows.net';
    const queueSvc = azureStorage.createQueueService(storageConnectionString);
    // Write incoming notifications onto a queue.
    queueSvc.createQueueIfNotExists('hook-recieved', (err, result, response) => {
        if (err){
            // this should be a log for the dev, not a message to the user
            session.send('There was an error creating the hook-recieved queue: '+ err);
            context.done(err);
            return;
        }
        // enqueue a message to process the webhook request.
        const request = data.body;
        context.log(request);
        const notificationCount = request.value.length;
        let processed = 0; 
        // web hook requests can contain multiple notifications
        for (let hook of request.value) {
            let msg = JSON.stringify(hook);
            context.log(msg);
            var queueMessageBuffer = new Buffer(msg).toString('base64');
            // only take action when the clientState is vaild.
            // return the 202 and do notthing if the clientState is vaild.
            if (hook.clientState !== cfg.CLIENT_STATE) {
                // add a log entry so that this can be investigated
                context.log('Unknown sender! Investigate the source of this message')
                context.log(data);
                context.res = { status: 202, body: 'All notifications processed' }; 
                context.done();
                return;
                break;
            } 
            queueSvc.createMessage('hook-recieved', queueMessageBuffer, (err, result, response) => {
                processed++;
                // fail on any message not getting queued properly
                if (err) {
                    context.log('error sending webhook messge on hook-recieved queue');
                    context.done(err);
                    return;
                }
                context.log('webhook messge on put onto hook-recieved queue');
                if (processed == notificationCount) {
                    context.res = { status: 202, body: 'All notifications processed' }; 
                    context.done();
                }
            });
        }
    });
}
