const GraphHelper = require('../common/graphHelper');

module.exports = function (context, myQueueItem) {
    context.log('Starting accept call');

    const acceptData = {
        comment: 'Accepted via the TravelTime Bot',
        sendResponse: true
    };
    const data = JSON.stringify(acceptData);
    context.log(`Accept call ${data}`);

    GraphHelper.postData(`/v1.0/${myQueueItem.meeting}/accept`, myQueueItem.accessToken, data, (error, data) => {
        if (error) {
            context.done(error);
            return;
        }
        context.done(null, data);
    });
};

