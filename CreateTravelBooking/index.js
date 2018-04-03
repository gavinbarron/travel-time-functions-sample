const GraphHelper = require('../common/graphHelper');

module.exports = function (context, myQueueItem) {
    context.log('Starting travel meeting booking');

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
    var graphHelper = new GraphHelper();
    GraphHelper.postData(`/v1.0/me/events`, myQueueItem.accessToken, JSON.stringify(meeting), (error, data) => {
        if (error) {
            context.done(error);
            return;
        }
        context.done(null, data);
    });
};