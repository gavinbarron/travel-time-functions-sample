const { DocumentClient, DocumentBase } = require('documentdb');
const UserRepository = require('./userRepository');
const https = require('https');

module.exports = function (context, notification) {
    context.log('Processsing calendar event notification');

    /**
   * Generates a GET request (of Content-type ```application/json```)
   * @param {string} path the path, relative to the host, to which this request will be sent
   * @param {string} token the acess token with which the request should be authenticated
   * @param {callback} callback
   */
    function getData (path, token, callback) {
        var self = this;
        const options = {
        host: 'graph.microsoft.com',
        path: path,
        method: 'GET',
        headers: {
            'Content-Type': 'application/json',
            Accept: 'application/json;odata.metadata=minimal;' +
                    'odata.streaming=true;IEEE754Compatible=false',
            Authorization: 'Bearer ' + token
        }
        };

        const req = https.request(options, res => {
            let endpointData = '';
    
            res.on('data', chunk => (endpointData += chunk));
            res.on('end', () => {
                if (res.statusCode === 200) callback(null, JSON.parse(endpointData));
                else callback(JSON.parse(endpointData), null);
            });
        });

        req.write('');
        req.end();

        req.on('error', error => callback(error, null));
    }

    // Init DocumentDb client and repo
    var documentDbOptions = {
        host: process.env.docDbHost || 'https://calendarbot.documents.azure.com:443', // 'https://localhost:8081',  Host for local DocDb emulator
        masterKey: process.env.docDbKey || 'YRWO4JRhpXT3P2n4ARUWYnG4U5HbY9zBgr24zfONkSRrxDoHM0L0FAea4ANdMPOqCoBjopcAASuAFzMzYj5xPQ==', //  'C2y6yDjf5/R+ob0N8A7Cgv30VRDJIWEHLM+4QDU5DE2nQ9nDuVTqobD4b8mGGyPMbIZnqyMsEcaGQy67XIw/Jw==', Fixed key for local DocDb emulator    
        database: process.env.docDbName || 'botdocdb',
        collection: process.env.docDbCollection ||'botdata'
    };

    var docDbClient = new DocumentClient(documentDbOptions.host, {masterKey: documentDbOptions.masterKey});
    var userRepo = new UserRepository(docDbClient, documentDbOptions.database, documentDbOptions.collection);
    userRepo.init((err, coll) => {
        if (err) {
            context.done(err);
            return;
        }
    });

    setTimeout(()=> {
        context.log('User Repository data connection initialized')
        context.log(`subscriptionId ${notification.subscriptionId}`)
        // Get the User from DocDB based on the subscriptionId in the message
        userRepo.getItemBySubscriptionId(notification.subscriptionId, (err, user) => {
            if (err) {
                context.done(err);
                return;
            } 
            if (!user) {
                context.done(new Error('unable to load user'))
                return;
            } 
            // Get the access token (should use the refresh token to renew this....)
            // Get the Calendar event from Microsoft Graph
            getData(`/v1.0/${notification.resource}`, user.data.accessToken, (error, data) => {
                if (error) {
                    context.log(error);
                    context.done(error);
                    return;
                }
                // Parse event for details
                const eventSummary = {
                    'action': 'NOTIFICATION_RECIEVED',
                    'address': user.data.address,
                    'resource': notification.resource,
                    'organizer': data.organizer.emailAddress.name || data.organizer.emailAddress.address,
                    'location':  data.location.displayName,
                    'subject': data.subject,
                    'start': `${data.start.dateTime}Z`,
                    'end': `${data.end.dateTime}Z`
                };
                //TODO: find a better way of getting the default mail address for the user.
                if(data.organizer.emailAddress.address.toLowerCase() === 'my.email@domain.com') {
                    //don't react to meetings I've organized.
                    context.done(new Error('Meeting from self'));
                    return;
                }
                context.log(eventSummary);
                context.done(null, eventSummary);
            });
        });
    }, 2000)
    
}