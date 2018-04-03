const https = require('https');

function GraphHelper() {
  this.host = 'graph.microsoft.com';
}

GraphHelper.prototype = {

  /**
   * Generates a POST request (of Content-type ```application/json```)
   * @param {string} path the path, relative to the host, to which this request will be sent
   * @param {string} token the access token with which the request should be authenticated
   * @param {string} data the data which will be 'POST'ed
   * @param {callback} callback
   */
  postData: function (path, token, data, callback) {
    var self = this;
    const options = {
      host: self.host,
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

      res.on('data', chunk => subscriptionData += chunk);
      res.on('end', () => {
        if (res.statusCode === 201) callback(null, JSON.parse(subscriptionData));
        else callback(JSON.parse(subscriptionData), null);
      });
    });

    req.write(data);
    req.end();

    req.on('error', error => callback(error, null));
  },

  /**
   * Generates a GET request (of Content-type ```application/json```)
   * @param {string} path the path, relative to the host, to which this request will be sent
   * @param {string} token the acess token with which the request should be authenticated
   * @param {callback} callback
   */
  getData: function (path, token, callback) {
    var self = this;
    const options = {
      host: self.host,
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

      res.on('data', chunk => endpointData += chunk);
      res.on('end', () => {
        if (res.statusCode === 200) callback(null, JSON.parse(endpointData));
        else callback(JSON.parse(endpointData), null);
      });
    });

    req.write('');
    req.end();

    req.on('error', error => callback(error, null));
  },

  /**
   * Generates a DELETE request
   * @param {string} path the path, relative to the host, to which this request will be sent
   * @param {string} token the acess token with which the request should be authenticated
   * @param {callback} callback
   */
  deleteData: function (path, token, callback) {
    var self = this;
    const options = {
      host: self.host,
      path: path,
      method: 'POST',
      headers: {
        'Content-Type': 'text/plain',
        'X-HTTP-Method': 'DELETE',
        Authorization: 'Bearer ' + token
      }
    };

    const req = https.request(options, res => {
      let endpointData = '';
      res.on('data', chunk => endpointData += chunk);
      res.on('end', () => callback(null));
    });

    req.end();

    req.on('error', error => callback(error));
  }
}
module.exports = GraphHelper;