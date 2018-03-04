var docdbUtils = require('./docdbUtils');

function UserRepository(documentDBClient, databaseId, collectionId) {
  this.client = documentDBClient;
  this.databaseId = databaseId;
  this.collectionId = collectionId;

  this.database = null;
  this.collection = null;
  self = this;
}

UserRepository.prototype.init = function(callback) {
  var self = this;

  docdbUtils.getOrCreateDatabase(self.client, self.databaseId, function(err, db) {
    if (err) {
      callback(err);
    }

    self.database = db;
    docdbUtils.getOrCreateCollection(self.client, self.database._self, self.collectionId, function(err, coll) {
      if (err) {
        callback(err);
      }

      self.collection = coll;
      callback(null, coll);
    });
  });
};

UserRepository.prototype.find = function(querySpec, callback) {
  var self = this;
  self.client.queryDocuments(self.collection._self, querySpec).toArray(function(err, results) {
    if (err) {
      callback(err);
    } else {
      callback(null, results);
    }
  });
};

UserRepository.prototype.addItem = function(item, callback) {
  var self = this;
  item.date = Date.now();
  item.completed = false;
  self.client.createDocument(self.collection._self, item, function(err, doc) {
    if (err) {
      callback(err);
    } else {
      callback(null);
    }
  });
};

UserRepository.prototype.updateItem = function(itemId, callback) {
  var self = this;

  self.getItem(itemId, function(err, doc) {
    if (err) {
      callback(err);
    } else {
      doc.completed = true;
      self.client.replaceDocument(doc._self, doc, function(err, replaced) {
        if (err) {
          callback(err);
        } else {
          callback(null);
        }
      });
    }
  });
};

UserRepository.prototype.getItemByGraphUserId = function(itemId, callback){
  var self = this;

  var querySpec = {
    query: 'SELECT * FROM root r WHERE r.data.oid=@id',
    parameters: [{
      name: '@id',
      value: itemId
    }]
  };

  self.client.queryDocuments(self.collection._self, querySpec).toArray(function(err, results) {
    if (err) {
      callback(err);
    } else {
      callback(null, results[0]);
    }
  });
};

UserRepository.prototype.getItemBySubscriptionId = function(itemId, callback){
  var self = this;

  var querySpec = {
    query: 'SELECT * FROM root r WHERE r.data.subscriptionId=@id',
    parameters: [{
      name: '@id',
      value: itemId
    }]
  };

  self.client.queryDocuments(self.collection._self, querySpec).toArray(function(err, results) {
    if (err) {
      callback(err);
    } else {
      callback(null, results[0]);
    }
  });
};

UserRepository.prototype.getItem = function(itemId, callback) {
  var self = this;

  var querySpec = {
    query: 'SELECT * FROM root r WHERE r.id=@id',
    parameters: [{
      name: '@id',
      value: itemId
    }]
  };

  self.client.queryDocuments(self.collection._self, querySpec).toArray(function(err, results) {
    if (err) {
      callback(err);
    } else {
      callback(null, results[0]);
    }
  });
};

module.exports = UserRepository;