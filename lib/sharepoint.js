'use strict';
const {Connector} = require('loopback-connector');
const debug = require('debug')('loopback:connector:sharepoint');
const {bootstrap} = require('pnp-auth');
const {sp, CamlQuery} = require('@pnp/sp');
const {getGUID} = require('@pnp/common');
const util = require('util');
const _ = require('lodash');
const async = require('async');
const {CamlBuilder} = require('./caml-builder');

function SharePointConnector(settings, dataSource) {
  Connector.call(this, 'sharepoint', settings);

  this.debug = settings.debug || debug.enabled;

  if (this.debug) {
    debug('Settings %j', settings);
  }
  this.dataSource = dataSource;
  this._models = this._models || this.dataSource.modelBuilder.definitions;
}

util.inherits(SharePointConnector, Connector);

/**
 * Connect to SharePoint
 * @param {Function} [callback] The callback function
 *
 * @callback callback
 * @param {Error} err The error object
 * @param {Sp} db The sharepoint object
 */
SharePointConnector.prototype.connect = function (callback) {
  const self = this;

  if (self.db) {
    process.nextTick(function () {
      if (callback) callback(null, self.db);
    });
  } else if ((self.dataSource.connecting)) {
    self.dataSource.once('connected', function () {
      process.nextTick(function () {
        if (callback) {
          callback(null, self.db);
        }
      });
    });
  } else {
    bootstrap(sp, self.settings.authConfig, self.settings.siteUrl);
    self.db = sp;
    callback(null, self.db);
  }
};

exports.initialize = function initializeDataSource(dataSource, callback) {
  const settings = dataSource.settings;

  dataSource.connector = new SharePointConnector(settings, dataSource);

  if (callback) {
    if (settings.lazyConnect) {
      process.nextTick(function () {
        callback();
      });
    } else {
      dataSource.connector.connect(callback);
    }
  }
};

SharePointConnector.prototype.create = function (modelName, data, options, callback) {
  const self = this;
  if (self.debug) {
    debug('create', modelName, data);
  }
  const spItem = this.toSPItem(modelName, data);
  const guid = getGUID();
  spItem.GUID = guid;
  sp.web.lists.getByTitle(this.getSPListTitle(modelName)).items.add(spItem)
    .then((result) => {
      const lbEntity = this.fromSPItem(modelName, result.data);
      callback(null, lbEntity);
    });
};

SharePointConnector.prototype.destroyAll = function (modelName, where, options, callback) {
  const model = this._models[modelName];
  let itemsToDelete;
  if (_.isEmpty(where)) {
    itemsToDelete = sp.web.lists.getByTitle(this.getSPListTitle(modelName)).items.select('ID').getAll();
  } else {
    const camlBuilder = new CamlBuilder(model);
    const camlQuery = camlBuilder.buildQuery({where});
    itemsToDelete = sp.web.lists.getByTitle(this.getSPListTitle(modelName)).getItemsByCAMLQuery(camlQuery);
  }
  itemsToDelete
    .then((items) => {
      async.each(items, (item, cb) => {
        sp.web.lists.getByTitle(this.getSPListTitle(modelName))
          .items
          .getById(item.ID)
          .delete()
          .then(() => {
            cb();
          });
      }, (err) => {
        if (err) {
          callback(err);
        } else {
          callback(null, {count: items.length});
        }
      });
    })
    .catch(err => {
      callback(err);
    });
};


SharePointConnector.prototype.destroy = function (modelName, id, options, callback) {
  const spItems = sp.web.lists.getByTitle(this.getSPListTitle(modelName)).items;
  const modelInfo = this._models[modelName];
  const idColumn = this.getSPColumnName(modelInfo, this.getIdPropertyName(modelInfo));
  if (idColumn === 'ID') {
    spItems.getById(id).delete().then(() => {
      callback();
    });
  } else {
    spItems
      .top(1)
      .filter(`${idColumn} eq '${id}'`)
      .get()
      .then((items) => {
        items[0].delete().then(() => {
          callback();
        });
      });
  }
};


/**
 * Find matching model instances by the filter
 *
 * @param {String} modelName The model name
 * @param {Object} filter The filter
 * @param {Function} [callback] The call back function
 */
SharePointConnector.prototype.all = function (modelName, filter, options, callback) {
  const self = this;
  if (self.debug) {
    debug('find', modelName, id);
  }
  let getItems;
  if (!filter) {
    getItems = sp.web.lists.getByTitle(this.getSPListTitle(modelName)).items.getAll();
  } else {
    const camlBuilder = new CamlBuilder(this._models[modelName]);
    const camlQuery = camlBuilder.buildQuery(filter);
    getItems = sp.web.lists.getByTitle(this.getSPListTitle(modelName)).getItemsByCAMLQuery();
  }
  getItems.then(items => {
    const entities = _.map(items, (item) => this.fromSPItem(modelName, item));
    callback(null, entities);
  });
};

SharePointConnector.prototype.find = function (modelName, id, options, callback) {
  const self = this;
  if (self.debug) {
    debug('find', modelName, id);
  }
  const idColumn = this.getSPColumnName(modelInfo, this.getIdPropertyName(modelName));
  sp.web.lists.getByTitle(this.getSPListTitle(modelName)).items
    .top(1)
    .filter(`${idColumn} eq '${id}'`).get().then((items) => {
    if (items.length === 0) {
      throw new Error(`Item not found`);
    } else {
      const entity = this.fromSPItem(items[0]);
      callback(null, entity);
    }
  });
};
Connector.defineAliases(SharePointConnector.prototype, 'find', 'findById');
/*!
 * Convert the data from SharePoint to LB entity
 *
 * @param {String} modelName The model name
 * @param {Object} data The data from SharePoint
 */
SharePointConnector.prototype.fromSPItem = function (modelName, spItem) {
  if (!spItem) {
    return null;
  }
  const modelInfo = this._models[modelName];

  let lbEntity = {};

  for (let propName in modelInfo.properties) {
    const spColumnName = this.getSPColumnName(modelInfo, propName);
    _.set(lbEntity, propName, _.get(spItem, spColumnName));
  }
  return lbEntity;
};


/*!
 * Convert the data from LB entity to SharePoint item
 *
 * @param {String} modelName The model name
 * @param {Object} lbItem The loopback model instance
 */
SharePointConnector.prototype.toSPItem = function (modelName, lbEntity) {
  if (!lbEntity) {
    return null;
  }
  const modelInfo = this._models[modelName];

  let spItem = {};

  for (let propName in modelInfo.properties) {
    const spColumnName = this.getSPColumnName(modelInfo, propName);
    _.set(spItem, spColumnName, _.get(lbEntity, propName));
  }
  return spItem;
};


/*!
 * Gets the title of the SharePoint list for the specified LB model
 *
 * @param {String} modelName The model name
 */
SharePointConnector.prototype.getSPListTitle = function (modelName) {
  const modelInfo = this._models[modelName];
  const listTitle = _.get(modelInfo, 'settings.sharepoint.list');
  return listTitle || modelName;
};

/*!
 * Gets the SharePoint list column name for specified LB model property
 *
 * @param {Object} modelInfo The model definition
 * @param {String} propName Property name
 */
SharePointConnector.prototype.getSPColumnName = function (modelInfo, propName) {
  const spSPPropName = _.get(modelInfo, `properties.${propName}.sharepoint.columnName`);
  return spSPPropName || propName;
};


SharePointConnector.prototype.getIdPropertyName = function (modelInfo) {
  for (let propName in modelInfo.properties) {
    if (modelInfo.properties[propName].id) {
      return propName;
    }
  }
};

exports.SharePointConnector = SharePointConnector;
