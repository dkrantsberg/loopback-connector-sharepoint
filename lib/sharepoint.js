'use strict';
const {Connector} = require('loopback-connector');
const debug = require('debug')('loopback:connector:sharepoint');
const {bootstrap} = require('pnp-auth');
const {sp, FieldTypes} = require('@pnp/sp');
const util = require('util');
const _ = require('lodash');
const {SPLib} = require('./sp-lib');
const Bluebird = require('bluebird');

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
SharePointConnector.prototype.connect = function(callback) {
  const self = this;
  if (self.sp) {
    process.nextTick(() => {
      if (callback) callback(null, self.sp);
    });
  } else {
    bootstrap(sp, self.settings.authConfig, self.settings.siteUrl);
    self.sp = sp;
    callback(null, self.sp);
  }
};

exports.initialize = function initializeDataSource(dataSource, callback) {
  const settings = dataSource.settings;
  dataSource.connector = new SharePointConnector(settings, dataSource);
  if (callback) {
    dataSource.connector.connect(callback);
  }
};

/**
 * Create a new model instance for the given data
 * @param {String} modelName The model name
 * @param {Object} data The model data
 * @param {Function} [callback] The callback function
 */
SharePointConnector.prototype.create = function(modelName, data, options, callback) {
  const self = this;
  if (self.debug) {
    debug('create', modelName, data);
  }
  const spItem = this.toSPItem(modelName, data);
  const idProp = this.getIdPropertyName(modelName);
  sp.web.lists.getByTitle(this.getSPListTitle(modelName)).items.add(spItem)
    .then((result) => {
      const lbEntity = this.fromSPItem(modelName, result.data);
      callback(null, lbEntity[idProp]);
    });
};

/**
 * Update all matching instances
 * @param {String} modelName The model name
 * @param {Object} where The search criteria
 * @param {Object} data The property/value pairs to be updated
 * @callback {Function} cb Callback function
 */
SharePointConnector.prototype.update =
  SharePointConnector.prototype.updateAll = function updateAll(modelName, where, data, options, callback) {
    const self = this;
    const listTitle = self.getSPListTitle(modelName);
    if (self.debug) {
      debug('updateAll', modelName, where, data);
    }
    const spData = self.toSPProperties(modelName, data);
    let affectedRows = 0;
    self.getItemsFiltered(modelName, {where, fields: ['ID']})
      .then(items => {
        const batch = sp.createBatch();
        for (const item of items) {
          sp.web.lists.getByTitle(listTitle).items.getById(item.ID).inBatch(batch).update(spData).then(() => {
            affectedRows++;
          });
        }
        return batch.execute();
      })
      .then((res) => {
        callback(null, {count: affectedRows});
      })
      .catch(err => {
        callback(err);
      });
  };

/**
 * Replace properties for the model instance data
 * @param {String} modelName The model name
 * @param {*} id The instance id
 * @param {Object} data The model data
 * @param {Object} options The options object
 * @param {Function} [cb] The callback function
 */
SharePointConnector.prototype.replaceById = function replace(modelName, id, data, options, cb) {
  if (this.debug) debug('replace', modelName, id, data);
  this.update(modelName, {id}, data, options, cb);
};

/**
 * Count the number of instances for the given model
 *
 * @param {String} modelName The model name
 * @param {Function} [callback] The callback function
 * @param {Object} filter The filter for where
 *
 */
SharePointConnector.prototype.count = function count(modelName, where, options, callback) {
  const self = this;
  if (self.debug) {
    debug('count', modelName, where);
  }
  self.getItemsFiltered(modelName, {where, fields: ['ID']})
    .then(items => {
      callback(null, items.length);
    })
    .catch(err => {
      callback(err, count);
    });
};

/**
 * Delete all instances for the given model
 * @param {String} modelName The model name
 * @param {Object} [where] The filter for where
 * @param {Function} [callback] The callback function
 */
SharePointConnector.prototype.destroyAll = function(modelName, where, options, callback) {
  const self = this;
  const listTitle = self.getSPListTitle(modelName);
  if (self.debug) {
    debug('destroyAll', modelName, where);
  }
  let affectedRows = 0;
  self.getItemsFiltered(modelName, {where, fields: ['ID']})
    .then(items => {
      const batch = sp.createBatch();
      for (const item of items) {
        sp.web.lists.getByTitle(listTitle).items.getById(item.ID).inBatch(batch).delete().then(() => {
          affectedRows++;
        });
      }
      return batch.execute();
    })
    .then((res) => {
      callback(null, {count: affectedRows});
    })
    .catch(err => {
      callback(err);
    });
};

/**
 * Find matching model instances by the filter
 *
 * @param {String} modelName The model name
 * @param {Object} filter The filter object
 * @param {Function} [callback] The call back function
 */
SharePointConnector.prototype.all = function(modelName, filter, options, callback) {
  const self = this;
  if (self.debug) {
    debug('all', modelName);
  }
  if (filter.skip) {
    /*
    Due to SharePoint lack of support for paging by skip/limit filters the following approach is used:
       1. Get only ID's of ALL the items matching specified criteria
       2. Take subset of the IDs by applying 'skip' and 'limit' filters to them
       3. Then get entire items for the subset of IDs
    */
    const originalFields = _.clone(filter.fields);
    const originalLimit = filter.limit;
    filter.fields = ['ID'];
    filter.limit = null;
    this.getItemsFiltered(modelName, filter)
      .then(items => {
        if (_.isEmpty(items)) {
          callback(null, []);
        }
        const end = originalLimit ? filter.skip + originalLimit : items.length;
        const ids = _.slice(_.map(items, 'ID'), filter.skip, end);
        filter.where = {ID: {inq: ids}};
        filter.fields = originalFields;
        return filter;
      })
      .then(filter => {
        return this.getItemsFiltered(modelName, filter);
      })
      .then(items => {
        const entities = _.map(items, (item) => this.fromSPItem(modelName, item));
        callback(null, entities);
      });
  } else {
    this.getItemsFiltered(modelName, filter, true)
      .then(items => {
        const entities = _.map(items, (item) => this.fromSPItem(modelName, item));
        callback(null, entities);
      });
  }
};

SharePointConnector.prototype.getItemsFiltered = function(modelName, filter, expandLookups) {
  const self = this;
  const spLib = new SPLib(self._models[modelName]);
  const camlQuery = spLib.buildQuery(filter);
  if (self.debug) {
    debug(`CAML: ${JSON.stringify(camlQuery)}`);
  }
  if (expandLookups) {
    // This will include FieldValuesAsText object which contains values of lookup fields (like users)
    return sp.web.lists.getByTitle(self.getSPListTitle(modelName)).getItemsByCAMLQuery(camlQuery, 'FieldValuesAsText');
  }
  return sp.web.lists.getByTitle(self.getSPListTitle(modelName)).getItemsByCAMLQuery(camlQuery);
};

/**
 * Perform automigrate for the given models. It drops the corresponding collections
 * and calls createIndex
 * @param {String[]} [models] A model name or an array of model names. If not present, apply to all models
 * @param {Function} [cb] The callback function
 */
SharePointConnector.prototype.automigrate = function(models, cb) {
  const self = this;
  if (self.debug) {
    debug('automigrate');
  }
  return Bluebird.mapSeries(models, model => {
    return self.createList(model);
  })
    .then(() => {
      cb();
    });
};

SharePointConnector.prototype.createList = function(modelName) {
  const listTitle = this.getSPListTitle(modelName);

  const model = this._models[modelName];
  const spLib = new SPLib(this._models[modelName]);

  return sp.web.lists.add(listTitle, '')
    .then(res => {
      return sp.web.lists.getByTitle(listTitle).fields.get();
    })
    .then((defaultFields) => {
      const existingSPFields = _.map(defaultFields, 'InternalName');
      const addFieldsBatch = sp.web.createBatch();
      const addFieldsToDefaultViewBatch = sp.web.createBatch();

      const list = sp.web.lists.getByTitle(listTitle);
      for (const prop of Object.keys(model.properties)) {
        const spFieldName = spLib.getSPFieldName(prop);
        // skip fields which already exist
        if (_.includes(existingSPFields, spFieldName)) {
          continue;
        }
        const fieldType = spLib.getSPFieldType(prop);
        const fieldProps = {FieldTypeKind: FieldTypes[fieldType]};
        list.fields.inBatch(addFieldsBatch).add(spFieldName, `SP.Field${fieldType}`, fieldProps);
        list.defaultView.fields.inBatch(addFieldsToDefaultViewBatch).add(spFieldName);
      }
      // add fields to list and then add those fields to default view
      return addFieldsBatch.execute()
        .then(() => {
          return addFieldsToDefaultViewBatch.execute();
        });
    });
};

Connector.defineAliases(SharePointConnector.prototype, 'find', 'findById');
/*!
 * Convert the data from SharePoint to LB entity
 *
 * @param {String} modelName The model name
 * @param {Object} data The data from SharePoint
 */
SharePointConnector.prototype.fromSPItem = function(modelName, spItem) {
  if (!spItem) {
    return null;
  }
  const lbEntity = {};
  const modelInfo = this._models[modelName];
  for (const propName in modelInfo.properties) {
    const spColumnName = this.getSPColumnName(modelName, propName);
    _.set(lbEntity, propName, _.get(spItem, spColumnName));
  }
  return lbEntity;
};

/*!
 * Convert the data from LB entity to SharePoint item
 *
 * @param {String} modelName The model name
 * @param {Object} lbItem The LoopBack model instance
 */
SharePointConnector.prototype.toSPItem = function(modelName, lbEntity) {
  const spItem = {};
  const modelInfo = this._models[modelName];
  for (const propName in modelInfo.properties) {
    const spColumnName = this.getSPColumnName(modelName, propName);
    _.set(spItem, spColumnName, _.get(lbEntity, propName));
  }
  return spItem;
};

/*!
 * Convert an object containing LB model properties to one with SharePoint properties
 *
 * @param {String} modelName The model name
 * @param {Object} lbProperties and object containing key-values where keys are LoopBack model properties
 */
SharePointConnector.prototype.toSPProperties = function(modelName, lbProperties) {
  if (!lbProperties) {
    return null;
  }
  const spProperties = {};

  for (const propName in lbProperties) {
    const spColumnName = this.getSPColumnName(modelName, propName);
    spProperties[spColumnName] = lbProperties[propName];
  }
  return spProperties;
};

/*
 * Gets the title of the SharePoint list for the specified LB model
 *
 * @param {String} modelName The model name
 */
SharePointConnector.prototype.getSPListTitle = function(modelName) {
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
SharePointConnector.prototype.getSPColumnName = function(modelName, propName) {
  const modelInfo = this._models[modelName];
  const spSPPropName = _.get(modelInfo, `properties.${propName}.sharepoint.columnName`);
  return spSPPropName || propName;
};

SharePointConnector.prototype.getIdPropertyName = function(modelName) {
  const modelInfo = this._models[modelName];
  for (const propName in modelInfo.properties) {
    if (modelInfo.properties[propName].id) {
      return propName;
    }
  }
};

exports.SharePointConnector = SharePointConnector;
