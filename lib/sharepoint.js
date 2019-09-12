'use strict';
const {Connector} = require('loopback-connector');
const debug = require('debug')('loopback:connector:sharepoint');
const {bootstrap} = require('pnp-auth');
const {sp} = require('@pnp/sp');
const util = require('util');

function SharePointConnector(settings, dataSource) {
  Connector.call(this, 'sharepoint', settings);

  this.debug = settings.debug || debug.enabled;

  if (this.debug) {
    debug('Settings %j', settings);
  }
  this.dataSource = dataSource;
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
    })
  } else if ((self.dataSource.connecting) ) {
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
