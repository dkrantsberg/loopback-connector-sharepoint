'use strict';
const juggler = require('loopback-datasource-juggler');
let DataSource = juggler.DataSource;

const config = {
  authConfig: {
    username: process.env.SP_USERNAME,
    password: process.env.SP_PASSWORD,
    online: true
  },
  siteUrl: process.env.SP_SITE_URL
};

global.config = config;

let db;
global.getDataSource = global.getSchema = function (customConfig, customClass) {
  const ctor = customClass || DataSource;
  db = new ctor(require('../'), customConfig || config);
  db.log = function (a) {
    console.log(a);
  };
  return db;
};

global.resetDataSourceClass = function (ctor) {
  DataSource = ctor || juggler.DataSource;
  const promise = db ? db.disconnect() : Promise.resolve();
  db = undefined;
  return promise;
};
