'use strict';
require('dotenv').config();
const path = require('path');
const nock = require('nock');
const timeout = 20000;
const juggler = require('loopback-datasource-juggler');
let DataSource = juggler.DataSource;

const config = {
  authConfig: {
    username: process.env.SP_USERNAME,
    password: process.env.SP_PASSWORD,
    online: true
  },
  siteUrl: process.env.SP_SITE_URL,
  debug: true
};

global.config = config;

let db;
let nockDone;

global.getDataSource = global.getSchema = (customConfig, customClass) => {
  const ctor = customClass || DataSource;
  db = new ctor(require('../'), customConfig || config);
  db.log = (a) => {
    console.log(a);
  };
  return db;
};

global.resetDataSourceClass = (ctor) => {
  DataSource = ctor || juggler.DataSource;
  const promise = db ? db.disconnect() : Promise.resolve();
  db = undefined;
  return promise;
};

if (!process.env.DISABLE_NOCK) {
  before(async () => {
    this.timeout(timeout);
    nock.back.fixtures = path.join(__dirname, 'nock-fixtures');
    nock.back.setMode('record');
    nock.enableNetConnect();
    // @ts-ignore
    global.nock = nock;
    nockDone = (await nock.back('sharepoint-connector.json', {
      before: beforeNock,
      afterRecord: filterNockOutput
    })).nockDone;
  });

  after(async () => {
    nockDone();
    this.timeout(5000);
  });
}

const beforeNock = (scope) => {
  // ignore request body when matching nock recorded fixtures
  scope.filteringRequestBody = (body, aRecordedBody) => {
    return aRecordedBody;
  };
};

const filterNockOutput = (outputs) => {
  // erase request body as it contains credentials
  return outputs.map(output => {
    output.body = '';
    return output;
  });
};
