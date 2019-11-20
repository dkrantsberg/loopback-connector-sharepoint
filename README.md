## loopback-connector-sharepoint [![Build Status](https://travis-ci.org/Synerzip/loopback-connector-sqlite.svg)](https://travis-ci.org/Synerzip/loopback-connector-sqlite)
[**LoopBack**](http://loopback.io/) is a highly-extensible, open-source Node.js framework that enables you to create dynamic end-to-end REST APIs with little or no coding. It also enables you to access data from major relational databases, MongoDB, SOAP and REST APIs.

**loopback-connector-sharepoint** is the Microsoft SharePoint connector module for [loopback-datasource-juggler](https://github.com/strongloop/loopback-datasource-juggler).

## Basic usage

#### Installation
Install the module using the command below in your projects root directory:
```sh
npm i loopback-connector-sharepoint
```

#### Configuration

Below is the sample datasource configuration file:

* `siteUrl`(string): SharePoint site url.
* `authConfig`(object): Object containing authentication credentials for SharePoint. See [node-sp-auth](https://github.com/s-KaiNet/node-sp-auth) and [node-sp-auth-config](https://github.com/koltyakov/node-sp-auth-config) documentation for different authentication strategies.
* `debug`: when true, prints debugging information (such as CAML queries) to console. 

```json
{
  "name": "sample-datasource",
  "connector": "sharepoint",
  "authConfig": {
    "username": "admin",
    "password": "secret",
    "online": true
  },
  "siteUrl": "https://sample.sharepoint.com/sites/my-site"
}
```


#### NOTE: Defining Models
SharePoint LB connector provides options for mapping between SharePoint lists and columns and Loopback models and their properties.
These options are set in the LB4 decorators inside `sharepoint` element. 

`list` - Name of SharePoint list to store model instances. If not specified then model class name is used. 
`columnName` - Name (InternalName) of SharePoint column to store model's property. If not specified then property name is used.

Example: user.model.ts
```typescript
import {Entity, property, model} from '@loopback/repository';
@model({
  settings: {
    sharepoint: {
      list: 'Users',
    },
  },
})

export class User extends Entity {
  @property({
    type: 'number',
    id: true,
    sharepoint: {
      columnName: 'ID',
    },
  })
  id?: number;

  @property({
    type: 'string',
    required: true,
    sharepoint: {
      columnName: 'Title',
    },
  })
  title: string;

  @property({
    type: 'string',
    sharepoint: {
      columnName: 'FirstName',
    },
  })
  firstName: string;

  @property({
    type: 'string',
    sharepoint: {
      columnName: 'LastName',
    },
  })
  lastName: string;

  @property({
    type: 'string',
    sharepoint: {
      columnName: 'Email',
    },
  })
  email: string;

  @property({
    type: 'number',
    sharepoint: {
      columnName: 'Age',
    },
  })
  age: number;

}
```
Notes: All SharePoint lists contain default columns: ID, GUID, Title, etc.. These columns are present in all lists and cannot be removed.  
You can map your model's properties to these columns. For identity properties `{id: true}` you have two options: map them to `ID` SP column, which is auto-generated integer, or to `GUID` SP column which is a 35-character UUID.
When mapping to GUID, you can set id value to your own generated GUID, SharePoint won't override it.       

## Debugging
loopback-connector-connector uses [debug](https://www.npmjs.com/package/debug) utility. To print debugging information you can set environment variable DEBUG=loopback-sharepoint-connector or DEBUG=*.
You can also set {debug: true} in the datasource configuration.
 

## Running the tests
* execute `npm install` for installing all the dependencies.
* execute `npm test` to run all the tests.
