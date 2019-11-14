'use strict';
const {SPLib} = require('../lib/sp-lib');
const {expect} = require('chai');

describe('SPLib tests', () => {
  let spLib;

  before(() => {
    const ds = global.getDataSource();
    const User = ds.define('User',
      {
        firstName: {type: String, sharepoint: {columnName: 'FirstName'}},
        lastName: {type: String, sharepoint: {columnName: 'LastName'}},
        email: {type: String, sharepoint: {columnName: 'Email'}},
        age: {type: Number, sharepoint: {columnName: 'Age'}},
        startDate: {type: Date, sharepoint: {columnName: 'StartDate'}},
        isEmployee: {type: Boolean, sharepoint: {columnName: 'IsEmployee'}},
        displayName: {sharepoint: {columnName: 'DisplayName'}}
      });
    spLib = new SPLib(User.definition);
  });

  describe('buildWhere()', () => {
    it('should return empty string if no condition is specified', () => {
      expect(spLib.buildWhere()).to.equal('');
    });
    it('simple key:value condition', () => {
      const where = {lastName: 'Doe'};
      const result = spLib.buildWhere(where);
      const expectedResult = '<Where><Eq><FieldRef Name="LastName"/><Value Type="Text">Doe</Value></Eq></Where>';
      expect(result).to.eql(expectedResult);
    });
    it('\'inq\' condition', () => {
      const where = {ID: {inq: [4, 5, 6, 7]}};
      const result = spLib.buildWhere(where);
      const expectedResult = '<Where><In><FieldRef Name="ID"/><Values><Value Type="Number">4</Value><Value Type="Number">5</Value><Value Type="Number">6</Value><Value Type="Number">7</Value></Values></In></Where>';
      expect(result).to.eql(expectedResult);
    });
    it('2 conditions with logical AND', () => {
      const where = {and: [{firstName: 'Joe'}, {lastName: 'Doe'}]};
      const expectedResult = '<Where><And><Eq><FieldRef Name="FirstName"/><Value Type="Text">Joe</Value></Eq><Eq><FieldRef Name="LastName"/><Value Type="Text">Doe</Value></Eq></And></Where>';
      const result = spLib.buildWhere(where);
      expect(result).to.eql(expectedResult);
    });
    it('3 conditions with logical AND', () => {
      const where = {and: [{firstName: 'Joe'}, {lastName: 'Doe'}, {age: 28}]};
      const expectedResult = '<Where><And><Eq><FieldRef Name="FirstName"/><Value Type="Text">Joe</Value></Eq><And><Eq><FieldRef Name="LastName"/><Value Type="Text">Doe</Value></Eq><Eq><FieldRef Name="Age"/><Value Type="Number">28</Value></Eq></And></And></Where>';
      const result = spLib.buildWhere(where);
      expect(result).to.eql(expectedResult);
    });
    it('4 conditions with logical AND', () => {
      const where = {and: [{firstName: 'Joe'}, {lastName: 'Doe'}, {age: 28}, {email: 'joe.doe@company.com'}]};
      const expectedResult = '<Where><And><Eq><FieldRef Name="FirstName"/><Value Type="Text">Joe</Value></Eq><And><Eq><FieldRef Name="LastName"/><Value Type="Text">Doe</Value></Eq><And><Eq><FieldRef Name="Age"/><Value Type="Number">28</Value></Eq><Eq><FieldRef Name="Email"/><Value Type="Text">joe.doe@company.com</Value></Eq></And></And></And></Where>';
      const result = spLib.buildWhere(where);
      expect(result).to.eql(expectedResult);
    });
    it('combination of AND and OR conditions', () => {
      const where = {and: [{or: [{firstName: 'Joe'}, {lastName: 'Doe'}]}, {email: 'joe.doe@company.com'}]};
      const expectedResult = '<Where><And><Or><Eq><FieldRef Name="FirstName"/><Value Type="Text">Joe</Value></Eq><Eq><FieldRef Name="LastName"/><Value Type="Text">Doe</Value></Eq></Or><Eq><FieldRef Name="Email"/><Value Type="Text">joe.doe@company.com</Value></Eq></And></Where>';
      const result = spLib.buildWhere(where);
      expect(result).to.eql(expectedResult);
    });
    it('lt, lte, gt, gte', () => {
      const where = {and: [{age: {'gt': 20}}, {age: {gte: 21}}, {age: {lt: 100}}, {age: {lte: 99}}]};
      const expectedResult = '<Where><And><Gt><FieldRef Name="Age"/><Value Type="Number">20</Value></Gt><And><Geq><FieldRef Name="Age"/><Value Type="Number">21</Value></Geq><And><Lt><FieldRef Name="Age"/><Value Type="Number">100</Value></Lt><Leq><FieldRef Name="Age"/><Value Type="Number">99</Value></Leq></And></And></And></Where>';
      const result = spLib.buildWhere(where);
      expect(result).to.eql(expectedResult);
    });
    it('neq, like, contains', () => {
      const where = {and: [{firstName: {'neq': 'Joe'}}, {lastName: {like: 'Doe'}}, {email: {contains: 'doe'}}]};
      const expectedResult = '<Where><And><Neq><FieldRef Name="FirstName"/><Value Type="Text">Joe</Value></Neq><And><BeginsWith><FieldRef Name="LastName"/><Value Type="Text">Doe</Value></BeginsWith><Contains><FieldRef Name="Email"/><Value Type="Text">doe</Value></Contains></And></And></Where>';
      const result = spLib.buildWhere(where);
      expect(result).to.eql(expectedResult);
    });
    it('date field', () => {
      const result = spLib.buildWhere({startDate: {'gt': new Date('1/1/2019')}});
      const expectedResult = '<Where><Gt><FieldRef Name="StartDate"/><Value Type="DateTime">2019-01-01T05:00:00.000Z</Value></Gt></Where>';
      expect(result).to.eql(expectedResult);
    });
    it('boolean field', () => {
      const result = spLib.buildWhere({isEmployee: true});
      const expectedResult = '<Where><Eq><FieldRef Name="IsEmployee"/><Value Type="Boolean">1</Value></Eq></Where>';
      expect(result).to.eql(expectedResult);
    });
    it('field with no type should default to text', () => {
      const result = spLib.buildWhere({displayName: 'Joe Doe'});
      const expectedResult = '<Where><Eq><FieldRef Name="DisplayName"/><Value Type="Text">Joe Doe</Value></Eq></Where>';
      expect(result).to.eql(expectedResult);
    });
  });

  describe('buildViewFields()', () => {
    it('should generate expected ViewFields element', () => {
      const result = spLib.buildViewFields(['firstName', 'lastName']);
      const expectedResult = '<ViewFields><FieldRef Name="FirstName"/><FieldRef Name="LastName"/></ViewFields>';
      expect(result).to.eql(expectedResult);
    });
    it('should return empty string if no fields are specified', () => {
      const result = spLib.buildViewFields();
      expect(result).to.eql('');
    });
  });

  describe('buildOrderBy()', () => {
    it('should order by descending ID by default', () => {
      const result = spLib.buildOrderBy();
      const expectedResult = '<OrderBy><FieldRef Name="ID" Ascending="False"/></OrderBy>';
      expect(result).to.eql(expectedResult);
    });
    it('should generate expected OrderBy for single field', () => {
      const result = spLib.buildOrderBy('firstName');
      const expectedResult = '<OrderBy><FieldRef Name="FirstName"/></OrderBy>';
      expect(result).to.eql(expectedResult);
    });
    it('should generate expected OrderBy for single field descending order', () => {
      const result = spLib.buildOrderBy('firstName DESC');
      const expectedResult = '<OrderBy><FieldRef Name="FirstName" Ascending="FALSE"/></OrderBy>';
      expect(result).to.eql(expectedResult);
    });
    it('should generate expected OrderBy for single field ascending order', () => {
      const result = spLib.buildOrderBy('firstName ASC');
      const expectedResult = '<OrderBy><FieldRef Name="FirstName" Ascending="TRUE"/></OrderBy>';
      expect(result).to.eql(expectedResult);
    });
    it('should generate expected OrderBy for multiple fields field descending order', () => {
      const result = spLib.buildOrderBy(['lastName', 'firstName DESC']);
      const expectedResult = '<OrderBy><FieldRef Name="LastName"/><FieldRef Name="FirstName" Ascending="FALSE"/></OrderBy>';
      expect(result).to.eql(expectedResult);
    });
    it('should throw error if expression is not a string', () => {
      expect(() => spLib.buildOrderBy({foo: 'bar'}))
        .to.throw('Invalid order expression. Must be a string.');
    });
    it('should throw error if order is not either ASC or DESC', () => {
      expect(() => spLib.buildOrderBy(['lastName', 'firstName WRONG']))
        .to.throw('Invalid order direction. Must be either ASC or DESC');
    });
  });

  describe('buildRowLimit()', () => {
    it('should return expected RowLimit element', () => {
      const result = spLib.buildRowLimit(10);
      const expectedResult = '<RowLimit>10</RowLimit>';
      expect(result).to.eql(expectedResult);
    });
    it('should return empty string if no argum RowLimit element', () => {
      const result = spLib.buildRowLimit();
      expect(result).to.eql('');
    });
  });
});
