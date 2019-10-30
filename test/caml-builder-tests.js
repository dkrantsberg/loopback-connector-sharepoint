const {CamlBuilder} = require('../lib/caml-builder');
const xml2js = require('xml2js');
const {expect} = require('chai');

describe('CamlBuilder', () => {
  let camlBuilder;

  before(() => {
    const ds = global.getDataSource();
    const User = ds.define('User',
      {
        firstName: {type: String, sharepoint: {columnName: 'FirstName'}},
        lastName: {type: String, sharepoint: {columnName: 'LastName'}},
        email: {type: String, sharepoint: {columnName: 'Email'}},
        age: {type: Number, sharepoint: {columnName: 'Age'}}
      });
    camlBuilder = new CamlBuilder(User.definition);
  });

  describe('buildWhere()', () => {
    it('test xml2js', (done) => {
      const xml = `<OrderBy>
                        <FieldRef Name ='ows_ArchiveOrder' Ascending='TRUE'/>
                        <FieldRef Name ='ows_ArchiveOrder2' Ascending='FALSE'/>
                   </OrderBy>`;
      const parser = new xml2js.Parser({explicitArray: false});
      parser.parseString(xml, (err, result) => {
        const backToXml = (new xml2js.Builder()).buildObject(result);
        done();
      })
    });

    it('simple condition', () => {
      const where = {lastName: 'Doe'};
      const result = camlBuilder.buildWhere(where);
      expect(result).to.exist;
    });
    it('2 conditions with logical AND', () => {
      const where = {and: [{firstName: 'Joe'}, {lastName: 'Doe'}]};
      const result = camlBuilder.buildWhere(where);
      expect(result).to.exist;
    });
    it('3 conditions with logical AND', () => {
      const where = {and: [{firstName: 'Joe'}, {lastName: 'Doe'}, {age: 28}]};
      const result = camlBuilder.buildWhere(where);
      expect(result).to.exist;
    });
    it('3 conditions with logical AND', () => {
      const where =  {and: [{or: [{firstName: 'Joe'},  {lastName: 'Doe'}]}, {email: 'joe.doe@company.com'}]};
      const result = camlBuilder.buildWhere(where);
      expect(result).to.exist;
    });
  })
});
