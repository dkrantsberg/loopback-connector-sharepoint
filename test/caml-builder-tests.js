const {CamlBuilder} = require('../lib/caml-builder');
const xml2js = require('xml2js');
const {expect} = require('chai');

describe('CamlBuilder', () => {
  const camlBuilder = new CamlBuilder();
  describe('buildWhere()', () => {
    it.only('test xml2js', (done) => {
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
      const where = {prop1: 'Val 1'};
      const result = camlBuilder.buildWhere(where);
      const expectedResult = `<Eq><FieldRef Name="prop1"/><Value Type="Text">Val 1</Value></Eq>`;
      expect(result).to.eql(expectedResult);
    });

    it(`'and'`, () => {
      const where = {and: [{prop1: 'Val 1'}, {prop2: 'Val 2'}, {prop3: 'Val 3'}]};
      const result = camlBuilder.buildWhere(where);
      const expectedResult = `<And><Eq><FieldRef Name="prop1"/><Value Type="Text">Val 1</Value></Eq><And><Eq><FieldRef Name="prop2"/><Value Type="Text">Val 2</Value></Eq><And><Eq><FieldRef Name="prop3"/><Value Type="Text">Val 3</Value></Eq></And></And></And>`;
      expect(result).to.eql(expectedResult);
    });

    it(`'or' within 'and'`, () => {
      const where = {and: [{or: [{prop1: 'Val 1'}, {prop2: 'Val 2'}]}, {prop3: 'Val 3'}]};
      const result = camlBuilder.buildWhere(where);
      const expectedResult = `<And><Or><Eq><FieldRef Name="prop1"/><Value Type="Text">Val 1</Value></Eq><Or><Eq><FieldRef Name="prop2"/><Value Type="Text">Val 2</Value></Eq></Or></Or><And><Eq><FieldRef Name="prop3"/><Value Type="Text">Val 3</Value></Eq></And></And>`;
      expect(result).to.eql(expectedResult);
    });
  })
});