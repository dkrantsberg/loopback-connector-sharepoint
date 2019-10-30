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
   //  it.skip('test xml2js', (done) => {
   //    const xml = `   <Where>
   //    <And>
   //       <Eq>
   //          <FieldRef Name='FirstName' />
   //          <Value Type='Text'>Joe</Value>
   //       </Eq>
   //       <Neq>
   //          <FieldRef Name='LastName' />
   //          <Value Type='Text'>Doe</Value>
   //       </Neq>
   //    </And>
   // </Where>`;
   //    const parser = new xml2js.Parser({explicitArray: false});
   //    parser.parseString(xml, (err, result) => {
   //      const backToXml = (new xml2js.Builder()).buildObject(result);
   //      done();
   //    })
   //  });

    it('simple key:value condition', () => {
      const where = {lastName: 'Doe'};
      const result = camlBuilder.buildWhere(where);
      const expectedResult = `<Where><Eq><FieldRef Name="LastName"/><Value Type="Text">Doe</Value></Eq></Where>`;
      expect(result).to.eql(expectedResult);
    });
    it(`'inq' condition`, () => {
      const where = {ID: {inq: [4, 5, 6, 7]}};
      const result = camlBuilder.buildWhere(where);
      const expectedResult = `<Where><In><FieldRef Name="ID"/><Values><Value Type="Number">4</Value><Value Type="Number">5</Value><Value Type="Number">6</Value><Value Type="Number">7</Value></Values></In></Where>`;
      expect(result).to.eql(expectedResult);
    });
    it('2 conditions with logical AND', () => {
      const where = {and: [{firstName: 'Joe'}, {lastName: 'Doe'}]};
      const expectedResult = `<Where><And><Eq><FieldRef Name="FirstName"/><Value Type="Text">Joe</Value></Eq><Eq><FieldRef Name="LastName"/><Value Type="Text">Doe</Value></Eq></And></Where>`;
      const result = camlBuilder.buildWhere(where);
      expect(result).to.eql(expectedResult);
    });
    it('3 conditions with logical AND', () => {
      const where = {and: [{firstName: 'Joe'}, {lastName: 'Doe'}, {age: 28}]};
      const expectedResult = `<Where><And><Eq><FieldRef Name="FirstName"/><Value Type="Text">Joe</Value></Eq><And><Eq><FieldRef Name="LastName"/><Value Type="Text">Doe</Value></Eq><Eq><FieldRef Name="Age"/><Value Type="Number">28</Value></Eq></And></And></Where>`;
      const result = camlBuilder.buildWhere(where);
      expect(result).to.eql(expectedResult);
    });
    it('4 conditions with logical AND', () => {
      const where = {and: [{firstName: 'Joe'}, {lastName: 'Doe'}, {age: 28}, {email: 'joe.doe@company.com'}]};
      const expectedResult = `<Where><And><Eq><FieldRef Name="FirstName"/><Value Type="Text">Joe</Value></Eq><And><Eq><FieldRef Name="LastName"/><Value Type="Text">Doe</Value></Eq><And><Eq><FieldRef Name="Age"/><Value Type="Number">28</Value></Eq><Eq><FieldRef Name="Email"/><Value Type="Text">joe.doe@company.com</Value></Eq></And></And></And></Where>`;
      const result = camlBuilder.buildWhere(where);
      expect(result).to.eql(expectedResult);
    });
    it('combination of AND and OR conditions', () => {
      const where = {and: [{or: [{firstName: 'Joe'}, {lastName: 'Doe'}]}, {email: 'joe.doe@company.com'}]};
      const expectedResult = `<Where><And><Or><Eq><FieldRef Name="FirstName"/><Value Type="Text">Joe</Value></Eq><Eq><FieldRef Name="LastName"/><Value Type="Text">Doe</Value></Eq></Or><Eq><FieldRef Name="Email"/><Value Type="Text">joe.doe@company.com</Value></Eq></And></Where>`;
      const result = camlBuilder.buildWhere(where);
      expect(result).to.eql(expectedResult);
    });
  })

  describe('buildViewFields()', () => {
    it('should generate expected ViewFields element', () => {
      const result = camlBuilder.buildViewFields(['firstName', 'lastName']);
      const expectedResult = '<ViewFields><FieldRef Name="FirstName"/><FieldRef Name="LastName"/></ViewFields>';
      expect(result).to.eql(expectedResult);
    });
    it('should return empty string if no fields are specufied', () => {
      const result = camlBuilder.buildViewFields();
      expect(result).to.eql('');
    })
  });

  describe('buildOrderBy()', () => {
    it('should order by descending ID by default', () => {
      const result = camlBuilder.buildOrderBy();
      const expectedResult = '<OrderBy><FieldRef Name="ID" Ascending="False"/></OrderBy>';
      expect(result).to.eql(expectedResult);
    });
    it('should generate expected OrderBy for single field', () => {
      const result = camlBuilder.buildOrderBy('firstName');
      const expectedResult = '<OrderBy><FieldRef Name="FirstName"/></OrderBy>';
      expect(result).to.eql(expectedResult);
    });
    it('should generate expected OrderBy for single field descending order', () => {
      const result = camlBuilder.buildOrderBy('firstName DESC');
      const expectedResult = '<OrderBy><FieldRef Name="FirstName" Ascending="FALSE"/></OrderBy>';
      expect(result).to.eql(expectedResult);
    });
    it('should generate expected OrderBy for single field ascending order', () => {
      const result = camlBuilder.buildOrderBy('firstName ASC');
      const expectedResult = `<OrderBy><FieldRef Name="FirstName" Ascending="TRUE"/></OrderBy>`;
      expect(result).to.eql(expectedResult);
    });
    it('should generate expected OrderBy for multiple fields field descending order', () => {
      const result = camlBuilder.buildOrderBy(['lastName', 'firstName DESC']);
      const expectedResult = '<OrderBy><FieldRef Name="LastName"/><FieldRef Name="FirstName" Ascending="FALSE"/></OrderBy>';
      expect(result).to.eql(expectedResult);
    });
  });

  describe('buildRowLimit()', () => {
     it('should return expected RowLimit element', () => {
       const result = camlBuilder.buildRowLimit(10);
       const expectedResult = '<RowLimit>10</RowLimit>';
       expect(result).to.eql(expectedResult);
     });
    it('should return empty string if no argum RowLimit element', () => {
      const result = camlBuilder.buildRowLimit();
      expect(result).to.eql('');
    });
  });
});
