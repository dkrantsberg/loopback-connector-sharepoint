'use strict';
const {expect} = require('chai');

describe('SharePoint connector tests', () => {
  const ds = global.getDataSource();
  const User = ds.define('User',
    {
      id: {type: String, id: true, sharepoint: {columnName: 'GUID'}},
      firstName: {type: String, sharepoint: {columnName: 'FirstName'}},
      lastName: {type: String, sharepoint: {columnName: 'LastName'}},
      email: {type: String, sharepoint: {columnName: 'Email'}},
      jobTitle: {type: String, sharepoint: {columnName: 'JobTitle'}},
      age: {type: Number, sharepoint: {columnName: 'Age'}}
    }, {
      sharepoint: {
        list: 'TestUsers'
      }
    });

  const testUsers = [{
    id: '4d10b1fb-727d-4fb9-a6e8-8f1fa43bd677',
    firstName: 'Joe',
    lastName: 'Wimplesnatch',
    email: 'joe.wimp@company.com',
    jobTitle: 'Paranoid In Chief',
    age: 45
  }, {
    id: 'a2c06707-dcc1-42d3-8d28-f7f7692e7f12',
    firstName: 'Heidrich',
    lastName: 'Fancypants',
    email: 'h.smart@company.com',
    jobTitle: 'Light Bender',
    age: 25
  }, {
    id: 'fe69dbef-8b7b-466e-835f-3680a4da64be',
    firstName: 'Piggie',
    lastName: 'Oinks',
    email: 'pretty.flowers@company.com',
    jobTitle: 'PR Manager',
    age: 33
  }, {
    id: '4924bdbf-e837-47d9-8fa6-6782948513c3',
    firstName: 'Winston',
    lastName: 'Oldscratch',
    email: 'old.scratch@company.com',
    jobTitle: 'Brand Evangelist',
    age: 69
  }];

  before((done) => {
    ds.automigrate((err) => {
      expect(err).to.not.exist;
      done();
    });
  });

  after((done) => {
    ds.connector.sp.web.lists
      .getByTitle('TestUsers')
      .delete()
      .then(() => {
        done();
      })
      .catch(err => {
        done(err);
      });
  });
  it('should create users', async () => {
    await User.create(testUsers);
  });

  it('should get user count', async () => {
    const cnt = await User.count();
    expect(cnt).to.eql(testUsers.length);
  });

  it('should get all users back', async () => {
    const result = await User.find();
    expect(result).to.exist;
    expect(result).to.have.lengthOf(testUsers.length);
    const users = result.map((r) => (r.__data));
    expect(users).to.have.deep.members(testUsers);
  });

  it('should get top 2 users ordered by age', async () => {
    const result = await User.find({order: 'age', limit: 2});
    const ages = result.map((r) => (r.__data.age));
    expect(ages).to.eql([25, 33]);
  });

  it('should get next 2 users ordered by age', async () => {
    const result = await User.find({order: 'age', skip: 2, limit: 2});
    const ages = result.map((r) => (r.__data.age));
    expect(ages).to.eql([45, 69]);
  });

  it('should find by id', async () => {
    const result = await User.findById(testUsers[0].id);
    expect(result.__data).to.eql(testUsers[0]);
  });

  it('should update by condition', async () => {
    const result = await User.update({age: 69}, {lastName: 'Newscratch'});
    expect(result).to.eql({count: 1});
    const updated = await User.find({where: {age: 69}});
    const expectedResult = Object.assign(testUsers[3], {lastName: 'Newscratch'});
    expect(updated[0].__data).to.eql(expectedResult);
  });

  it('should replace by id', async () => {
    const newUser = {
      firstName: 'Jimmy',
      lastName: 'Different',
      email: 'imnotjoe@company.com',
      jobTitle: 'CTO',
      age: 48
    };
    await User.replaceById('4d10b1fb-727d-4fb9-a6e8-8f1fa43bd677', newUser);
    const replaced = await User.findById('4d10b1fb-727d-4fb9-a6e8-8f1fa43bd677');
    expect(replaced).to.include(newUser);
  });

  it('should delete by id', async () => {
    const deleteResult = await User.deleteById(testUsers[0].id);
    expect(deleteResult).to.eql({count: 1});
    expect(await User.findById(testUsers[0].id)).to.be.null;
  });

  it('should delete by condition', async () => {
    const deleteResult = await User.deleteAll({lastName: 'Newscratch'});
    expect(deleteResult).to.eql({count: 1});
    expect((await User.find({where: {lastName: 'Newscratch'}})).length).to.equal(0);
  });

  it('should delete all users if no condition is specified', async () => {
    const deleteResult = await User.deleteAll();
    expect(deleteResult).to.eql({count: 2});
    expect(await User.count()).to.equal(0);
  });
});

