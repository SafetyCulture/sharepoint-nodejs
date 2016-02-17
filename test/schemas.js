import { expect } from 'chai';
import jsen from 'jsen';
import _ from 'lodash';
import auditSchema from '../../src/sharepoint/schemas/audit.json';
import itemSchema from '../../src/sharepoint/schemas/item.json';
import auditBasic from './fixtures/auditBasic';
import itemBasic from './fixtures/itemBasic';
import { generateSchemaErrorMsg } from '../utils';

describe('Audit Schema', () => {
  const validate = jsen(auditSchema);

  it('should be a valid schema', () => {
    const validateSchema = jsen({'$ref': 'http://json-schema.org/draft-04/schema#'});
    expect(validateSchema(auditSchema)).to.equal(true, generateSchemaErrorMsg(validateSchema.errors));
  });

  it('should validate a valid audit', () => {
    expect(validate(auditBasic)).to.equal(true, generateSchemaErrorMsg(validate.errors));
  });

  it('should not validate an invalid audit', () => {
    const invalidAudit = _.assign({}, auditBasic, {
      'Audit Id': 'some fake id'
    });
    expect(validate(invalidAudit)).to.be.equal(false, 'Invalid data has validated incorrectly');
  });
});

describe('Item Schema', () => {
  const validate = jsen(itemSchema);

  it('should be a valid schema', () => {
    const validateSchema = jsen({'$ref': 'http://json-schema.org/draft-04/schema#'});
    expect(validateSchema(itemSchema)).to.equal(true, generateSchemaErrorMsg(validateSchema.errors));
  });

  it('should validate a valid item', () => {
    expect(validate(itemBasic)).to.equal(true, generateSchemaErrorMsg(validate.errors));
  });

  it('should not validate an invalid item', () => {
    const invalidItem = _.assign({}, itemBasic, {
      audit_id: 'some fake id'
    });
    expect(validate(invalidItem)).to.be.equal(false, 'Invalid data has validated incorrectly');
  });
});
