import { expect } from 'chai';
import safetyCultureAudit from './fixtures/safetyCultureAudit';
import {
  toResponse,
  toItem,
  toAudit
} from '../../src/sharepoint/transformers';

describe('SharePoint-transformers', () => {
  describe('#toResponse', () => {
    it('should return a correctly transformed selected response', () => {
      const responses = safetyCultureAudit.items[2].responses;
      const result = toResponse(responses);
      const expected = {
        'R Id': '8bcfbf00-e11b-11e1-9b23-0800200c9a66',
        'R Type': 'text',
        'R Label': 'Yes',
        'R Short Label': 'Y',
        'R Colour': '162,182,58',
        'R Image': '',
        'R Enable Score': true,
        'R Score': 1
      };

      expect(result).to.deep.equal(expected);
    });

    it('should default to type: label', () => {
      const responses = safetyCultureAudit.header_items[4].responses;
      const result = toResponse(responses);
      const expected = {
        'R Type': 'datetime',
        'R Label': '2015-09-19T04:08:16.382Z'
      };

      expect(result).to.deep.equal(expected);
    });
  });

  describe('#toItem', () => {
    it('should return a correctly transformed item', () => {
      const item = safetyCultureAudit.items[4];
      const result = toItem(item);
      const expected = {
        '__metadata': { type: 'SP.Data.SafetyCulture_x0020_ItemsListItem' },
        'Title': 'The ferry is battled and veiny?',
        'Parent Id': 'f9b1af05-fb1b-4470-bbf8-29a409db964d',
        'Item Id': '90c41537-7cec-4419-9297-cc0dc13c8fe2',
        'Label': 'The ferry is battled and veiny?',
        'Type': 'question',
        'Score': 1,
        'Max Score': 1,
        'Percentage': 100,
        'R Id': '8bcfbf00-e11b-11e1-9b23-0800200c9a66',
        'R Type': 'text',
        'R Label': 'Yes',
        'R Short Label': 'Y',
        'R Colour': '162,182,58',
        'R Enable Score': true,
        'R Score': 1
      };

      expect(result).to.deep.equal(expected);
    });

    it('should return a correctly transformed header item', () => {
      const item = safetyCultureAudit.header_items[5];
      const result = toItem(item);
      const expected = {
        '__metadata': { type: 'SP.Data.SafetyCulture_x0020_ItemsListItem' },
        'Title': 'Prepared by',
        'Parent Id': '10E39E36-B1EE-4E22-BCB1-F358E8AE7151',
        'Item Id': 'f3245d43-ea77-11e1-aff1-0800200c9a66',
        'Label': 'Prepared by',
        'Type': 'textsingle',
        'R Type': 'text',
        'R Label': 'Voluptas et excepturi fuga et sint quis. Tempora sed inventore enim totam. Aliquid accusantium magnam quia numquam consequuntur nobis laboriosam.'
      };

      expect(result).to.deep.equal(expected);
    });

    it('must ensure text type responses are strings', () => {
      const item = safetyCultureAudit.header_items[2];
      const result = toItem(item);
      const expected = {
        '__metadata': { type: 'SP.Data.SafetyCulture_x0020_ItemsListItem' },
        'Title': 'Document No.',
        'Parent Id': '10E39E36-B1EE-4E22-BCB1-F358E8AE7151',
        'Item Id': 'f3245d46-ea77-11e1-aff1-0800200c9a66',
        'Label': 'Document No.',
        'Type': 'textsingle',
        'R Type': 'text',
        'R Label': '397051'
      };

      expect(result).to.deep.equal(expected);
    });
  });

  describe('#toAudit', () => {
    it('must return a correctly transformed audit', () => {
      const result = toAudit(safetyCultureAudit, [0]);
      const expected = {
        '__metadata': { type: 'SP.Data.SafetyCulture_x0020_AuditsListItem' },
        'Audit Id': 'audit_bf885775e4364ca4be1c1c044cc67d11',
        'Title': 'Halting Organ Audit',
        'Score': 178,
        'Total Score': 331,
        'Score Percentage': 53.776,
        'Date Modified': '2015-09-19T04:08:16.382Z',
        'Date Started': '2015-09-19T04:08:16.382Z',
        'Duration': 0,
        'SafetyCulture Owner': 'Nicholas Matenaar',
        'SafetyCulture Author': 'Nicholas Matenaar',
        'Template Name': 'Unkenned Eggnog Inspection',
        'ItemsId': { results: [0] }
      };

      expect(result).to.deep.equal(expected);
    });

    it('must include passed in item ids', () => {
      const result = toAudit(safetyCultureAudit, [1, 2, 3]);

      expect(result.ItemsId.results).to.deep.equal([1, 2, 3]);
    });
  });
});
