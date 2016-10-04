import { expect } from 'chai';

import { listType } from '../src/lists.js';


describe('Lists', function test() {
  describe('formats the name correctly', () => {
    it('replaces slash with special SharePoint-escaped character', () => {
      expect(listType('Cats/Dogs')).to.be.equal('SP.Data.Cats_x002f_DogsListItem');
      expect(listType('Cats-Dogs')).to.be.equal('SP.Data.Cats_x002d_DogsListItem');
      expect(listType('Cats-Dogs/Birds')).to.be.equal('SP.Data.Cats_x002d_Dogs_x002f_BirdsListItem');
    });

    it('replaces spaces with symbol', () => {
      expect(listType('Cats Dogs')).to.be.equal('SP.Data.Cats_x0020_DogsListItem');
    });
  });
});
