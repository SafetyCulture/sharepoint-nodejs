import { expect } from 'chai';

import { listType } from '../src/lists.js';


describe('Lists', function test() {
  describe('formats the name correctly', () => {
    it('replace slash with nothing', () => {
      expect(listType('Cats/Dogs')).to.be.equal('SP.Data.CatsDogsItem');
      expect(listType('Cats-Dogs')).to.be.equal('SP.Data.CatsDogsItem');
      expect(listType('Cats-Dogs/Birds')).to.be.equal('SP.Data.CatsDogsBirdsItem');
    });

    it('replaces spaces with symbol', () => {
      expect(listType('Cats Dogs')).to.be.equal('SP.Data.Cats_x0020_DogsItem');
    });
  });
});
