import {expect} from 'chai';
import {asStringWithoutSensitiveFields} from '../src/oauth2';

describe('when logging', () => {
  it('should omit authorisation code from response', () => {
    const request = {
      code: 'value'
    };

    expect(asStringWithoutSensitiveFields(request)).to.be.equal('code=omitted');
  });

  it('should omit refresh_token from request ', () => {
    const request = {
      refresh_token: 'value'
    };

    expect(asStringWithoutSensitiveFields(request)).to.be.equal('refresh_token=omitted');
  });

  it('should omit client_secret from request ', () => {
    const request = {
      client_secret: 'secret'
    };

    expect(asStringWithoutSensitiveFields(request)).to.be.equal('client_secret=omitted');
  });

  it('should omit all sensitive fields and leave other fields', () => {
    const request = {
      refresh_token: 'value',
      client_id: 'keep',
      client_secret: 'secret',
      code: 'something'
    };

    expect(asStringWithoutSensitiveFields(request)).to.be.equal('refresh_token=omitted&client_id=keep&client_secret=omitted&code=omitted');
    expect(request.refresh_token).to.be.equal('value');
  });
});
