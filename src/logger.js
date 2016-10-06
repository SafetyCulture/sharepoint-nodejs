import bunyan from 'bunyan';

export const log = bunyan.createLogger({
  name: 'sharepoint-core',
  serializers: {
    err: bunyan.stdSerializers.err
  }
});
