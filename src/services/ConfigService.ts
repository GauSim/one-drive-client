

interface IApiCreds {
  "redirectUrl": string;
  "clientID": string;
  "clientSecret": string;
  "identityMetadata": string;
  "allowHttpForRedirectUrl": boolean;
  "responseType": string;
  "validateIssuer": boolean;
  "responseMode": string;
  "scope": string[];
}

export class ConfigService {

  constructor() {

  }

  getApiCreds(): IApiCreds {
    return require('../../secrets');
  }
}

