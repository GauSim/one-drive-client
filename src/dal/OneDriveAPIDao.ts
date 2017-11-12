import request = require('superagent');
import logger = require('morgan');

export interface ODataResponse<T> {
  '@odata.context': string,
  '@odata.nextLink'?: string;
  value: T
}

export interface IOneDriveUser {
  displayName: string;
  id: string;
}

export interface IDrive {
  id: string,
  driveType: 'personal'
  owner: IOneDriveUser;
  quota: {
    used: number;
    total: number;
    remaining: number;
    deleted: number;
    state: 'normal';
  }
}

export interface IDriveItem {
  createdBy: {
    user: IOneDriveUser
  };
  createdDateTime: string;
  cTag: string;
  eTag: string;
  fileSystemInfo: {
    createdDateTime: string;
    lastModifiedDateTime: string;
  };
  '@microsoft.graph.downloadUrl'?: string;
  image?: { height: number, width: number }

  file?: {
    mimeType: 'image/jpeg'
  }
  folder?: {
    childCount: number;
  };
  id: string;
  name: string;
  lastModifiedBy: {
    user: IOneDriveUser;
  }
  lastModifiedDateTime: string;
  parentReference: {
    driveId: string;
    id: string;
    path: string;
  };
  size: 0 | number;
  webUrl: string;
}


export class OneDriveAPIDao {

  private static readonly BASE_URL = 'https://graph.microsoft.com';

  private static getHttpConnection() {
    return request
  }

  public static browseByUrl(accessToken: string, url: string) {
    return new Promise<{ items: IDriveItem[], nextLink: string | null | undefined }>((ok, fail) => {
      OneDriveAPIDao.getHttpConnection()
        .get(url)
        .set('Authorization', 'Bearer ' + accessToken)
        .end((error, response) => {
          // Returns 200 OK and the photo in the body. If no photo exists, returns 404 Not Found.
          if (error) {
            fail({ error, response })
          } else {
            const items = (response.body as ODataResponse<IDriveItem[]>).value;


            ok({
              items,
              nextLink: (response.body as ODataResponse<IDriveItem[]>)['@odata.nextLink']
            })
          }
        })
        .on('error', e => fail(e));
    });
  }

  public static browseById(accessToken: string, driveId: string, itemId: string) {
    return this.browseByUrl(accessToken, `${this.BASE_URL}/beta/drives/${driveId}/items/${itemId}/children`);
  }

  public static getDriveItems(accessToken: string, driveId: string) {
    return new Promise<IDriveItem[]>((ok, fail) => {
      OneDriveAPIDao.getHttpConnection()
        .get(`${this.BASE_URL}/beta/drives/${driveId}/root/children`)
        .set('Authorization', 'Bearer ' + accessToken)
        .end((error, response) => {
          // Returns 200 OK and the photo in the body. If no photo exists, returns 404 Not Found.
          if (error) {
            fail({ error, response })
          } else {
            const items = (response.body as ODataResponse<IDriveItem[]>).value
            ok(items)
          }
        })
        .on('error', e => fail(e));
    });
  }

  public static getDrives(accessToken: string) {
    return new Promise<IDrive[]>((ok, fail) => {
      OneDriveAPIDao.getHttpConnection()
        .get(`${this.BASE_URL}/beta/drives`)
        .set('Authorization', 'Bearer ' + accessToken)
        .end((error, response) => {
          // Returns 200 OK and the photo in the body. If no photo exists, returns 404 Not Found.
          if (error) {
            fail({ error, response })
          } else {

            ok((response.body as ODataResponse<IDrive[]>).value)
          }
        })
        .on('error', e => fail(e));
    });
  }

  public static getUserProfile(accessToken: string) {
    return new Promise<{ displayName: string; emails: { address: string; }[]; }>((ok, fail) => {
      request
        .get(`${this.BASE_URL}/beta/me`)
        .set('Authorization', 'Bearer ' + accessToken)
        .end((error, response) => {
          // Returns 200 OK and the photo in the body. If no photo exists, returns 404 Not Found.
          if (error) {
            fail({ error, response })
          } else {
            const profile = {
              displayName: response.body.displayName,
              emails: [{ address: response.body.mail || response.body.userPrincipalName }]
            };

            ok(profile)
          }
        })
        .on('error', e => fail(e));
    });
  }

  public static async getProfilePhoto(accessToken: string) {
    return new Promise<string | Buffer>((ok, fail) => {
      // Get the profile photo of the current user (from the user's mailbox on Exchange Online).
      // This operation in version 1.0 supports only work or school mailboxes, not personal mailboxes.
      OneDriveAPIDao.getHttpConnection()
        .get(`${this.BASE_URL}/beta/me/photo/$value`)
        .set('Authorization', 'Bearer ' + accessToken)
        .end((error, response) => {
          // Returns 200 OK and the photo in the body. If no photo exists, returns 404 Not Found.
          if (error) {
            fail({ error, response })
          } else {
            ok(response.body)
          }
        })
        .on('error', e => fail(e));
    })
  }


  public static uploadFile(accessToken, file: string | Buffer) {
    return new Promise<{ id: string }>((ok, fail) => {
      // This operation only supports files up to 4MB in size.
      // To upload larger files, see `https://developer.microsoft.com/graph/docs/api-reference/v1.0/api/item_createUploadSession`.
      OneDriveAPIDao.getHttpConnection()
        .put(`${this.BASE_URL}/beta/me/drive/root/children/mypic.jpg/content`)
        .send(file)
        .set('Authorization', 'Bearer ' + accessToken)
        .set('Content-Type', 'image/jpg')
        .end((error, response) => {
          // Returns 200 OK and the file metadata in the body.
          if (error) {
            fail({ error, response })
          } else {
            ok(response.body)
          }
        })
        .on('error', e => fail(e));
    });
  }

  // See https://developer.microsoft.com/en-us/graph/docs/api-reference/v1.0/api/item_createlink
  public static getSharingLink(accessToken, id) {
    return new Promise<{ webUrl: string }>((ok, fail) => {
      OneDriveAPIDao.getHttpConnection()
        .post(`${this.BASE_URL}/beta/me/drive/items/${id}/createLink`)
        .send({ type: 'view' })
        .set('Authorization', 'Bearer ' + accessToken)
        .set('Content-Type', 'application/json')
        .end((error, response) => {
          // Returns 200 OK and the file metadata in the body.
          if (error) {
            fail({ error, response })
          } else {
            ok(response.body.link);
          }
        })
        .on('error', e => fail(e));
    });
  }

  /**
   * Per issue #53 for BadRequest when message uses utf-8 characters:
   * `.set('Content-Length': Buffer.byteLength(mailBody,'utf8'))`
   */
  public static postSendMail(accessToken: string, message: string) {
    return new Promise<request.Response>((ok, fail) => {
      OneDriveAPIDao.getHttpConnection()
        .post(`${this.BASE_URL}/beta/me/sendMail`)
        .send(message)
        .set('Authorization', 'Bearer ' + accessToken)
        .set('Content-Type', 'application/json')
        .set('Content-Length', message.length.toString())
        .end((error, response) => {
          // Returns 202 if successful.
          // Note: If you receive a 500 - Internal Server Error
          // while using a Microsoft account (outlook.com, hotmail.com or live.com),
          // it's possible that your account has not been migrated to support this flow.
          // Check the inner error object for code 'ErrorInternalServerTransientError'.
          // You can try using a newly created Microsoft account or contact support.
          if (error) {
            fail({ error, response })
          } else {
            ok(response.body.link);
          }
        })
        .on('error', e => fail(e));
    });
  }
}


