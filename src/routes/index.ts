import passport = require('passport');
import express = require('express');

import { ExpressRequest, ExpressResponse, UserProfile } from '../models/ExpressRequest';
import { Render } from '../render/Render';
import { isAuthenticated } from '../middleware/isAuthenticated';
import { OneDriveAPIDao, IDriveItem } from '../dal/OneDriveAPIDao';
import { EmailService } from '../services/EmailService';

export type SendMailFormBody = { body?: Partial<{ default_email: string }> };

function isPromise(it: any): it is Promise<void> {
  if (it !== null && typeof it === 'object') {
    return it && typeof (it as Promise<void>).then === 'function';
  }
  return false;
}


// helpers
function hasAccessTokenExpired(e: string | Error | any) {
  if (typeof e === 'string') {
    return false;
  }

  let expired: boolean;

  if (!e.innerError) {
    expired = false;
  } else {
    expired = e.forbidden &&
      e.message === 'InvalidAuthenticationToken' &&
      e.response && e.response.error.message === 'Access token has expired.';
  }
  return expired;
}

const router = express.Router();

export function _register(verb: 'GET' | 'POST', path: string, bofore?: any) {

  console.log('register', verb, path);

  return (target, propertyKey: string, descriptor: PropertyDescriptor) => {

    const ctrlAction: (...params) => void | Promise<void> = descriptor.value;

    const middleWare = !bofore
      ? (req, res, next) => next()
      : bofore


    router[verb.toLowerCase()](path, middleWare, async (req: ExpressRequest, res: ExpressResponse) => {

      try {

        await ctrlAction(req, res);

      } catch (ex) {

        if (hasAccessTokenExpired(ex)) {
          ex.message += ' Expired token. Please sign out and sign in again.';
        }

        Render.error(ex, res);

      }

    });
  }
}


class Http {
  public static get = (path: string, ...args) => _register('GET', path, ...args)
  public static post = (path: string, ...args) => _register('POST', path, ...args)
}

const openIdConnect = passport.authenticate('azuread-openidconnect', { failureRedirect: '/' });


export class Router {

  @Http.get('/')
  public static async index(req: ExpressRequest, res: ExpressResponse) {
    if (!req.isAuthenticated()) {
      res.redirect('login-page');
    } else {
      res.redirect('sendMail');
    }
  }

  @Http.get('/login-page')
  public static async lodingPage(req: ExpressRequest, res: ExpressResponse) {
    if (req.isAuthenticated()) {
      res.redirect('/')
    } else {
      res.render('login');
    }
  }

  @Http.get('/login', openIdConnect)
  public static async login(req: ExpressRequest, res: ExpressResponse) {
    res.redirect('/');
  }

  @Http.get('/token', openIdConnect)
  public static async token(req: ExpressRequest, res: ExpressResponse) {
    res.redirect('/');
  }

  @Http.get('/disconnect', isAuthenticated)
  public static async disconnect(req: ExpressRequest, res: ExpressResponse) {
    if (!req.session) {
      throw new Error('Session missing');
    }
    req.session.destroy(() => {
      req.logOut();
      res.clearCookie('graphNodeCookie');
      res.status(200);
      res.redirect('/');
    });
  }

  @Http.get('/sendMail', isAuthenticated)
  public static async sendMailGET(req: ExpressRequest & SendMailFormBody, res: ExpressResponse) {
    const profile = await OneDriveAPIDao.getUserProfile(req.user.accessToken);
    Render.sendMail(profile, undefined, res);
  }

  @Http.post('/sendMail', isAuthenticated)
  public static async sendMailPOST(req: ExpressRequest & SendMailFormBody, res: ExpressResponse) {

    const toEmail = req.body.default_email;

    const profile = await OneDriveAPIDao.getUserProfile(req.user.accessToken);
    const mailBody = await EmailService.prepForEmailMessage(req.user.accessToken, profile.displayName, toEmail);
    await OneDriveAPIDao.postSendMail(req.user.accessToken, JSON.stringify(mailBody));

    Render.sendMail(profile, toEmail, res);
  }


  @Http.get('/files', isAuthenticated)
  public static async files(req: ExpressRequest & SendMailFormBody, res: ExpressResponse) {
    const drives = await OneDriveAPIDao.getDrives(req.user.accessToken);
    const driveId = drives[0].id;
    const items = await OneDriveAPIDao.getDriveItems(req.user.accessToken, driveId);
    Render.folder({ driveId, items }, res);
  }

  @Http.get('/browse/:driveId/:itemId', isAuthenticated)
  public static async browse(req: ExpressRequest, res: ExpressResponse) {

    const firstPage = await OneDriveAPIDao.browseById(
      req.user.accessToken,
      req.params.driveId,
      req.params.itemId
    );

    const getNext = async (url: string, pushInto: IDriveItem[]): Promise<IDriveItem[]> => {

      const { items, nextLink } = await OneDriveAPIDao.browseByUrl(req.user.accessToken, url);
      pushInto = [...pushInto, ...items];

      return nextLink
        ? getNext(nextLink, pushInto)
        : pushInto;
    };

    const items = firstPage.nextLink
      ? await getNext(firstPage.nextLink, firstPage.items)
      : firstPage.items;


    Render.folder({ driveId: req.params.driveId, items }, res);
  }

  public static register() {
    return router;
  }
}
