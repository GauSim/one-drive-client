/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
// application dependencies
import express = require('express');
import session = require('express-session');
import path = require('path');
import favicon = require('serve-favicon');
import logger = require('morgan');
import cookieParser = require('cookie-parser');
import bodyParser = require('body-parser');
import passport = require('passport');
const OIDCStrategy = (require('passport-azure-ad') as any).OIDCStrategy;
import uuid = require('uuid');
import { ConfigService } from './services/ConfigService';
import { Router } from './routes/index';

const app = express();

// **IMPORTANT
// Note that production apps will need to create a self-signed cert and use a secure server,
// and change dev settings marked 'For development only' in app.js and config.js.
// Below is an example after you have the key cert pair:
// const https = require('https');
// const certConfig = {
//  key: fs.readFileSync('./utils/cert/server.key', 'utf8'),
//  cert: fs.readFileSync('./utils/cert/server.crt', 'utf8')
// };
// const server = https.createServer(certConfig, app);

// authentication setup
const callback = (iss, sub, profile, accessToken, refreshToken, done) => {
  done(null, { profile, accessToken, refreshToken });
};

type User = {};
const users: { [userId: string]: User } = {};
const config = new ConfigService();

passport.use(new OIDCStrategy(config.getApiCreds(), callback));
passport.serializeUser<User, string>((user, done) => {
  const id = uuid.v4();
  users[id] = user;
  done(null, id);
});
passport.deserializeUser<User, string>((id, done) => {
  const user = users[id];
  done(null, user);
});

// view engine setup
app.set('views', path.join(__dirname + '/../', 'views'));
app.set('view engine', 'hbs');

app.use(favicon(path.join(__dirname + '/../', 'public', 'img', 'favicon.ico')));
app.use(logger('dev'));
app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: false }));
app.use(cookieParser());
app.use(session({
  secret: '12345QWERTY-SECRET',
  name: 'graphNodeCookie',
  resave: false,
  saveUninitialized: false,
  //cookie: {secure: true} // For development only
}));
app.use(express.static(path.join(__dirname + '/../', 'public')));
app.use(passport.initialize());
app.use(passport.session());
app.use('/', Router.register());

// error handlers
// catch 404 and forward to error handler
app.use(function (req, res, next) {
  const err = new Error('Not Found');
  res.status(404);
  next(err);
});

// development error handler
// will print stacktrace
if (app.get('env') === 'development') {
  app.use((err: Error & Partial<{ status: number }>, req: express.Request, res: express.Response, _) => {
    res.status(err.status || 500);
    res.render('error', {
      message: err.message,
      error: err
    });
  });
}

// production error handler
// no stacktraces leaked to user
app.use((err: Error & Partial<{ status: number }>, req: express.Request, res: express.Response, _) => {
  res.status(err.status || 500);
  res.render('error', {
    message: err.message,
    error: {}
  });
});

export {
  app
}
