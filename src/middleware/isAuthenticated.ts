
import express = require('express');
import { Render } from '../render/Render';
import { ExpressRequest, ExpressResponse } from '../models/ExpressRequest';

export function isAuthenticated(req: ExpressRequest, res: ExpressResponse, next: Function) {
  if (req.isAuthenticated()) {
    next();
    return;
  }

  Render.error(new Error('401 Authentication Error'), res);
}