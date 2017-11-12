import express = require('express');

export type UserProfile = { displayName: string, emails: { address: string }[] };

export type ExpressRequest = express.Request
  & Partial<{ user: { accessToken: string } }>

export type ExpressResponse = express.Response
  & { render: <T>(view: 'login' | 'sendMail', vm?: T) => void };
