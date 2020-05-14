var express = require('express');
var jsonwebtoken = require('jsonwebtoken');
var fetch = require('node-fetch');
var form = require('form-urlencoded').default;
var arrayBuffer = require('base64-arraybuffer')

const https = require('https');
var router = express.Router();

// Auth URL: https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=b0d311d0-8b1a-4c03-ace7-07b5438952d7&response_type=code&redirect_uri=https://localhost:3001&response_mode=fragment&scope=Files.ReadWrite

/* GET home page. */
router.get('/downloadpicture', async function(req, res, next) {
    const url = req.get('PictureUrl');
    var response = await fetch(url);
    var buffer = await response.arrayBuffer();
    var base64 = arrayBuffer.encode(buffer);
    res.send(base64);
});


router.get('/', async function(req, res, next) {
    const authorization = req.get('Authorization');
    if (authorization == null) {
        throw new Error('No Authorization header was found.');
    }
    const [schema, jwt] = authorization.split(' ');
  
    const decoded = jsonwebtoken.decode(jwt, { complete: true });
    const v2Params = {
      client_id: process.env.CLIENT_ID,
      client_secret: process.env.CLIENT_SECRET,
      grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer',
      assertion: jwt,
      requested_token_use: 'on_behalf_of',
      scope: ['Files.ReadWrite'].join(' ')
    };
  
    const stsDomain = 'https://login.microsoftonline.com';
    const tenant = 'common';
    const tokenURLSegment = 'oauth2/v2.0/token';
  
    const tokenResponse = await fetch(`${stsDomain}/${tenant}/${tokenURLSegment}`, {
      method: 'POST',
      body: form(v2Params),
      headers: {
          'Accept': 'application/json',
          'Content-Type': 'application/x-www-form-urlencoded'
      }
    });

    const token = (await tokenResponse.json()).access_token;


    const pictureFilesResponse = await fetch(`https://graph.microsoft.com/v1.0/me/drive/root:/images:/children`, {
        method: 'GET',
        headers: {
            'Accept': 'application/json',
            'Content-Type': 'application/x-www-form-urlencoded',
            'Authorization': 'Bearer ' + token
        }
      });
    const json = await pictureFilesResponse.json();
    const response = json.value.map((item) => {
        return {downloadUrl: item["@microsoft.graph.downloadUrl"], name: item["name"]};
    });

    res.send(response);
});

module.exports = router;
