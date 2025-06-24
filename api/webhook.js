export const config = { runtime: 'nodejs' };

console.log('[BUILD TIME] TOKEN:', process.env.EBAY_VERIFICATION_TOKEN ?? 'undefined');

const crypto = require('crypto');

// 生のリクエストボディをパースする関数
async function getRawBody(req) {
  return new Promise((resolve, reject) => {
    let data = '';
    req.on('data', chunk => {
      data += chunk;
    });
    req.on('end', () => {
      resolve(data);
    });
    req.on('error', err => {
      reject(err);
    });
  });
}

export default async function handler(req, res) {
  console.log('[RUN TIME] TOKEN:', process.env.EBAY_VERIFICATION_TOKEN ?? 'undefined');

  let body = req.body;
  if (req.method === 'POST' && !body) {
    // JSONボディを手動でパース
    const rawBody = await getRawBody(req);
    try {
      body = JSON.parse(rawBody);
    } catch (e) {
      body = {};
    }
  }

  console.log('[RUN TIME] BODY:', JSON.stringify(body));

  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type, X-EBAY-SIGNATURE');

  if (req.method === 'OPTIONS') {
    res.status(200).end();
    return;
  }

  if (req.method === 'GET') {
    const challengeCode = req.query.challengeCode || req.query.challenge_code;
    const verificationToken = process.env.EBAY_VERIFICATION_TOKEN;
    const endpointUrl = process.env.EBAY_ENDPOINT_URL;

    if (!challengeCode || !verificationToken || !endpointUrl) {
      res.status(400).json({ error: 'Missing required parameters' });
      return;
    }

    const stringToHash = challengeCode + verificationToken + endpointUrl;
    const hash = crypto.createHash('sha256').update(stringToHash).digest('hex');
    res.status(200).json({ challengeResponse: hash });
    return;
  }

  if (req.method === 'POST') {
    const challengeCode = body.challengeCode || body.challenge_code;
    const verificationToken = process.env.EBAY_VERIFICATION_TOKEN;
    const endpointUrl = process.env.EBAY_ENDPOINT_URL;

    if (!challengeCode || !verificationToken || !endpointUrl) {
      res.status(400).json({ error: 'Missing required parameters' });
      return;
    }

    const stringToHash = challengeCode + verificationToken + endpointUrl;
    const hash = crypto.createHash('sha256').update(stringToHash).digest('hex');
    res.status(200).json({ challengeResponse: hash });
    return;
  }

  res.status(405).json({ error: 'Method not allowed' });
} 