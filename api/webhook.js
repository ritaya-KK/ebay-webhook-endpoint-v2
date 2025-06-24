const { createHmac } = require('crypto');
const getRawBody = require('raw-body');

const config = {
  api: { bodyParser: false },       // rawBody を読むため
  runtime: 'nodejs'                // Edge を回避
};

console.log('[BUILD TIME] TOKEN:', process.env.EBAY_VERIFICATION_TOKEN ?? 'undefined');

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

async function handler(req, res) {
  console.log('[RUN TIME] TOKEN:', process.env.EBAY_VERIFICATION_TOKEN ?? 'undefined');

  if (req.method === 'GET') {
    const challengeCode = req.query.challengeCode || req.query.challenge_code;
    const verificationToken = process.env.EBAY_VERIFICATION_TOKEN;
    const endpointUrl = process.env.EBAY_ENDPOINT_URL;

    if (!challengeCode || !verificationToken || !endpointUrl) {
      res.status(400).json({ error: 'Missing required parameters' });
      return;
    }

    const stringToHash = challengeCode + verificationToken + endpointUrl;
    const hash = createHmac('sha256', verificationToken)
                     .update(stringToHash)
                     .digest('hex');
    res.status(200).json({ challengeResponse: hash });
    return;
  }

  if (req.method === 'POST') {
    // 1. 生バイト列を取得
    const raw = await getRawBody(req);
    const signature = req.headers['x-ebay-signature'];

    // 2. HMAC を計算
    const expected = createHmac('sha256', process.env.EBAY_VERIFICATION_TOKEN)
                     .update(raw)
                     .digest('base64');

    if (signature !== expected) {
      console.warn('Signature mismatch');
      return res.status(400).end('Invalid signature');
    }

    // 3. ここで JSON.parse(raw) して業務ロジックへ
    const payload = JSON.parse(raw.toString());
    // 例: payload.notification.data.username など
    // 非同期処理に回す場合はここでキューイング等

    // 4. eBay 推奨の 204 空応答
    return res.status(204).end();
  }

  res.setHeader('Allow', ['GET', 'POST']);
  res.status(405).end();
}

module.exports = handler;
module.exports.config = config; 