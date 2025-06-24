const { createHmac } = require('crypto');
const getRawBody = require('raw-body');

const config = {
  api: { bodyParser: false },
  runtime: 'nodejs'
};

async function handler(req, res) {
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
    const raw = await getRawBody(req);
    const signature = req.headers['x-ebay-signature'];

    const expected = createHmac('sha256', process.env.EBAY_VERIFICATION_TOKEN)
                     .update(raw)
                     .digest('base64');

    if (signature !== expected) {
      return res.status(400).end('Invalid signature');
    }

    // 必要ならここでpayloadを非同期処理へ
    // const payload = JSON.parse(raw.toString());

    return res.status(204).end();
  }

  res.setHeader('Allow', ['GET', 'POST']);
  res.status(405).end();
}

module.exports = handler;
module.exports.config = config;