const crypto = require('crypto');

module.exports = async (req, res) => {
  console.log('=== Webhook Request Received ===');
  console.log('Method:', req.method);
  console.log('Headers:', JSON.stringify(req.headers, null, 2));
  console.log('Query:', JSON.stringify(req.query, null, 2));
  console.log('Body:', JSON.stringify(req.body, null, 2));

  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type, X-EBAY-SIGNATURE');

  if (req.method === 'OPTIONS') {
    res.status(200).end();
    return;
  }

  if (req.method === 'GET') {
    const challengeCode = req.query.challengeCode || req.query.challenge_code;
    const verificationToken = req.query.verificationToken || req.query.verification_token;
    const endpointUrl = req.query.endpointUrl || req.query.endpoint_url;

    console.log('GET - Challenge Code:', challengeCode);
    console.log('GET - Verification Token:', verificationToken);
    console.log('GET - Endpoint URL:', endpointUrl);

    if (!challengeCode) {
      console.log('Challenge code is missing. Responding with an empty 200 OK.');
      res.status(200).end();
      return;
    }

    if (!verificationToken || !endpointUrl) {
      console.log('Missing verificationToken or endpointUrl.');
      res.status(400).json({ error: 'Missing verificationToken or endpointUrl' });
      return;
    }

    try {
      const stringToHash = challengeCode + verificationToken + endpointUrl;
      const hash = crypto.createHash('sha256').update(stringToHash).digest('hex');
      console.log('GET - Calculated hash:', hash);
      res.setHeader('Content-Type', 'application/json');
      res.status(200).json({ challengeResponse: hash });
    } catch (error) {
      console.error('Error processing GET request:', error);
      res.status(500).json({ error: 'Internal server error' });
    }
    return;
  }

  if (req.method === 'POST') {
    console.log('=== Processing POST request ===');
    
    const challengeCode = req.body.challengeCode || req.body.challenge_code;
    const verificationToken = process.env.EBAY_VERIFICATION_TOKEN;
    const endpointUrl = process.env.EBAY_ENDPOINT_URL;

    console.log('POST - Challenge Code:', challengeCode);
    console.log('POST - Verification Token (from env):', verificationToken ? '***' : 'NOT SET');
    console.log('POST - Endpoint URL (from env):', endpointUrl);

    if (challengeCode) {
      if (!verificationToken || !endpointUrl) {
        console.log('Missing environment variables: EBAY_VERIFICATION_TOKEN or EBAY_ENDPOINT_URL');
        res.status(400).json({ error: 'Missing required environment variables' });
        return;
      }

      try {
        const stringToHash = challengeCode + verificationToken + endpointUrl;
        const hash = crypto.createHash('sha256').update(stringToHash).digest('hex');
        console.log('POST - Calculated hash:', hash);
        res.setHeader('Content-Type', 'application/json');
        res.status(200).json({ challengeResponse: hash });
      } catch (error) {
        console.error('Error processing POST verification request:', error);
        res.status(500).json({ error: 'Internal server error' });
      }
      return;
    }

    console.log('POST - Processing notification request');
    res.status(200).json({ message: 'Notification received successfully' });
    return;
  }

  res.status(405).json({ error: 'Method not allowed' });
}; 