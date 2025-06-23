const crypto = require('crypto');

module.exports = async (req, res) => {
  console.log('=== Webhook Request Received ===');
  console.log('Method:', req.method);
  console.log('Query:', JSON.stringify(req.query, null, 2));

  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type, X-EBAY-SIGNATURE');

  if (req.method === 'OPTIONS') {
    res.status(200).end();
    return;
  }

  if (req.method === 'GET') {
    const { challengeCode, verificationToken, endpointUrl } = req.query;
    
    // eBay sends a "ping" without a challenge code.
    // Respond with an empty 200 OK to acknowledge.
    if (!challengeCode) {
      console.log('Challenge code is missing. Responding with an empty 200 OK.');
      res.status(200).end(); // Send an empty response
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
      console.log('Calculated hash:', hash);
      res.setHeader('Content-Type', 'text/plain');
      res.status(200).send(hash);
    } catch (error) {
      console.error('Error processing GET request:', error);
      res.status(500).json({ error: 'Internal server error' });
    }
    return;
  }

  if (req.method === 'POST') {
    console.log('=== Processing POST request (eBay notification) ===');
    res.status(200).json({ message: 'Notification received.' });
    return;
  }

  res.status(405).json({ error: 'Method not allowed' });
};