const crypto = require("crypto");

module.exports = async (req, res) => {
  console.log("=== Webhook Request Received ===");
  console.log("Method:", req.method);
  console.log("Headers:", JSON.stringify(req.headers, null, 2));
  console.log("Body:", JSON.stringify(req.body, null, 2));
  console.log("Query:", JSON.stringify(req.query, null, 2));

  // CORS headers
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type, X-EBAY-SIGNATURE");

  // Handle OPTIONS request (preflight)
  if (req.method === "OPTIONS") {
    res.status(200).end();
    return;
  }

  // Handle GET request (eBay validation)
  if (req.method === "GET") {
    console.log("=== Processing GET request (eBay validation) ===");
    
    const { challengeCode, verificationToken, endpointUrl } = req.query;
    
    console.log("Challenge Code:", challengeCode);
    console.log("Verification Token:", verificationToken);
    console.log("Endpoint URL:", endpointUrl);
    
    if (!challengeCode || !verificationToken || !endpointUrl) {
      console.log("Missing required parameters");
      res.status(400).json({ error: "Missing required parameters" });
      return;
    }

    try {
      // Create the string to hash
      const stringToHash = challengeCode + verificationToken + endpointUrl;
      console.log("String to hash:", stringToHash);
      
      // Calculate SHA256 hash
      const hash = crypto.createHash("sha256").update(stringToHash).digest("hex");
      console.log("Calculated hash:", hash);
      
      // Return the hash as plain text
      res.setHeader("Content-Type", "text/plain");
      res.status(200).send(hash);
      
      console.log("=== GET request completed successfully ===");
    } catch (error) {
      console.error("Error processing GET request:", error);
      res.status(500).json({ error: "Internal server error" });
    }
    return;
  }

  // Handle POST request (eBay notifications)
  if (req.method === "POST") {
    console.log("=== Processing POST request (eBay notification) ===");
    
    try {
      // Log the notification data
      console.log("Notification data:", JSON.stringify(req.body, null, 2));
      
      // Here you would typically process the notification
      // For now, we will just acknowledge receipt
      
      res.status(200).json({ 
        status: "success",
        message: "Notification received successfully",
        timestamp: new Date().toISOString()
      });
      
      console.log("=== POST request completed successfully ===");
    } catch (error) {
      console.error("Error processing POST request:", error);
      res.status(500).json({ error: "Internal server error" });
    }
    return;
  }

  // Handle unsupported methods
  res.status(405).json({ error: "Method not allowed" });
};
