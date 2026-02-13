const app = require("./app");
const http = require("http");

// Health check server for Docker/Kubernetes
// Runs on a separate path to avoid interfering with Teams Bot endpoints
const HEALTH_PORT = process.env.HEALTH_PORT || 3979;

const healthServer = http.createServer((req, res) => {
  if (req.url === "/api/health" || req.url === "/health") {
    res.writeHead(200, { "Content-Type": "application/json" });
    res.end(JSON.stringify({
      status: "healthy",
      service: "keeper-teams-bot",
      timestamp: new Date().toISOString(),
      uptime: process.uptime(),
    }));
  } else {
    res.writeHead(404, { "Content-Type": "application/json" });
    res.end(JSON.stringify({ error: "Not found" }));
  }
});

// Start the application
(async () => {
  // Start health check server
  healthServer.listen(HEALTH_PORT, () => {
    console.log(`[Health] Health check endpoint available at http://localhost:${HEALTH_PORT}/api/health`);
  });
  
  // Start Teams bot
  await app.start();
  console.log(`\nBot started, app listening to`, process.env.PORT || process.env.port || 3978);
})();
