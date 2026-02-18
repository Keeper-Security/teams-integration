const app = require("./app");
const http = require("http");
const { createLogger } = require("./services");

const log = createLogger('Server');
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

(async () => {
  healthServer.listen(HEALTH_PORT, () => {
    log.info(`Health check endpoint available at http://localhost:${HEALTH_PORT}/api/health`);
  });
  
  await app.start();
  log.info(`Bot started, app listening on port ${process.env.PORT || process.env.port || 3978}`);
})();
