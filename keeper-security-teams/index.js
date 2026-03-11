const http = require("http");

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
  // Initialize configuration from KSM (if available) before loading app
  const { initializeConfig, getConfig } = require("./config");
  await initializeConfig();
  const config = getConfig();
  
  // Set environment variables from KSM config for Teams SDK
  // The Teams SDK reads credentials from process.env, not from config object
  if (config.MicrosoftAppId) process.env.CLIENT_ID = config.MicrosoftAppId;
  if (config.MicrosoftAppPassword) process.env.CLIENT_SECRET = config.MicrosoftAppPassword;
  if (config.MicrosoftAppTenantId) process.env.TENANT_ID = config.MicrosoftAppTenantId;
  if (config.MicrosoftAppType) process.env.BOT_TYPE = config.MicrosoftAppType;
  
  // Now load the app (config is already initialized)
  const app = require("./app");
  const { createLogger } = require("./services");
  const log = createLogger('Server');
  
  healthServer.listen(HEALTH_PORT);
  
  await app.start();
  log.info(`Bot started, app listening on port ${process.env.PORT || process.env.port || 3978}`);
  
  // Log KSM status
  if (config.ksm?.enabled) {
    log.info(`Configuration loaded from KSM: ${config.ksm.loadedSections.join(', ')}`);
  }
})();
