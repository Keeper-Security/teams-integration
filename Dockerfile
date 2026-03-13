# Keeper Teams App Dockerfile
FROM node:20-slim

# Set working directory
WORKDIR /app

# Set environment variables
ENV NODE_ENV=production

# Install curl for health checks
RUN apt-get update && apt-get install -y --no-install-recommends \
    curl \
    && rm -rf /var/lib/apt/lists/*

# Copy package files first (for better caching)
COPY package*.json ./

# Install dependencies (production only)
RUN npm ci --only=production

# Copy application code
COPY . .

# Create data directory for persistence (conversation references, etc.)
RUN mkdir -p /app/data

# Expose ports (3978 for app, 3979 for health check)
EXPOSE 3978 3979

# Health check (uses dedicated health check port)
HEALTHCHECK --interval=30s --timeout=10s --start-period=15s --retries=3 \
  CMD curl -f http://localhost:3979/api/health || exit 1

# Run the application
CMD ["npm", "start"]
