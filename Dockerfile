# =============================================================================
# PMO Project Initiation Portal — Dockerfile
# =============================================================================
# Multi-stage build for minimal production image.
# =============================================================================

# ---- Build Stage ----
FROM node:20-alpine AS builder

WORKDIR /app

COPY package.json package-lock.json ./
RUN npm ci --only=production

COPY . .

# ---- Production Stage ----
FROM node:20-alpine

WORKDIR /app

# Create non-root user
RUN addgroup -g 1001 -S nodejs && \
    adduser -S express -u 1001

COPY --from=builder --chown=express:nodejs /app/node_modules ./node_modules
COPY --from=builder --chown=express:nodejs /app/server.js ./server.js
COPY --from=builder --chown=express:nodejs /app/src ./src
COPY --from=builder --chown=express:nodejs /app/ProjectInitiationForm.html ./ProjectInitiationForm.html

USER express

EXPOSE 3000

# Health check
HEALTHCHECK --interval=30s --timeout=3s --start-period=5s --retries=3 \
  CMD node -e "require('http').get('http://localhost:3000/api/health', (r) => { process.exit(r.statusCode === 200 ? 0 : 1); })"

CMD ["node", "server.js"]
