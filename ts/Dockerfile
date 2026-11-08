# Remote MCP for mail, calendars and Teams (src/http.ts). Node 26 runs the TypeScript sources as they are: no build step.
# Same base image digest as OE Copilot's agente-ts (index amd64 + arm64); to update: docker buildx imagetools inspect node:26-slim
FROM node:26-slim@sha256:193fe51b64e77981119c98c2002c9e32a70e2f006fb4d25068ce0558998917f0
ENV NODE_ENV=production BIND_HOST=0.0.0.0 PORT=8190
WORKDIR /app
COPY package.json package-lock.json ./
RUN npm ci --omit=dev --ignore-scripts --no-audit --no-fund && npm cache clean --force
COPY src ./src
USER node
EXPOSE 8190
HEALTHCHECK --interval=30s --timeout=5s --start-period=30s --retries=3 \
  CMD ["node", "-e", "fetch('http://127.0.0.1:8190/healthz').then((r) => process.exit(r.ok ? 0 : 1), () => process.exit(1))"]
CMD ["node", "src/http.ts"]
