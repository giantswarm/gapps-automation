# Browsers, fonts and Xvfb are part of the Playwright image, its version must match playwright-core in package.json
FROM mcr.microsoft.com/playwright:v1.64.1-noble

WORKDIR /app
COPY package.json package-lock.json ./
RUN npm ci --omit=dev && npm cache clean --force
COPY src ./src
RUN mkdir -p /state && chown pwuser /state

ENV NODE_ENV=production \
    MEETBOT_STORAGE_STATE=/secrets/storage-state.json \
    MEETBOT_STATE_DIR=/state

# pwuser is the unprivileged user of the Playwright image
USER pwuser
EXPOSE 8080

# Meet treats headless browsers with suspicion, so run headed on a virtual display
CMD ["xvfb-run", "--auto-servernum", "--server-args=-screen 0 1280x800x24", "node", "src/server.mjs"]
