# Use the lightweight Playwright-ready image
FROM mcr.microsoft.com/playwright:v1.41.0-jammy
WORKDIR /app
COPY package*.json ./
# Install only production deps to save memory
RUN npm install --only=production
COPY . .
RUN npm run build
# Railway requires the app to listen on 0.0.0.0
ENV PORT=8080
EXPOSE 8080
CMD ["npm", "start"]
