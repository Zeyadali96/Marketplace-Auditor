FROM mcr.microsoft.com/playwright:v1.49.1-jammy
WORKDIR /app
COPY package*.json ./
RUN npm install
COPY . .
RUN npm run build
# Railway dynamically assigns a port; we must be ready for it
ENV NODE_ENV=production
ENV IS_RAILWAY=true
EXPOSE 3000
CMD ["npm", "start"]
