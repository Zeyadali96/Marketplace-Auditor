FROM mcr.microsoft.com/playwright:v1.44.0-jammy
WORKDIR /app
COPY package*.json ./
RUN npm install
COPY . .
RUN npm run build
# Railway dynamically assigns a port; we must be ready for it
ENV NODE_ENV=production
EXPOSE 8080
CMD ["npm", "start"]
