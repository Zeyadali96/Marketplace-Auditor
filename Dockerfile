FROM node:20-bookworm

WORKDIR /app

COPY package*.json ./
RUN npm install

RUN npx playwright install --with-deps chromium

COPY . .

ENV NODE_ENV=production
RUN npm run build

CMD ["npm", "start"]
