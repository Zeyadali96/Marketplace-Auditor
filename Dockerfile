FROM node:20-bookworm

WORKDIR /app

COPY package*.json ./
RUN npm install

RUN npx playwright install --with-deps chromium

COPY . .

RUN npm run build

CMD ["npm", "start"]
