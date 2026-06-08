FROM node:20-bookworm-slim

RUN apt-get update \
  && apt-get install -y --no-install-recommends python3 make g++ \
  && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY package.json ./
COPY backend/package.json backend/
COPY frontend/package.json frontend/

RUN npm run install:all

COPY backend backend/
COPY frontend frontend/

RUN npm run build

ENV NODE_ENV=production
ENV PORT=3001
ENV DATA_DIR=/data

RUN mkdir -p /data

EXPOSE 3001

CMD ["npm", "start"]
