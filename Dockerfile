# Stage 1: Install dependencies
FROM node:18-alpine AS deps
WORKDIR /usr/src/app
COPY package.json package-lock.json* ./
RUN npm install

# Stage 2: Build the application
FROM node:18-alpine AS builder
WORKDIR /usr/src/app
COPY --from=deps /usr/src/app/node_modules ./node_modules
COPY . .
RUN npm run build

# Stage 3: Production image
FROM node:18-alpine AS final
WORKDIR /usr/src/app
COPY --from=builder /usr/src/app/dist ./dist
COPY --from=builder /usr/src/app/package*.json ./
COPY --from=builder /usr/src/app/server.mjs .

EXPOSE 3001
CMD ["node", "server.mjs"]
