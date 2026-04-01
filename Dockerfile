FROM node:22-alpine AS builder

WORKDIR /app

RUN corepack enable && corepack prepare pnpm@latest --activate

COPY package.json pnpm-lock.yaml ./
RUN pnpm install --frozen-lockfile

COPY tsconfig.json tsdown.config.ts ts-builds.config.json ./
COPY src/ src/
RUN pnpm build

FROM node:22-alpine

WORKDIR /app

RUN corepack enable && corepack prepare pnpm@latest --activate

COPY package.json pnpm-lock.yaml ./
RUN pnpm install --frozen-lockfile --prod

COPY --from=builder /app/dist/ dist/

ENV NODE_ENV=production
ENV TRANSPORT_TYPE=httpStream
ENV PORT=8080

EXPOSE 8080

CMD ["node", "dist/bin.js"]
