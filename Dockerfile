FROM node:18-alpine

WORKDIR /app

# Copy manifest and install production deps only
COPY package*.json ./
RUN npm ci --only=production

# Copy application source
COPY server.js ./
COPY public ./public

# Create runtime directories (data + temp are mounted or created by the server)
RUN mkdir -p temp

EXPOSE 3000

CMD ["node", "server.js"]
