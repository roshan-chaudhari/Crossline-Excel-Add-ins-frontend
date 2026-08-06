# Build the frontend assets for the Excel Add-in
FROM node:20-alpine
WORKDIR /app

# Copy package files and install dependencies
COPY package*.json ./
RUN npm install

# Copy source code
COPY . .

# Build argument for the client URL
ARG CLIENT_SERVER_URL=""
ENV CLIENT_SERVER_URL=$CLIENT_SERVER_URL

# Build client files to /app/dist
RUN npm run build

# By default, copy the build files to the volume mount point
CMD ["sh", "-c", "cp -R /app/dist/* /app/output/"]
