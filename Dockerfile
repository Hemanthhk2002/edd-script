FROM node:18-alpine

WORKDIR /app

# Install dependencies first for better caching
COPY package*.json ./

# Install production dependencies only
RUN npm install --production

# Copy application files
COPY . .

# Expose the port the app runs on (default Express port is 3000)
EXPOSE 4000

# Set environment variables
ENV NODE_ENV=production
ENV PORT=4000

# Command to run the application
CMD ["node", "index.js"]
