FROM node:20-bookworm-slim

# LibreOffice does the PDF export; Python + openpyxl applies table styling
# (borders, header row, fit-to-width) before that export happens.
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice \
    fonts-liberation \
    fonts-dejavu \
    python3 \
    python3-openpyxl \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY package*.json ./
RUN npm install --omit=dev

COPY . .

ENV PORT=3000
EXPOSE 3000

CMD ["node", "server.js"]
