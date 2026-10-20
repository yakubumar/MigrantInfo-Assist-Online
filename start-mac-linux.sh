#!/usr/bin/env bash
cd "$(dirname "$0")"
if ! command -v node >/dev/null 2>&1; then
  echo "Node.js topilmadi. https://nodejs.org dan LTS versiyani o'rnating."; exit 1
fi
[ -d node_modules ] || { echo "Kutubxonalar o'rnatilmoqda..."; npm install --no-audit --no-fund || exit 1; }
if [ ! -f .env ]; then
  read -r -p "Gemini API kalitini kiriting: " KEY
  echo "GEMINI_API_KEY=$KEY" > .env
fi
[ -f dist/server.cjs ] || { echo "Dastur yig'ilmoqda..."; npm run build || exit 1; }
export NODE_ENV=production
( sleep 3; (open http://localhost:3000 || xdg-open http://localhost:3000) >/dev/null 2>&1 ) &
node dist/server.cjs
