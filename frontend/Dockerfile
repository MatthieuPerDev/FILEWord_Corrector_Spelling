# Linux léger + Node.js
FROM node:22-alpine

WORKDIR /app

# la liste des librairies d'abord (pour le cache)
COPY package.json package-lock.json ./
RUN npm install # run + commande

# le reste du code React
COPY . .

EXPOSE 5173

# "--host" : accepter les connexions venant de l'extérieur du conteneur
CMD ["npm", "run", "dev", "--", "--host"]