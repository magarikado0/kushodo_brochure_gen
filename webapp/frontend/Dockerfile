FROM node:22-alpine AS build

WORKDIR /app
COPY webapp/frontend/package.json /app/package.json
RUN npm install
COPY webapp/frontend /app
RUN npm run build

FROM nginx:1.27-alpine
COPY --from=build /app/dist /usr/share/nginx/html
COPY webapp/frontend/nginx.conf /etc/nginx/conf.d/default.conf
