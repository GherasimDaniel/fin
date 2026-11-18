# Serves the month delta as a static page: docker build -t month-delta . && docker run --rm -p 8080:80 month-delta
# The page is rebuilt from data/month.js at build time, so the image always has the latest downloaded data.
# To get new trading days first, run `node scripts/update.js` (or open-month-delta.cmd), then rebuild the image.

# 1) build: put data/month.js into the page template (no network needed)
FROM node:22-alpine AS build
WORKDIR /app
COPY scripts/lib/page.js scripts/lib/page.js
COPY scripts/template.html scripts/template.html
COPY data/month.js data/month.js
RUN node -e "require('./scripts/lib/page').write()"

# 2) serve: nginx with the page as the site index
FROM nginx:1.27-alpine
COPY --from=build /app/month-delta.html /usr/share/nginx/html/index.html
COPY --from=build /app/month-delta.html /usr/share/nginx/html/month-delta.html
EXPOSE 80
