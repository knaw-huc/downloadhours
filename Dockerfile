FROM node:13.10

RUN mkdir /app

WORKDIR /app

COPY package.json /app
COPY yarn.lock /app

RUN yarn install

COPY index.js /app

ENTRYPOINT ["node", "index.js"]