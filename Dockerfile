FROM node:20-bookworm-slim

WORKDIR /usr/src/app


RUN apt-get update && apt-get install -y git && rm -rf /var/lib/apt/lists/*
COPY package*.json ./
RUN node -v
RUN npm -v
RUN npm i
RUN git config --global --add safe.directory /usr/src/app
