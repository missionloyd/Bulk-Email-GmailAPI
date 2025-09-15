#!/bin/bash
touch last_sent.txt

echo "Killing old Docker processes"
docker compose rm -fs

echo "Spinning up Docker containers"
docker compose build --force-rm && \
docker compose up --detach && \
docker compose logs --follow
