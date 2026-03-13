#!/bin/bash

# Simple git sync script
# Pulls changes, then adds and pushes local changes

REMOTE="origin"
BRANCH="main"

echo "Starting automated sync for $REMOTE/$BRANCH..."

while true; do
  # Pull changes
  git pull $REMOTE $BRANCH --rebase
  
  # Check if there are changes to commit
  if [[ -n $(git status -s) ]]; then
    echo "Changes detected, syncing..."
    git add .
    git commit -m "auto: sync changes"
    git push $REMOTE $BRANCH
    echo "Sync complete."
  fi
  
  sleep 30
done
