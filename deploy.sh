#!/bin/bash
set -e

PROJECT="project-d5a525d1-72ba-431c-80c"
SERVICE="byd-search"
REGION="europe-west1"

GCLOUD="${CLOUDSDK_PYTHON:+CLOUDSDK_PYTHON=$CLOUDSDK_PYTHON} gcloud"

echo "=== BYD Search — Cloud Run Deployment ==="
echo "Project: $PROJECT"
echo "Service: $SERVICE"
echo "Region:  $REGION"
echo

if ! gcloud auth list 2>/dev/null | grep -q ACTIVE; then
  echo "Not logged in. Opening browser for authentication..."
  gcloud auth login --project "$PROJECT"
fi

echo
echo "Building and deploying to Cloud Run..."
gcloud run deploy "$SERVICE" \
  --source . \
  --region "$REGION" \
  --project "$PROJECT" \
  --memory 1Gi \
  --timeout 300 \
  --allow-unauthenticated

echo
echo "Done! Service URL:"
gcloud run services describe "$SERVICE" \
  --region "$REGION" \
  --project "$PROJECT" \
  --format "value(status.url)"
