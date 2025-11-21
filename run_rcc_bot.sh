#!/bin/bash
# Helper script to run the Robocorp bot with LibreOffice in Docker

set -e

echo "=========================================="
echo "Robocorp LibreOffice Bot Runner"
echo "=========================================="

# Check if Docker is running
if ! docker info > /dev/null 2>&1; then
    echo "❌ Error: Docker is not running"
    exit 1
fi

# Build and run the bot
echo ""
echo "🔨 Building Docker image..."
docker-compose -f docker-compose.rcc.yml build

echo ""
echo "🚀 Running Robocorp bot..."
echo "(Showing last 50 lines of output...)"
echo ""
docker-compose -f docker-compose.rcc.yml up 2>&1 | tail -n 50

echo ""
echo "✅ Bot execution completed!"
echo ""
echo "📁 Check the 'output' directory for results"
