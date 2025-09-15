#!/bin/bash

# VARO REBILLING - Start Script
# This starts both the frontend (live-server) and backend (FastAPI) servers

echo "🛢️  Starting VARO REBILLING Application..."
echo ""

# Start Python backend in background
echo "🐍 Starting Python backend (FastAPI) on port 8000..."
python3 server.py &
PYTHON_PID=$!

# Wait a moment for Python server to start
sleep 2

# Start frontend server in background
echo "🌐 Starting frontend server (live-server) on port 8080..."
npm start &
FRONTEND_PID=$!

echo ""
echo "✅ Both servers are starting..."
echo ""
echo "📊 Dashboard:     http://127.0.0.1:8080"
echo "🔧 API Backend:   http://127.0.0.1:8000"
echo "📖 API Docs:      http://127.0.0.1:8000/docs"
echo ""
echo "Press Ctrl+C to stop both servers"

# Function to cleanup processes
cleanup() {
    echo ""
    echo "🛑 Stopping servers..."
    kill $PYTHON_PID 2>/dev/null
    kill $FRONTEND_PID 2>/dev/null
    echo "✅ Servers stopped"
    exit 0
}

# Set trap to cleanup on exit
trap cleanup SIGINT SIGTERM

# Wait for background processes
wait