#!/usr/bin/env bash
set -e

echo "============================================"
echo "Compiling Student Performance Analytics..."
echo "============================================"
javac StudentPerformanceAnalytics.java

echo "Compilation successful."
echo ""
echo "============================================"
echo "Running Student Performance Analytics Tool..."
echo "============================================"
java StudentPerformanceAnalytics
