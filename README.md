# ENALYSIS: Gold (XAU/USD) ML Engine

A high-performance, zero-dependency machine learning system for real-time XAU/USD price forecasting in the browser.

## Overview

ENALYSIS is a B.Tech Capstone Thesis project that implements a Tri-Factor Ensemble combining:
- **Random Forest** classifier for supervised learning
- **Kalman Filter** for Bayesian state estimation
- **Q-Learning** reinforcement learning agent

All components run client-side in Vanilla JavaScript with <50ms latency for full ML inference.

## Project Structure

- `gold-bot.html` - Main browser-based application interface
- `report.js` - Node.js script to generate thesis documentation in Word format
- `package.json` - Dependencies (docx library for document generation)

## Quick Start

### Running the Application

Open `gold-bot.html` in a modern web browser to launch the real-time trading interface.

### Generating Documentation

To generate the thesis in Word format:

```bash
npm install
node report.js
```

This creates `ENALYSIS_Thesis.docx` with complete technical documentation.

## Features

- ✅ Real-time XAU/USD price streaming via WebSocket
- ✅ 12-dimensional feature engineering pipeline
- ✅ Three-model ensemble voting mechanism
- ✅ <50ms end-to-end latency (95th percentile)
- ✅ 60 FPS real-time dashboard rendering
- ✅ No external ML dependencies (Random Forest, Kalman Filter, Q-Learning from scratch)
- ✅ Smart Money Concepts liquidity zone detection
- ✅ Regime-aware risk management

## Technical Architecture

### Data Pipeline
- REST API for historical data (Twelve Data)
- WebSocket for real-time tick streaming
- Hybrid fallback for connectivity resilience

### ML Models
- **Random Forest**: 20 trees, max depth 5, 12-dimensional feature space
- **Kalman Filter**: Scalar state estimation with process/observation noise tuning
- **Q-Learning**: Tabular 108-state discretized action space

### Signal Generation
- Weighted ensemble voting with rolling accuracy weights
- Volume profile-based support/resistance levels
- ATR-based risk management with regime adaptation
- Confidence calibration across market conditions

## Performance Metrics

- Directional accuracy: Tested across multiple market regimes
- Latency: <50ms for complete ML pipeline + rendering
- Memory footprint: <5MB for full system
- Browser compatibility: Chrome, Firefox, Safari, Edge (ES6+)

## Research Gap Addressed

This project bridges the gap between:
1. Architecturally-constrained ML in browser environments
2. Heterogeneous ensemble design for financial forecasting
3. Real-time regime detection and risk adaptation
4. Integration of Smart Money Concepts with quantitative signals

## Deployment

The system is designed as a single-file standalone application. Simply serve `gold-bot.html` on any web server or open locally.

## License

Academic use - B.Tech Capstone Project

## Authors

Anuja Verma - CSJMU, UIET, Computer Science & Engineering, 2024-25

---

**Note**: This system generates trading signals only. It does not execute trades and should not be used as sole basis for investment decisions without proper risk management and professional financial advice.
