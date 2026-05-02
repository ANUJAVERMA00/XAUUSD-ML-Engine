# ENALYSIS · Gold ML Engine

> **Live Dashboard:** [xauusd-bot.dpsharma.workers.dev](https://xauusd-bot.dpsharma.workers.dev)  
> **Live API:** [enalysis-gold-api.dpsharma.workers.dev](https://enalysis-gold-api.dpsharma.workers.dev/health)

A real-time **XAU/USD (Gold) trading analysis dashboard** hosted on Cloudflare Workers. It combines live price feeds, classical technical indicators, and a multi-model ML ensemble (Random Forest, Gradient Boosting, Kalman Filter, Reinforcement Learning) to generate BUY / SELL / NEUTRAL signals with confidence scores, take-profit, and stop-loss levels — all in a single-file HTML app with zero build steps.

---

##  Features

| Feature | Details |
|---|---|
| **Live Price Feed** | WebSocket + REST polling via [Twelve Data API](https://twelvedata.com) |
| **Technical Indicators** | RSI, MACD, Bollinger Bands, ATR, EMA (9/21/50), ADX, Stochastic |
| **ML Ensemble** | Random Forest · Gradient Boosting · Kalman Filter · Q-Learning (RL) |
| **Signal Output** | BUY / SELL / NEUTRAL with confidence %, TP, SL, and Risk:Reward ratio |
| **Support & Resistance** | Auto-detected zones with liquidity levels |
| **Market Regime** | Trend / Range / Volatile classification |
| **Performance Tracking** | Trades, Win Rate, PnL, Max Drawdown |
| **Activity Log** | Real-time tagged log (TRADE · SIGNAL · ML · RL · INFO · ERROR) |
| **Report Generator** | Node.js script (`report.js`) to export a formatted `.docx` report |
| **Hosted on Cloudflare** | Zero cold-starts, global edge deployment |

---

## Live Demo

| Resource | URL |
|---|---|
|  Dashboard | https://xauusd-bot.dpsharma.workers.dev |
|  API Health | https://enalysis-gold-api.dpsharma.workers.dev/health |
|  Signal | https://enalysis-gold-api.dpsharma.workers.dev/signal |
|  Signal (1h) | https://enalysis-gold-api.dpsharma.workers.dev/signal?tf=1h |
|  Indicators | https://enalysis-gold-api.dpsharma.workers.dev/indicators |
|  Zones | https://enalysis-gold-api.dpsharma.workers.dev/zones |

---

## API Reference

Base URL: `https://enalysis-gold-api.dpsharma.workers.dev`

### `GET /health`
Returns service status and live XAU/USD price.

```json
{
  "status": "ok",
  "service": "ENALYSIS Gold ML API",
  "symbol": "XAU/USD",
  "live_price": 4613.25,
  "timestamp": "2026-05-02T08:26:17.819Z",
  "endpoints": ["/signal", "/signal?tf=1h", "/indicators", "/zones", "/health"],
  "version": "2.0"
}
```

### `GET /signal`
Returns the current ML ensemble signal.

```json
{
  "signal": "BUY",
  "confidence": 78,
  "entry": 4610.50,
  "tp": 4645.00,
  "sl": 4590.00,
  "rr": 1.72
}
```

### `GET /signal?tf=1h`
Same as `/signal` but computed on the specified timeframe.  
Supported values: `1m` `5m` `15m` `1h` `4h` `1D`

### `GET /indicators`
Returns all computed technical indicator values (RSI, MACD, BB, ATR, EMA, ADX, Stochastic).

### `GET /zones`
Returns auto-detected support, resistance, and liquidity zones.

---

## ML Models Explained

### Random Forest (RF)
An ensemble of decision trees trained on rolling windows of OHLCV + indicator features. Each tree votes on direction (BUY/SELL/NEUTRAL) and the majority label becomes the RF signal.

### Gradient Boosting (GB)
Sequentially trained weak learners that minimize prediction error. More sensitive to recent price action than RF — tends to react faster to regime changes.

### Kalman Filter (KF)
A recursive Bayesian estimator used to smooth noisy price data. The filter's predicted vs. actual deviation generates a mean-reversion signal.

### Reinforcement Learning (Q-Learning)
An ε-greedy Q-learning agent with a discrete state space derived from RSI bucket, MACD direction, and EMA alignment. Learns optimal HOLD/BUY/SELL actions through simulated episode rewards.

### Ensemble
All four model outputs are weighted by their recent confidence scores and combined into a single consensus signal.

---

##  Signal Interpretation

```
Signal:  BUY / SELL / NEUTRAL
Entry:   Suggested entry price
TP:      Take-Profit target
SL:      Stop-Loss level
R:R:     Risk-to-Reward ratio
Conf:    Ensemble confidence (0–100%)
```

Signals are regenerated automatically every minute (configurable via `autoAnalyze` in source).

---

##  Run Locally

### Dashboard (Browser)

```bash
git clone https://github.com/YOUR_USERNAME/gold-ml-engine.git
cd gold-ml-engine

# No server needed — open directly
open gold-bot.html        # macOS
start gold-bot.html       # Windows
xdg-open gold-bot.html    # Linux
```

> Replace `API_KEY` near the top of `gold-bot.html` with your own key from [twelvedata.com](https://twelvedata.com) if needed.

### Report Generator (Node.js)

```bash
npm install
node report.js
# Outputs a formatted .docx report in the project root
```

---

##  Deploy to Cloudflare Workers

```bash
npm install -g wrangler
wrangler login
wrangler deploy
```

The dashboard and API are each deployed as separate Cloudflare Workers under the `dpsharma.workers.dev` subdomain.

---

##  Project Structure

```
gold-ml-engine/
├── gold-bot.html       # Main dashboard — self-contained single-file app
├── report.js           # Cloudflare Worker API + DOCX report generator
├── package.json        # Node dependencies (docx library)
└── node_modules/       # Auto-generated after npm install
```

---

##  Configuration

Key settings at the top of `gold-bot.html`:

```js
const API_KEY = 'YOUR_TWELVE_DATA_API_KEY';
const SYMBOL  = 'XAU/USD';   // Any Twelve Data supported symbol
const BASE    = 'https://api.twelvedata.com';
```

Timeframe can be switched live from the UI: `1m · 5m · 15m · 1h · 4h · 1D`

---

##  Dependencies

| Context | Package | Purpose |
|---|---|---|
| Browser | *(none)* | Dashboard is vanilla JS — zero dependencies |
| Node.js | `docx ^9.6.1` | DOCX report generation |

---

##  Disclaimer

This project is for **educational and research purposes only**. The signals generated are not financial advice. Gold trading involves significant risk. Always do your own research and consult a qualified financial advisor before making any trades.

---

##  License

MIT — free to use, modify, and distribute.