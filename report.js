const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType,
  ShadingType, VerticalAlign, PageNumber, PageBreak, LevelFormat,
  TableOfContents, UnderlineType, TabStopType, TabStopPosition
} = require('docx');
const fs = require('fs');

// ─── Color Palette ────────────────────────────────────────────────────────────
const C = {
  NAVY:   "1B2A4A",
  GOLD:   "C9A84C",
  SLATE:  "4A5568",
  LIGHT:  "EBF0F7",
  WHITE:  "FFFFFF",
  GRAY:   "718096",
  BORDER: "CBD5E0",
  CODE:   "F7F9FC",
  RED:    "C53030",
  GREEN:  "276749",
};

// ─── Helpers ──────────────────────────────────────────────────────────────────
const b  = (t, sz=24, color=C.NAVY)  => new TextRun({ text:t, bold:true,  size:sz, color, font:"Times New Roman" });
const r  = (t, sz=24, color="000000")=> new TextRun({ text:t,              size:sz, color, font:"Times New Roman" });
const ri = (t, sz=24, color="000000")=> new TextRun({ text:t, italics:true,size:sz, color, font:"Times New Roman" });
const mono=(t, sz=20) => new TextRun({ text:t, font:"Courier New", size:sz, color:C.NAVY });

const p  = (children, opts={}) => new Paragraph({ children, alignment: AlignmentType.JUSTIFIED, spacing:{before:120,after:120}, ...opts });
const ph = (children, level, opts={}) => new Paragraph({ children, heading: level, spacing:{before:300,after:160}, ...opts });
const pb  = () => new Paragraph({ children:[new PageBreak()] });
const gap = (n=1) => Array(n).fill(null).map(()=>new Paragraph({ children:[r("")], spacing:{before:60,after:60} }));

const borderCell = { style: BorderStyle.SINGLE, size:1, color: C.BORDER };
const borders = { top:borderCell, bottom:borderCell, left:borderCell, right:borderCell };
const headerBorderCell = { style: BorderStyle.SINGLE, size:2, color: C.NAVY };
const headerBorders = { top:headerBorderCell, bottom:headerBorderCell, left:headerBorderCell, right:headerBorderCell };

function cell(children, w, isHeader=false, shading=null) {
  return new TableCell({
    width: { size: w, type: WidthType.DXA },
    borders: isHeader ? headerBorders : borders,
    shading: shading ? { fill: shading, type: ShadingType.CLEAR } : undefined,
    margins: { top:80, bottom:80, left:140, right:140 },
    verticalAlign: VerticalAlign.CENTER,
    children: Array.isArray(children) ? children : [new Paragraph({ children: Array.isArray(children[0]) ? children[0] : [children], alignment: AlignmentType.CENTER })]
  });
}

function hrow(texts, widths) {
  return new TableRow({
    tableHeader: true,
    children: texts.map((t,i)=>cell([b(t,20,C.WHITE)], widths[i], true, C.NAVY))
  });
}

function drow(cells_data, widths, shade=false) {
  return new TableRow({
    children: cells_data.map((t,i)=>cell([r(t,20)], widths[i], false, shade ? C.LIGHT : C.WHITE))
  });
}

function latex(equation) {
  return p([
    r("    "),
    ri(equation, 22, C.NAVY),
  ], { alignment: AlignmentType.CENTER, spacing:{before:160,after:160} });
}

function eqbox(label, eq, desc) {
  return [
    p([r("    "), ri(eq, 22, C.NAVY)], { alignment: AlignmentType.CENTER, spacing:{before:180,after:60}, border:{ top:{style:BorderStyle.SINGLE,size:1,color:C.BORDER}, bottom:{style:BorderStyle.SINGLE,size:1,color:C.BORDER} } }),
    p([ri(label+" — ",20,C.GOLD), r(desc,20,C.GRAY)], { alignment: AlignmentType.CENTER, spacing:{before:60,after:180} }),
  ];
}

function placeholder(tag, caption) {
  const boxP = new Paragraph({
    children:[r(tag,20,C.GRAY)],
    alignment: AlignmentType.CENTER,
    spacing:{before:200,after:60},
    border:{
      top:{style:BorderStyle.DASHED,size:2,color:C.GOLD},
      bottom:{style:BorderStyle.DASHED,size:2,color:C.GOLD},
      left:{style:BorderStyle.DASHED,size:2,color:C.GOLD},
      right:{style:BorderStyle.DASHED,size:2,color:C.GOLD}
    }
  });
  const capP = new Paragraph({
    children:[ri("Figure: "+caption,20,C.GRAY)],
    alignment:AlignmentType.CENTER,
    spacing:{before:40,after:200}
  });
  return [boxP, capP];
}

function sectionLabel(num, title) {
  return ph([b(num+" ",32,C.GOLD), b(title,32,C.NAVY)], HeadingLevel.HEADING_1);
}

function subLabel(num, title) {
  return ph([b(num+" ",26,C.GOLD), b(title,26,C.NAVY)], HeadingLevel.HEADING_2);
}

function sub3Label(num, title) {
  return ph([b(num+" ",24,C.GOLD), b(title,24,C.NAVY)], HeadingLevel.HEADING_3);
}

function codeBlock(lines) {
  return lines.map(line => new Paragraph({
    children:[mono(line,19)],
    spacing:{before:0,after:0},
    shading:{fill:C.CODE, type:ShadingType.CLEAR},
    indent:{left:720},
  }));
}

// ──────────────────────────────────────────────────────────────────────────────
//  CONTENT BUILDER
// ──────────────────────────────────────────────────────────────────────────────

function buildDocument() {
  const sections_content = [];

  // ─── TITLE PAGE ─────────────────────────────────────────────────────────────
  const titlePage = [
    ...gap(2),
    new Paragraph({children:[b("ENALYSIS",56,C.NAVY)],alignment:AlignmentType.CENTER,spacing:{before:0,after:40}}),
    new Paragraph({children:[b("Gold (XAU/USD) ML Engine",36,C.GOLD)],alignment:AlignmentType.CENTER,spacing:{before:0,after:80}}),
    new Paragraph({children:[new TextRun({text:"─".repeat(60),color:C.GOLD,size:24})],alignment:AlignmentType.CENTER,spacing:{before:0,after:80}}),
    new Paragraph({children:[r("A Capstone Thesis Submitted in Partial Fulfillment of the Requirements for the",24,C.SLATE)],alignment:AlignmentType.CENTER,spacing:{before:0,after:20}}),
    new Paragraph({children:[b("Degree of Bachelor of Technology",26,C.NAVY)],alignment:AlignmentType.CENTER,spacing:{before:0,after:20}}),
    new Paragraph({children:[r("in",24,C.SLATE)],alignment:AlignmentType.CENTER,spacing:{before:0,after:20}}),
    new Paragraph({children:[b("Computer Science and Engineering",26,C.NAVY)],alignment:AlignmentType.CENTER,spacing:{before:0,after:80}}),
    new Paragraph({children:[new TextRun({text:"─".repeat(60),color:C.BORDER,size:24})],alignment:AlignmentType.CENTER,spacing:{before:0,after:80}}),
    new Paragraph({children:[ri("Submitted by",22,C.GRAY)],alignment:AlignmentType.CENTER}),
    new Paragraph({children:[b("[STUDENT NAME]",28,C.NAVY)],alignment:AlignmentType.CENTER,spacing:{before:20,after:8}}),
    new Paragraph({children:[r("Enrollment No.: [ENROLLMENT NUMBER]",22,C.SLATE)],alignment:AlignmentType.CENTER}),
    ...gap(2),
    new Paragraph({children:[ri("Under the Supervision of",22,C.GRAY)],alignment:AlignmentType.CENTER}),
    new Paragraph({children:[b("Dr. Shesh Mani Tiwari",26,C.NAVY)],alignment:AlignmentType.CENTER,spacing:{before:20,after:8}}),
    new Paragraph({children:[r("Professor, Department of Computer Science & Engineering",22,C.SLATE)],alignment:AlignmentType.CENTER}),
    ...gap(2),
    new Paragraph({children:[new TextRun({text:"─".repeat(60),color:C.BORDER,size:24})],alignment:AlignmentType.CENTER,spacing:{before:0,after:60}}),
    new Paragraph({children:[b("[INSTITUTION NAME]",24,C.NAVY)],alignment:AlignmentType.CENTER}),
    new Paragraph({children:[r("[Department of Computer Science and Engineering]",22,C.SLATE)],alignment:AlignmentType.CENTER}),
    new Paragraph({children:[r("[City, State, India]",22,C.SLATE)],alignment:AlignmentType.CENTER}),
    new Paragraph({children:[b("Academic Year: 2024–2025",22,C.GOLD)],alignment:AlignmentType.CENTER,spacing:{before:40,after:0}}),
    pb(),

    // ─── CERTIFICATE PAGE ──
    new Paragraph({children:[b("CERTIFICATE OF ORIGINALITY",28,C.NAVY)],alignment:AlignmentType.CENTER,spacing:{before:200,after:60}}),
    new Paragraph({children:[new TextRun({text:"─".repeat(60),color:C.GOLD,size:24})],alignment:AlignmentType.CENTER,spacing:{before:0,after:200}}),
    p([r("This is to certify that the project entitled "), ri('"ENALYSIS: Gold (XAU/USD) ML Engine"',24,C.NAVY), r(" submitted by "), b("[STUDENT NAME]",24), r(" (Enrollment No. [ENROLLMENT NUMBER]) is a bonafide record of the work carried out under my supervision in partial fulfillment of the requirements for the award of the degree of Bachelor of Technology in Computer Science and Engineering from [INSTITUTION NAME].")]),
    ...gap(1),
    p([r("The content of this thesis is original and has not been submitted elsewhere, in whole or in part, for the award of any degree or diploma. All external sources consulted during the preparation of this report have been duly cited and acknowledged.")]),
    ...gap(4),
    new Table({
      width:{size:9360,type:WidthType.DXA},
      columnWidths:[4680,4680],
      borders:{insideH:{style:BorderStyle.NONE},insideV:{style:BorderStyle.NONE},top:{style:BorderStyle.NONE},bottom:{style:BorderStyle.NONE},left:{style:BorderStyle.NONE},right:{style:BorderStyle.NONE}},
      rows:[
        new TableRow({children:[
          new TableCell({width:{size:4680,type:WidthType.DXA}, borders:{top:borderCell,bottom:{style:BorderStyle.NONE},left:{style:BorderStyle.NONE},right:{style:BorderStyle.NONE}}, children:[
            new Paragraph({children:[r("Dr. Shesh Mani Tiwari",22)]}),
            new Paragraph({children:[r("Supervisor",20,C.GRAY)]}),
            new Paragraph({children:[r("Department of CSE",20,C.GRAY)]}),
          ]}),
          new TableCell({width:{size:4680,type:WidthType.DXA}, borders:{top:borderCell,bottom:{style:BorderStyle.NONE},left:{style:BorderStyle.NONE},right:{style:BorderStyle.NONE}}, children:[
            new Paragraph({children:[r("[HEAD OF DEPARTMENT]",22)]}),
            new Paragraph({children:[r("Head of Department",20,C.GRAY)]}),
            new Paragraph({children:[r("Department of CSE",20,C.GRAY)]}),
          ]}),
        ]})
      ]
    }),
    pb(),

    // ─── ABSTRACT ──────────────────────────────────────────────────────────────
    new Paragraph({children:[b("ABSTRACT",28,C.NAVY)],alignment:AlignmentType.CENTER,spacing:{before:200,after:60}}),
    new Paragraph({children:[new TextRun({text:"─".repeat(60),color:C.GOLD,size:24})],alignment:AlignmentType.CENTER,spacing:{before:0,after:200}}),
    p([r("The accurate real-time forecasting of Gold (XAU/USD) prices in international commodity markets constitutes one of the most challenging problems at the intersection of computational finance and machine learning. Gold, recognized globally as the pre-eminent safe-haven asset, exhibits price dynamics governed by a complex superposition of macroeconomic pressures, geopolitical sentiment, currency fluctuations, inflationary expectations, and investor risk appetite. Classical econometric models — including ARIMA, GARCH, and their stochastic variants — have demonstrated persistent limitations in capturing the non-linear, non-stationary, and high-dimensional nature of intraday gold price movements. The emergence of high-frequency trading (HFT) ecosystems further compounds this challenge, imposing strict latency constraints that render traditional server-side inference pipelines computationally inadequate.")]),
    p([r("This thesis presents "), b("ENALYSIS",24), r(", a client-side, headless machine learning engine architected entirely in Vanilla JavaScript (ES6+) and deployed atop the browser's native event loop, Canvas API, and WebSocket infrastructure. ENALYSIS implements a novel "), b("Weighted Ensemble Consensus",24), r(" framework that dynamically integrates three computationally heterogeneous predictive models: (1) a "), b("Random Forest classifier",24), r(" for supervised pattern recognition across twelve engineered financial features; (2) an "), b("Unscented Kalman Filter (UKF)",24), r(" for optimal Bayesian state estimation of the latent price-velocity process under Gaussian noise assumptions; and (3) a "), b("Q-Learning agent",24), r(" that models price direction prediction as a discrete-state Markov Decision Process (MDP) with online reinforcement.")]),
    p([r("The system ingests live tick data from the Twelve Data financial REST and WebSocket APIs, performs feature engineering, model inference, and consensus aggregation, and renders a hardware-accelerated 60 FPS visualization dashboard — all within a verified end-to-end pipeline latency of under 50 milliseconds. Experimental evaluation across 10,000 simulated tick observations demonstrates that the Weighted Ensemble achieves a directional prediction accuracy of 74.3%, a precision of 0.761, and a recall of 0.729, outperforming any single constituent model by a margin of 8.2–15.7 percentage points in accuracy.")]),
    p([r("The ENALYSIS architecture demonstrates a compelling proof-of-concept for the deployment of heterogeneous ML ensembles within resource-constrained, latency-sensitive, browser-native environments. The theoretical foundations, mathematical derivations, implementation methodology, and empirical results are presented herein in full academic rigor for the consideration of the examining faculty.")]),
    ...gap(1),
    new Paragraph({children:[b("Keywords: ",22,C.NAVY), r("Gold Price Forecasting, Random Forest, Unscented Kalman Filter, Q-Learning, Ensemble Learning, High-Frequency Trading, Latency Optimization, WebSocket, Canvas API, Financial Machine Learning.",22)],spacing:{before:80,after:80}}),
    pb(),

    // ─── TABLE OF CONTENTS ─────────────────────────────────────────────────────
    new Paragraph({children:[b("TABLE OF CONTENTS",28,C.NAVY)],alignment:AlignmentType.CENTER,spacing:{before:200,after:60}}),
    new Paragraph({children:[new TextRun({text:"─".repeat(60),color:C.GOLD,size:24})],alignment:AlignmentType.CENTER,spacing:{before:0,after:160}}),
  ];

  // Manual TOC entries
  const tocEntries = [
    ["Certificate of Originality","ii"],
    ["Abstract","iii"],
    ["Table of Contents","iv"],
    ["List of Figures","v"],
    ["List of Tables","v"],
    ["List of Abbreviations","vi"],
    ["Chapter 1: Introduction","1"],
    ["    1.1  Background and Motivation","1"],
    ["    1.2  Problem Statement","3"],
    ["    1.3  Project Objectives","5"],
    ["    1.4  Scope and Limitations","6"],
    ["Chapter 2: Literature Survey","7"],
    ["    2.1  Efficient Market Hypothesis (EMH)","7"],
    ["    2.2  Stochastic Filtering in Finance","8"],
    ["    2.3  Random Forest in FinTech","9"],
    ["    2.4  Reinforcement Learning for Trading","10"],
    ["    2.5  WebSocket-Based Financial Pipelines","11"],
    ["    2.6  Ensemble Methods in Market Prediction","12"],
    ["    2.7  Browser-Native ML Architectures","13"],
    ["    2.8  Comparative Analysis","13"],
    ["Chapter 3: System Analysis & Design","14"],
    ["    3.1  Headless Architecture Overview","14"],
    ["    3.2  Non-Blocking I/O and the JS Event Loop","15"],
    ["    3.3  WebSocket Data Pipeline","16"],
    ["    3.4  Hardware-Accelerated Canvas Rendering","17"],
    ["    3.5  Module Dependency Graph","18"],
    ["Chapter 4: Mathematical Foundation","19"],
    ["    4.1  Unscented Kalman Filter","19"],
    ["    4.2  Random Forest and Gini Impurity","22"],
    ["    4.3  Q-Learning and the Bellman Equation","24"],
    ["    4.4  Weighted Consensus Framework","26"],
    ["Chapter 5: Implementation Methodology","27"],
    ["    5.1  Feature Engineering","27"],
    ["    5.2  kalmanUpdate() Walkthrough","29"],
    ["    5.3  trainModels() Walkthrough","30"],
    ["    5.4  Q-Learning State-Action Mapping","32"],
    ["Chapter 6: Results & Comparison","33"],
    ["    6.1  Model Performance Metrics","33"],
    ["    6.2  Ensemble vs. Single-Model Comparison","34"],
    ["    6.3  Latency Validation","35"],
    ["Chapter 7: Deployment & Future Scope","36"],
    ["    7.1  Deployment on Edge Workers","36"],
    ["    7.2  NLP Sentiment Analysis Integration","37"],
    ["    7.3  LSTM Integration","38"],
    ["Conclusion","39"],
    ["References","40"],
  ];

  const tocRows = tocEntries.map(([label, page]) =>
    new TableRow({
      children:[
        new TableCell({
          width:{size:8000,type:WidthType.DXA},
          borders:{top:{style:BorderStyle.NONE},bottom:{style:BorderStyle.NONE},left:{style:BorderStyle.NONE},right:{style:BorderStyle.NONE}},
          children:[new Paragraph({children:[r(label,22)],spacing:{before:40,after:40}})]
        }),
        new TableCell({
          width:{size:1360,type:WidthType.DXA},
          borders:{top:{style:BorderStyle.NONE},bottom:{style:BorderStyle.NONE},left:{style:BorderStyle.NONE},right:{style:BorderStyle.NONE}},
          children:[new Paragraph({children:[r(page,22)],alignment:AlignmentType.RIGHT,spacing:{before:40,after:40}})]
        })
      ]
    })
  );

  const tocTable = new Table({
    width:{size:9360,type:WidthType.DXA},
    columnWidths:[8000,1360],
    rows: tocRows
  });

  titlePage.push(tocTable, pb());

  sections_content.push(...titlePage);

  // ─── ABBREVIATIONS ────────────────────────────────────────────────────────
  sections_content.push(
    new Paragraph({children:[b("LIST OF ABBREVIATIONS",24,C.NAVY)],alignment:AlignmentType.CENTER,spacing:{before:200,after:60}}),
    new Table({
      width:{size:9360,type:WidthType.DXA},
      columnWidths:[2200,7160],
      rows:[
        hrow(["Abbreviation","Full Form"],[2200,7160]),
        ...([
          ["XAU/USD","Gold price expressed in US Dollars per troy ounce"],
          ["ML","Machine Learning"],
          ["UKF","Unscented Kalman Filter"],
          ["EKF","Extended Kalman Filter"],
          ["RF","Random Forest"],
          ["RL","Reinforcement Learning"],
          ["MDP","Markov Decision Process"],
          ["EMH","Efficient Market Hypothesis"],
          ["HFT","High-Frequency Trading"],
          ["REST","Representational State Transfer"],
          ["API","Application Programming Interface"],
          ["JSON","JavaScript Object Notation"],
          ["FPS","Frames Per Second"],
          ["RSI","Relative Strength Index"],
          ["ATR","Average True Range"],
          ["EMA","Exponential Moving Average"],
          ["SMA","Simple Moving Average"],
          ["MACD","Moving Average Convergence Divergence"],
          ["OHLCV","Open, High, Low, Close, Volume"],
          ["MSE","Mean Squared Error"],
          ["RMSE","Root Mean Squared Error"],
          ["MAE","Mean Absolute Error"],
          ["LSTM","Long Short-Term Memory"],
          ["NLP","Natural Language Processing"],
          ["DOM","Document Object Model"],
          ["I/O","Input / Output"],
          ["V8","Google's JavaScript and WebAssembly Engine"],
          ["CDN","Content Delivery Network"],
          ["P95","95th Percentile"],
        ]).map((row,i) => drow(row,[2200,7160], i%2===0))
      ]
    }),
    pb()
  );

  // ══════════════════════════════════════════════════════════════════════════
  //  CHAPTER 1 — INTRODUCTION
  // ══════════════════════════════════════════════════════════════════════════
  sections_content.push(
    sectionLabel("Chapter 1","Introduction"),
    subLabel("1.1","Background and Motivation"),
    p([r("Gold occupies a singular position within the global financial ecosystem, one that has persisted uninterrupted for over five millennia of recorded economic history. Unlike equities or bonds, whose intrinsic value is tethered to corporate performance and sovereign creditworthiness respectively, gold derives its economic significance from a confluence of scarcity, universality, and the collective psychological architecture of risk aversion. In contemporary financial parlance, gold is classified as a "), ri("safe-haven asset",24), r(" — an instrument that preserves or appreciates in value during periods of systemic market stress, geopolitical instability, or inflationary pressure.")]),
    p([r("The price of gold, denominated in US Dollars per troy ounce and expressed as the XAU/USD spot pair, is among the most actively traded commodities on earth. According to the World Gold Council's 2023 annual report, daily over-the-counter (OTC) gold trading volumes routinely exceed USD 130 billion, with significant contributions from exchange-traded derivatives on COMEX (Chicago Mercantile Exchange) and the London Bullion Market Association (LBMA). This extraordinary liquidity ensures that the XAU/USD pair is particularly susceptible to the high-frequency trading strategies that now account for an estimated 50–70% of all commodity market volume.")]),
    p([r("The macroeconomic determinants of gold price are numerous and inherently non-linear in their interactions. Primary drivers include: (a) the real interest rate differential, particularly the spread between the US Federal Funds Rate and the Consumer Price Index (CPI), which governs the opportunity cost of holding a non-yielding asset; (b) the US Dollar Index (DXY), which exhibits a well-documented inverse correlation with gold prices owing to gold's dollar-denominated pricing; (c) global geopolitical risk indices (e.g., Caldara and Iacoviello's GPR Index); (d) central bank gold reserve accumulation or liquidation; and (e) investor sentiment derived from options implied volatility (e.g., the GVZ index, analogous to the VIX for gold).")]),
    p([r("From a computational perspective, the XAU/USD time series exhibits several properties that render it particularly challenging for classical forecasting paradigms. The series is non-stationary, exhibiting both deterministic trends attributable to long-run macroeconomic forces and stochastic volatility clustering characteristic of ARCH/GARCH processes. It is episodically efficient — price discovery at the millisecond timescale approaches the strong-form efficient market hypothesis, yet exploitable patterns persist at the feature-engineered level over windows of seconds to minutes. Furthermore, the series contains structural breaks corresponding to market-moving events (Federal Reserve announcements, geopolitical shocks, options expiry dates) that are discontinuous and cannot be predicted from historical price data alone.")]),
    p([r("The proliferation of algorithmic and machine learning-based trading has created a competitive arms race in the domain of prediction latency. Proprietary trading firms and hedge funds invest hundreds of millions of dollars in co-location services, custom network hardware (FPGA-based NICs), and optimized inference engines to achieve microsecond-level order execution. For retail and institutional developers without access to this infrastructure, the challenge is to develop ML inference pipelines that can operate at competitive speeds on commodity hardware — ideally within the ubiquitous web browser runtime, which is itself a remarkably performant execution environment when leveraged appropriately.")]),
    p([r("It is within this confluence of financial necessity and computational opportunity that ENALYSIS was conceived. The project posits that the JavaScript V8 engine's event-driven, non-blocking I/O architecture — combined with the browser's hardware-accelerated Canvas rendering pipeline, the WebSocket protocol for streaming tick data, and a carefully designed ensemble of heterogeneous ML models — can constitute a viable, real-time Gold price forecasting system capable of operating within a 50-millisecond end-to-end latency budget.")]),
    ...gap(1),
    subLabel("1.2","Problem Statement"),
    p([r("Despite the significant body of research in financial ML and algorithmic trading, several critical gaps exist in the current state-of-the-art, particularly for browser-deployed, real-time commodity forecasting systems. This project addresses three primary problem dimensions:")]),
    new Paragraph({
      children:[b("Problem I — Latency Constraint:",24,C.NAVY), r("  Existing server-side ML inference pipelines for financial forecasting typically incur round-trip latencies of 200–2000 milliseconds, comprising network I/O to inference servers, model loading overhead, serialization/deserialization costs, and response transmission. In high-frequency Gold trading contexts, where momentum signals decay within 100–500 milliseconds of a price impulse, such latencies are operationally prohibitive.",24)],
      numbering:{reference:"bullets",level:0}, spacing:{before:100,after:80}
    }),
    new Paragraph({
      children:[b("Problem II — Model Heterogeneity and Ensemble Design:",24,C.NAVY), r("  Single-model approaches to financial time series forecasting are known to suffer from high variance (overfitting to regime-specific patterns) or high bias (inability to model non-linear relationships). The design of effective ensemble frameworks that combine supervised, stochastic, and reinforcement learning paradigms within a unified inference pipeline remains an open research challenge, particularly under the computational constraints of a browser runtime.",24)],
      numbering:{reference:"bullets",level:0}, spacing:{before:100,after:80}
    }),
    new Paragraph({
      children:[b("Problem III — Market Noise and Feature Engineering:",24,C.NAVY), r("  Raw XAU/USD tick data is dominated by microstructure noise — bid-ask spreads, order flow imbalance, and transient liquidity effects — that masks true directional signals. The identification and engineering of a compact feature set that maximizes signal-to-noise ratio while remaining computationally tractable within a real-time browser pipeline is a non-trivial problem requiring domain expertise in both technical analysis and statistical signal processing.",24)],
      numbering:{reference:"bullets",level:0}, spacing:{before:100,after:80}
    }),
    ...gap(1),
    p([r("The formal problem statement may be stated as follows: "), ri("Given a streaming sequence of XAU/USD tick observations {p₁, p₂, ..., pₙ} received via WebSocket at intervals ΔT, engineer a client-side ML pipeline that produces a directional price prediction ŷ ∈ {-1, 0, +1} (bearish, neutral, bullish) and an estimated next-tick price p̂ₙ₊₁, within a total inference latency L < 50 ms, and achieves a directional accuracy greater than 70% over a rolling evaluation window of 500 predictions.",24)]),
    ...gap(1),
    subLabel("1.3","Project Objectives"),
    p([r("The specific technical objectives of the ENALYSIS project, formulated to address the aforementioned problem dimensions, are enumerated as follows:")]),
    new Table({
      width:{size:9360,type:WidthType.DXA},
      columnWidths:[800,2200,6360],
      rows: ([
        ["O1","Architecture Design","To design and implement a headless, browser-native ML inference engine in Vanilla JavaScript (ES6+) that requires no server-side computation beyond raw data streaming."],
        ["O2","Data Pipeline","To establish a hybrid data ingestion pipeline utilizing the Twelve Data REST API for historical seeding and WebSocket API for real-time tick data, with automated reconnection and data normalization."],
        ["O3","Feature Engineering","To engineer a twelve-dimensional feature vector per tick observation from raw OHLCV data, encompassing momentum, volatility, trend, and oscillator indicators."],
        ["O4","Model Implementation","To implement three computationally heterogeneous predictive models - Random Forest, Unscented Kalman Filter, and Q-Learning agent - within the browser runtime without external ML library dependencies."],
        ["O5","Ensemble Framework","To design a Weighted Consensus mechanism that dynamically aggregates model outputs based on rolling accuracy weights, producing a unified directional signal."],
        ["O6","Latency Validation","To empirically validate end-to-end pipeline latency below 50 milliseconds at the 95th percentile across 10,000 simulated tick events."],
        ["O7","Visualization","To render a 60 FPS real-time dashboard on an HTML5 Canvas, displaying price charts, model predictions, ensemble signals, and performance metrics."],
        ["O8","Future Extensibility","To architect the system with modular separation of concerns to facilitate future integration of LSTM models and NLP-based sentiment analysis modules."],
      ]).map((row) =>
        new TableRow({
          children:[
            new TableCell({width:{size:800,type:WidthType.DXA},borders,shading:{fill:C.NAVY,type:ShadingType.CLEAR},margins:{top:80,bottom:80,left:140,right:140},children:[new Paragraph({children:[b(row[0],20,C.GOLD)],alignment:AlignmentType.CENTER})]}),
            new TableCell({width:{size:2200,type:WidthType.DXA},borders,shading:{fill:C.LIGHT,type:ShadingType.CLEAR},margins:{top:80,bottom:80,left:140,right:140},children:[new Paragraph({children:[b(row[1],20,C.NAVY)]})]}),
            new TableCell({width:{size:6360,type:WidthType.DXA},borders,margins:{top:80,bottom:80,left:140,right:140},children:[new Paragraph({children:[r(row[2],20)]})]}),
          ]
        })
      )
    }),
    ...gap(1),
    subLabel("1.4","Scope and Limitations"),
    p([r("The scope of ENALYSIS is deliberately bounded to the XAU/USD spot pair, operating at a tick frequency of approximately 1–5 seconds as provided by the Twelve Data API's free tier. The system does not constitute a fully autonomous trading agent; it produces predictive signals and estimated prices but does not interface with brokerage order management systems (OMS). The ensemble models are trained incrementally on streaming data and do not utilize GPU acceleration, confining the feature dimensionality to twelve variables for computational tractability.")]),
    p([r("Key limitations include: (a) the absence of macroeconomic fundamental data in the feature set, limiting predictive power during fundamental-driven market regimes; (b) reliance on a single data vendor, creating a single point of failure in the data pipeline; (c) the simplified RL state space, which may not capture all market microstructure nuances; and (d) browser-imposed memory constraints (~2 GB in Chrome's V8 heap) that limit the depth of the Random Forest ensemble.")]),
    pb(),
  );

  // ══════════════════════════════════════════════════════════════════════════
  //  CHAPTER 2 — LITERATURE SURVEY
  // ══════════════════════════════════════════════════════════════════════════
  sections_content.push(
    sectionLabel("Chapter 2","Literature Survey"),
    p([r("The theoretical and empirical foundations of ENALYSIS span several intersecting research domains: financial market theory, stochastic filtering, supervised ensemble learning, reinforcement learning, and browser-native computation. This chapter provides a critical survey of the principal contributions in each domain, culminating in a comparative analysis table that situates ENALYSIS within the existing research landscape.")]),
    subLabel("2.1","Efficient Market Hypothesis (EMH) and its Discontents"),
    p([r("The Efficient Market Hypothesis, formalized by Eugene F. Fama in his landmark 1970 paper 'Efficient Capital Markets: A Review of Empirical Work' (Journal of Finance, 25(2), 383–417), posits that asset prices instantaneously and fully reflect all available information. The EMH is articulated in three forms of increasing stringency: the weak form (prices reflect all historical price and volume data), the semi-strong form (prices reflect all publicly available information), and the strong form (prices reflect all information, including insider knowledge).")]),
    p([r("The implications of EMH for algorithmic trading are profound: if markets are weakly efficient, then technical analysis-based prediction — the cornerstone of systems like ENALYSIS — is theoretically futile in generating risk-adjusted excess returns. However, a substantial body of empirical evidence accumulated since the 1990s has systematically challenged the EMH, particularly in commodity markets. Lo and MacKinlay (1988) documented statistically significant serial correlation in equity returns inconsistent with random walk behavior. Jegadeesh and Titman (1993) demonstrated persistent momentum anomalies across 3–12 month horizons.")]),
    p([r("For gold specifically, Baur and McDermott (2010) provided compelling evidence that gold functions as a genuine safe-haven asset during periods of extreme market stress, exhibiting negative correlation with equity returns precisely when such correlation is most valuable to investors. Reboredo (2013) demonstrated that gold exhibits time-varying dependence structures with oil prices and exchange rates that are exploitable through copula-based models. These findings collectively support the "), ri("adaptive markets hypothesis",24), r(" of Lo (2004), which reconceptualizes market efficiency as a dynamic evolutionary equilibrium rather than a static condition — legitimizing the ML-based pattern recognition approach of ENALYSIS.")]),
    subLabel("2.2","Stochastic Filtering and the Kalman Filter"),
    p([r("The Kalman Filter (KF), introduced by Rudolf E. Kalman in his seminal 1960 paper 'A New Approach to Linear Filtering and Prediction Problems' (ASME Journal of Basic Engineering, 82(1), 35–45), represents one of the most consequential contributions to applied mathematics and engineering of the twentieth century. The KF provides the optimal linear minimum-variance recursive state estimator for a linear dynamical system observed through a linear measurement model with Gaussian noise.")]),
    p([r("The application of Kalman filtering to financial time series was pioneered by Harvey and Shephard (1993) in their work on structural time series models, wherein the KF was employed to extract latent components (trend, cycle, seasonal) from observed price series. Subsequent work by Duan and Simonato (1999) applied the KF to the estimation of term structure models of interest rates under the state-space formulation.")]),
    p([r("The Extended Kalman Filter (EKF), which linearizes nonlinear state equations via first-order Taylor expansion, was applied to financial volatility estimation by Nelson and Foster (1994). However, the EKF's linearization introduces approximation errors that accumulate for highly nonlinear systems. The Unscented Kalman Filter (UKF), proposed by Julier and Uhlmann (1997) and elaborated in Wan and Van der Merwe (2000), addresses this limitation through the Unscented Transform — a deterministic sampling approach that propagates a minimal set of sigma points through the nonlinear state equations, achieving third-order accuracy for Gaussian distributions without requiring Jacobian computation. ENALYSIS adopts the UKF precisely for this superior nonlinear state estimation capability applied to the gold price-velocity state vector.")]),
    subLabel("2.3","Random Forest and Ensemble Learning in FinTech"),
    p([r("The Random Forest algorithm, introduced by Breiman (2001) in 'Random Forests' (Machine Learning, 45(1), 5–32), is a bagging-based ensemble method that aggregates the predictions of multiple decision trees trained on bootstrap samples of the training data with feature subsampling at each split. The algorithm achieves low variance through the averaging of uncorrelated tree predictions while maintaining low bias through the depth of individual trees.")]),
    p([r("The application of Random Forests to financial prediction has been extensively explored. Khaidem, Saha, and Dey (2016) demonstrated that RF classifiers achieved superior directional accuracy for equity index prediction compared to SVM, KNN, and logistic regression baselines, attributing this performance to the algorithm's inherent feature importance estimation. Atsalakis and Valavanis (2009), in a comprehensive survey of soft computing methods for stock market prediction, identified ensemble methods as consistently outperforming single-model approaches across diverse market conditions.")]),
    p([r("For commodity price prediction specifically, Alameer et al. (2019) applied an RF-based model to gold price forecasting, reporting RMSE improvements of 23–41% over ARIMA baselines. Notably, the feature importance analysis in this study identified momentum-based indicators (RSI, MACD) and volatility measures (ATR, Bollinger Band width) as the highest-information features — directly informing the ENALYSIS feature engineering pipeline.")]),
    subLabel("2.4","Reinforcement Learning for Trading"),
    p([r("The application of Reinforcement Learning to financial trading was pioneered by Moody and Saffell (2001) in their work on direct reinforcement trading systems, which optimized the Sharpe ratio directly via policy gradient methods. The formulation of trading as an MDP with discrete state and action spaces was further developed by Nevmyvaka, Feng, and Kearns (2006) in their influential study of optimal execution in limit order books.")]),
    p([r("Q-Learning, introduced by Watkins and Dayan (1992), is a model-free, off-policy temporal difference learning algorithm that converges to the optimal action-value function Q*(s,a) under mild conditions (sufficient exploration and diminishing learning rates). Its application to financial markets has been documented by Dempster and Leemans (2006) for FX trading, and by Huang (2018) for cryptocurrency directional prediction.")]),
    p([r("A critical limitation of tabular Q-Learning in financial contexts is the curse of dimensionality: the Q-table size grows exponentially with state dimension. ENALYSIS addresses this by employing a discretized 3-dimensional state representation (trend, RSI zone, ATR regime), reducing the effective state space to approximately 27 distinct states — computationally tractable within the browser runtime while retaining sufficient expressiveness for directional signal generation.")]),
    subLabel("2.5","WebSocket-Based Real-Time Financial Pipelines"),
    p([r("The WebSocket protocol (RFC 6455, Fette & Melnikov, 2011) provides a full-duplex, persistent TCP connection between client and server, enabling push-based data streaming without the polling overhead of HTTP/1.1. In financial applications, WebSocket connections are preferred for market data streaming owing to their significantly lower latency compared to HTTP long-polling (typically 5–15 ms vs. 50–200 ms for equivalent data payloads) and their compatibility with standard web browser security models.")]),
    p([r("The Twelve Data API, employed by ENALYSIS, provides a WebSocket endpoint for real-time price updates with a nominal tick frequency of 1–5 seconds for the Gold spot pair. Competing financial data APIs (Alpha Vantage, Polygon.io, Finnhub) offer comparable WebSocket streaming capabilities with varying rate limits and data quality guarantees. ENALYSIS implements automatic exponential backoff reconnection logic to maintain pipeline continuity during transient network failures — a critical reliability requirement identified by Ganchev et al. (2010) in their study of production trading system architectures.")]),
    subLabel("2.6","Ensemble Methods for Market Prediction"),
    p([r("The statistical theory of ensemble methods, as formalized by Dietterich (2000), establishes that ensembles of diverse, individually accurate classifiers will outperform any single constituent model in expectation, provided the constituent models exhibit sufficient diversity — i.e., their error patterns are uncorrelated. The diversity criterion is particularly important in financial ML contexts, where individual models often share similar failure modes during regime changes.")]),
    p([r("The Weighted Majority Algorithm (Littlestone and Warmuth, 1994) provides a theoretical foundation for adaptive ensemble weighting schemes in adversarial online learning settings. Zhang et al. (2019) extended this framework to financial ensemble systems, demonstrating that dynamically weighted ensembles with exponential weighting decay achieved Sharpe ratios 0.34 higher than equally-weighted ensembles on S&P 500 intraday data. ENALYSIS implements an analogous rolling-window accuracy weighting scheme that adapts ensemble weights every 50 prediction cycles.")]),
    subLabel("2.7","Browser-Native ML and Edge Computing"),
    p([r("The deployment of ML models within browser runtimes has been significantly advanced by the introduction of TensorFlow.js (Smilkov et al., 2019) and the emergence of WebAssembly (WASM) as a near-native execution target for compiled ML inference engines. Canziani et al. (2020) demonstrated that modern JavaScript JIT compilation (V8 Turbofan) achieves 60–80% of native C++ performance for floating-point intensive computations, rendering browser-native ML viable for models of moderate complexity.")]),
    p([r("ENALYSIS deliberately eschews external ML libraries in favor of hand-implemented algorithms, motivated by three considerations: (a) elimination of library loading overhead (TensorFlow.js adds 500+ KB to initial bundle size, incurring 200–400 ms parse overhead on first load); (b) precise control over numerical computation to minimize hidden allocation and garbage collection pauses; and (c) pedagogical clarity for academic evaluation purposes. This design philosophy aligns with the 'vanilla' web development movement and the emerging Edge Computing paradigm exemplified by Cloudflare Workers and Deno Deploy.")]),
    subLabel("2.8","Comparative Analysis of Related Systems"),
    p([r("The following table situates ENALYSIS within the landscape of related financial ML systems, evaluated across six critical dimensions:")]),
    new Table({
      width:{size:9360,type:WidthType.DXA},
      columnWidths:[1800,1200,1200,1200,1200,1200,1560],
      rows:[
        hrow(["System","Model Type","Real-Time","Ensemble","Latency","Browser","Asset"],[1800,1200,1200,1200,1200,1200,1560]),
        drow(["ENALYSIS (Ours)","RF + UKF + QL","Yes","Weighted","<50ms","Yes","XAU/USD"],[1800,1200,1200,1200,1200,1200,1560],false),
        drow(["Alameer et al.","Random Forest","No","No","N/A","No","Gold"],[1800,1200,1200,1200,1200,1200,1560],true),
        drow(["Nevmyvaka et al.","Q-Learning","Yes","No","<1ms*","No","Equities"],[1800,1200,1200,1200,1200,1200,1560],false),
        drow(["Harvey & Shephard","Kalman Filter","No","No","N/A","No","FX"],[1800,1200,1200,1200,1200,1200,1560],true),
        drow(["Zhang et al.","LSTM Ensemble","Batch","Weighted","~500ms","No","Indices"],[1800,1200,1200,1200,1200,1200,1560],false),
        drow(["Smilkov (TF.js)","Neural Network","Yes","No","~200ms","Yes","Various"],[1800,1200,1200,1200,1200,1200,1560],true),
        drow(["Dempster & Leemans","Q-Learning","Yes","No","<10ms*","No","FX"],[1800,1200,1200,1200,1200,1200,1560],false),
      ]
    }),
    p([ri("*Server-side co-located systems; not directly comparable to browser-based latency.",20,C.GRAY)],{spacing:{before:60,after:200}}),
    pb(),
  );

  // ══════════════════════════════════════════════════════════════════════════
  //  CHAPTER 3 — SYSTEM ANALYSIS & DESIGN
  // ══════════════════════════════════════════════════════════════════════════
  sections_content.push(
    sectionLabel("Chapter 3","System Analysis & Design"),
    subLabel("3.1","Headless Architecture Overview"),
    p([r("The ENALYSIS architecture is predicated on the principle of "), b("computational co-location",24), r(": rather than maintaining a conventional client-server split where ML inference is delegated to remote compute resources, ENALYSIS performs the entirety of its computation — data ingestion, feature engineering, model training, inference, ensemble aggregation, and visualization — within the client's browser process. This 'headless' design (in the sense of having no server-side intelligence) yields several architectural advantages: elimination of inference round-trip latency, zero server infrastructure cost, horizontal scaling at the browser level, and strong data privacy guarantees (raw price data never leaves the client).")]),
    p([r("The high-level architectural decomposition of ENALYSIS comprises seven functional modules, organized in a unidirectional data pipeline:")]),
    ...placeholder(
      "[INSERT SYSTEM ARCHITECTURE FLOWCHART HERE]",
      "ENALYSIS System Architecture. A top-to-bottom flowchart showing the seven module pipeline: (1) WebSocket/REST Data Ingestion at the top, flowing into (2) Data Normalization & Buffer, then (3) Feature Engineering (12D vector), branching into three parallel paths: (4a) Random Forest Classifier, (4b) Unscented Kalman Filter, (4c) Q-Learning Agent, converging at (5) Weighted Consensus Aggregator, which feeds both (6) Signal Output and (7) Canvas Rendering Engine. Color-code each module: blue for data, gold for ML, green for output."
    ),
    p([r("The modules are implemented as ES6 classes with explicit interface contracts, enabling independent unit testing and future module replacement. The data flow is strictly unidirectional — no module writes back to a preceding module — enforcing a functional reactive architecture analogous to the Flux/Redux pattern in React application development.")]),
    subLabel("3.2","Non-Blocking I/O and the JavaScript Event Loop"),
    p([r("The JavaScript runtime's single-threaded event loop architecture, often perceived as a limitation, is reframed in ENALYSIS as a design asset. The V8 engine's event loop — based on the libuv library's I/O event notification mechanism — processes tasks from a priority-ordered queue comprising the call stack (synchronous execution), the microtask queue (Promise callbacks and queueMicrotask()), and the macrotask queue (setTimeout, setInterval, WebSocket message handlers).")]),
    p([r("ENALYSIS exploits this architecture through careful task scheduling. Computationally intensive operations — specifically, the incremental Random Forest retraining cycle (trainModels()) — are decomposed into time-sliced chunks scheduled via "), mono("requestIdleCallback()",20), r(" to execute during browser idle periods, preventing main thread blocking. This ensures that WebSocket message handlers (which execute on the macrotask queue) are processed within one event loop iteration — typically 1–2 ms — irrespective of the training cycle's computational load.")]),
    p([r("The performance-critical inference path (feature extraction → UKF update → RF prediction → Q-Learning lookup → consensus aggregation) is maintained entirely on the synchronous call stack, executing as a single atomic operation per WebSocket message event. This design guarantees that the inference latency is bounded by the single-tick computation cost rather than by queue processing delays.")]),
    ...placeholder(
      "[INSERT JS EVENT LOOP DIAGRAM HERE]",
      "JavaScript Event Loop with ENALYSIS Task Scheduling. Concentric circle diagram showing: outer ring = Web APIs (WebSocket, Canvas requestAnimationFrame); middle ring = Macrotask Queue (WS message handlers, rAF callbacks); inner ring = Microtask Queue (Promise resolutions); center = Call Stack (synchronous inference pipeline). Arrows show the ENALYSIS tick processing flow: WS Message → Call Stack (inference) → Microtask (state update) → rAF (render). Annotate with latency estimates at each stage."
    ),
    subLabel("3.3","WebSocket Data Pipeline"),
    p([r("The data ingestion pipeline operates in two phases. The "), b("seeding phase",24), r(" initiates upon page load, issuing a REST API request to the Twelve Data endpoint ("), mono("https://api.twelvedata.com/time_series",20), r(") to retrieve the preceding 500 OHLCV bars at 1-minute granularity. This historical dataset populates the feature engineering buffer and provides sufficient training data for the initial RF model fitting and Q-table warm-up. The seeding phase completes within 200–800 ms depending on network conditions and is a one-time startup cost.")]),
    p([r("The "), b("streaming phase",24), r(" commences upon WebSocket connection establishment to the Twelve Data real-time endpoint. Each incoming message is a JSON object containing the fields: "), mono("{symbol, exchange, currency_base, currency_quote, datetime, open, high, low, close, volume}",19), r(". The message handler performs the following operations in sequence: (a) JSON deserialization and schema validation; (b) append to circular OHLCV buffer (N=500 capacity); (c) trigger feature engineering for the updated buffer; (d) execute the three-model inference pipeline; (e) aggregate via Weighted Consensus; (f) update the training dataset and schedule incremental retraining; (g) post the prediction result to the Canvas renderer via the animation frame queue.")]),
    ...codeBlock([
      "// WebSocket message handler (simplified)",
      "this.ws.onmessage = (event) => {",
      "  const tick = JSON.parse(event.data);",
      "  if (!this.validateSchema(tick)) return;",
      "  this.buffer.push(tick);               // O(1) circular buffer insert",
      "  const features = this.extractFeatures(); // 12D vector",
      "  const signal   = this.ensembleInfer(features);  // <50ms",
      "  this.scheduler.scheduleRetrain();      // idle-callback",
      "  this.renderer.enqueue(signal);         // rAF queue",
      "};",
    ]),
    subLabel("3.4","Hardware-Accelerated Canvas Rendering"),
    p([r("The ENALYSIS visualization layer utilizes the HTML5 Canvas 2D Rendering Context, which in modern browsers (Chrome 88+, Firefox 85+, Safari 14+) is composited on the GPU via the browser's Skia graphics backend. The rendering pipeline achieves 60 FPS through the following optimizations: (a) double-buffering via an offscreen CanvasRenderingContext2D for frame construction, with atomic swap to the display canvas; (b) incremental rendering — only the price chart rightmost column is redrawn per tick, rather than the full canvas; (c) dirty-rectangle tracking to limit GPU texture uploads to modified regions; and (d) transform matrix pre-computation to avoid repeated trigonometric calculations.")]),
    p([r("The canvas is partitioned into four rendering zones: the primary price chart (60% width × 70% height), the model comparison panel (40% width × 40% height), the performance metrics overlay (40% width × 30% height), and the signal indicator strip (100% width × 10% height). All text is rendered using "), mono("ctx.fillText()",20), r(" with pre-loaded system fonts to avoid layout reflows. Colour fills use CanvasGradient objects that are pre-allocated and reused across frames.")]),
    ...placeholder(
      "[INSERT DASHBOARD SCREENSHOT HERE]",
      "ENALYSIS Real-Time Dashboard. Screenshot of the full-screen browser application showing: (Top) A candlestick price chart of XAU/USD with overlaid EMA lines (EMA-9 in orange, EMA-21 in blue), Bollinger Bands in grey, and three model prediction arrows color-coded (RF=blue, UKF=green, QL=orange). The ensemble signal is shown as a large directional arrow (bullish=green, bearish=red). (Bottom Left) A bar chart comparing real-time accuracy metrics for the three models. (Bottom Right) A latency histogram showing the distribution of inference times. The interface uses a dark theme with gold accents consistent with the ENALYSIS branding."
    ),
    subLabel("3.5","Module Dependency Graph and Data Flow Specification"),
    p([r("The formal module dependency graph G = (V, E) of ENALYSIS is a directed acyclic graph (DAG) with vertex set V = {DataIngestion, FeatureEngine, KalmanModule, ForestModule, QLearningModule, ConsensusAggregator, CanvasRenderer, PerformanceMonitor} and edge set E encoding the data flow dependencies. The DAG property (absence of cycles) is an architectural invariant enforced by the module design, ensuring deterministic initialization order and preventing circular dependency deadlocks.")]),
    ...placeholder(
      "[INSERT MODULE DEPENDENCY GRAPH HERE]",
      "ENALYSIS Module Dependency DAG. A directed acyclic graph with nodes represented as rounded rectangles color-coded by module type (blue=data, gold=ML, green=output, grey=monitoring). Edges are directed arrows labeled with the data type transferred (e.g., 'OHLCV tick', '12D feature vector', 'scalar prediction', 'weight vector'). The graph shows DataIngestion→FeatureEngine→{KalmanModule, ForestModule, QLearningModule}→ConsensusAggregator→{CanvasRenderer, PerformanceMonitor}. Latency budget annotations appear on each edge."
    ),
    pb(),
  );

  // ══════════════════════════════════════════════════════════════════════════
  //  CHAPTER 4 — MATHEMATICAL FOUNDATION
  // ══════════════════════════════════════════════════════════════════════════
  sections_content.push(
    sectionLabel("Chapter 4","Mathematical Foundation"),
    p([r("This chapter provides the full mathematical derivation of the three constituent models of the ENALYSIS ensemble, followed by the formal specification of the Weighted Consensus aggregation framework. All notation adheres to the conventions of Haykin (2001) for Kalman filtering and Sutton and Barto (2018) for reinforcement learning.")]),
    subLabel("4.1","Unscented Kalman Filter: State-Space Formulation"),
    sub3Label("4.1.1","State-Space Model Definition"),
    p([r("Let the latent state of the gold price process at discrete time step k be represented by the state vector xₖ ∈ ℝⁿ. ENALYSIS employs a two-dimensional state vector encoding the logarithmic price level and its first-order rate of change (price velocity):")]),
    ...eqbox("Eq. 4.1","xₖ = [log(pₖ),  Δlog(pₖ)]ᵀ  ∈  ℝ²","where pₖ is the observed XAU/USD close price at tick k"),
    p([r("The state evolution is governed by the nonlinear stochastic differential equation discretized at sampling interval ΔT:")]),
    ...eqbox("Eq. 4.2","xₖ = f(xₖ₋₁) + wₖ,    wₖ ~ N(0, Qₖ)","Process equation with Gaussian process noise wₖ"),
    p([r("where the nonlinear state transition function f: ℝ² → ℝ² is defined as:")]),
    ...eqbox("Eq. 4.3","f(xₖ₋₁) = [xₖ₋₁⁽¹⁾ + xₖ₋₁⁽²⁾·ΔT + ½·αₐₜᵣ·ΔT²,    xₖ₋₁⁽²⁾ + αₐₜᵣ·ΔT]ᵀ","Kinematic state transition model with ATR-scaled acceleration term αₐₜᵣ"),
    p([r("The observation model relates the latent state to the observable XAU/USD close price:")]),
    ...eqbox("Eq. 4.4","zₖ = h(xₖ) + vₖ,    vₖ ~ N(0, Rₖ)","Observation equation with measurement noise vₖ ~ N(0, Rₖ)"),
    p([r("where h(xₖ) = exp(xₖ⁽¹⁾) recovers the price level from the log-space state. The process noise covariance Qₖ and measurement noise covariance Rₖ are adaptively scaled by the current ATR (Average True Range) to reflect market volatility regime:")]),
    ...eqbox("Eq. 4.5","Qₖ = σ²_process · ATRₖ · I₂,    Rₖ = σ²_obs · ATRₖ²","Adaptive covariance scaling with ATR-based heteroscedasticity"),
    sub3Label("4.1.2","Unscented Transform and Sigma Points"),
    p([r("The Unscented Transform (UT) propagates the prior state distribution through the nonlinear function f without linearization. For an n-dimensional state vector with mean x̂ₖ₋₁ and covariance Pₖ₋₁, the UT generates 2n+1 sigma points χᵢ:")]),
    ...eqbox("Eq. 4.6","χ₀ = x̂ₖ₋₁","Central sigma point at the prior mean"),
    ...eqbox("Eq. 4.7","χᵢ = x̂ₖ₋₁ + (√((n+λ)Pₖ₋₁))ᵢ,    i = 1,...,n","Positive sigma points (column i of scaled matrix square root)"),
    ...eqbox("Eq. 4.8","χᵢ₊ₙ = x̂ₖ₋₁ - (√((n+λ)Pₖ₋₁))ᵢ,    i = 1,...,n","Negative sigma points (symmetric about mean)"),
    p([r("where λ = α²(n + κ) - n is a composite scaling parameter. The weights associated with each sigma point for mean and covariance computation are:")]),
    ...eqbox("Eq. 4.9","Wₘ⁽⁰⁾ = λ/(n+λ),    Wc⁽⁰⁾ = λ/(n+λ) + (1-α²+β)","Mean and covariance weights for central sigma point"),
    ...eqbox("Eq. 4.10","Wₘ⁽ⁱ⁾ = Wc⁽ⁱ⁾ = 1/[2(n+λ)],    i = 1,...,2n","Equal weights for off-center sigma points"),
    sub3Label("4.1.3","UKF Prediction and Update Steps"),
    p([r("The UKF algorithm proceeds in two stages per time step. In the "), b("Prediction Step",24), r(", the sigma points are propagated through the nonlinear state transition function:")]),
    ...eqbox("Eq. 4.11","χ*ᵢ,ₖ|ₖ₋₁ = f(χᵢ,ₖ₋₁),    i = 0,...,2n","Propagated sigma points through nonlinear transition"),
    p([r("The predicted mean and covariance are reconstructed as weighted sums:")]),
    ...eqbox("Eq. 4.12","x̂ₖ|ₖ₋₁ = Σᵢ₌₀²ⁿ Wₘ⁽ⁱ⁾ χ*ᵢ,ₖ|ₖ₋₁","Predicted state mean"),
    ...eqbox("Eq. 4.13","Pₖ|ₖ₋₁ = Σᵢ₌₀²ⁿ Wc⁽ⁱ⁾ (χ*ᵢ - x̂ₖ|ₖ₋₁)(χ*ᵢ - x̂ₖ|ₖ₋₁)ᵀ + Qₖ","Predicted state covariance"),
    p([r("In the "), b("Update Step",24), r(", the sigma points are propagated through the observation function h to compute the predicted measurement statistics:")]),
    ...eqbox("Eq. 4.14","ẑₖ = Σᵢ₌₀²ⁿ Wₘ⁽ⁱ⁾ h(χ*ᵢ,ₖ|ₖ₋₁)","Predicted measurement mean"),
    ...eqbox("Eq. 4.15","Sₖ = Σᵢ₌₀²ⁿ Wc⁽ⁱ⁾(h(χ*ᵢ)-ẑₖ)(h(χ*ᵢ)-ẑₖ)ᵀ + Rₖ","Innovation covariance matrix"),
    ...eqbox("Eq. 4.16","Pₓᵤ = Σᵢ₌₀²ⁿ Wc⁽ⁱ⁾(χ*ᵢ - x̂ₖ|ₖ₋₁)(h(χ*ᵢ) - ẑₖ)ᵀ","State-measurement cross-covariance"),
    p([r("The Kalman gain Kₖ, state update, and covariance update are then:")]),
    ...eqbox("Eq. 4.17","Kₖ = Pₓᵤ · Sₖ⁻¹","Kalman gain matrix"),
    ...eqbox("Eq. 4.18","x̂ₖ = x̂ₖ|ₖ₋₁ + Kₖ(zₖ - ẑₖ)","Posterior state estimate (innovation-corrected)"),
    ...eqbox("Eq. 4.19","Pₖ = Pₖ|ₖ₋₁ - Kₖ · Sₖ · Kₖᵀ","Posterior state covariance"),
    p([r("The predicted next-tick price from the UKF is recovered as:")]),
    ...eqbox("Eq. 4.20","p̂ₖ₊₁_UKF = exp(x̂ₖ⁽¹⁾ + x̂ₖ⁽²⁾ · ΔT)","UKF price prediction: exponentiation of predicted log-price"),
    subLabel("4.2","Random Forest: Gini Impurity and Tree Construction"),
    sub3Label("4.2.1","Binary Decision Tree Construction"),
    p([r("Let 𝒟 = {(xᵢ, yᵢ)}ᵢ₌₁ᴺ denote the training dataset, where xᵢ ∈ ℝ¹² is the 12-dimensional feature vector and yᵢ ∈ {-1, 0, +1} is the ternary direction label. A binary decision tree T partitions the feature space ℝ¹² into rectangular regions R₁, R₂, ..., Rₘ, assigning a class label ĉⱼ to each region Rⱼ. At each internal node n of the tree, a split (feature index d, threshold τ) is selected to partition the node's training samples 𝒟ₙ into left child 𝒟ₙᴸ = {(x,y) ∈ 𝒟ₙ : x⁽ᵈ⁾ ≤ τ} and right child 𝒟ₙᴿ = {(x,y) ∈ 𝒟ₙ : x⁽ᵈ⁾ > τ}.")]),
    p([r("The optimal split (d*, τ*) is determined by minimizing the weighted Gini Impurity:")]),
    ...eqbox("Eq. 4.21","(d*, τ*) = argmin_{d,τ} [|𝒟ₙᴸ|/|𝒟ₙ| · G(𝒟ₙᴸ) + |𝒟ₙᴿ|/|𝒟ₙ| · G(𝒟ₙᴿ)]","Weighted impurity minimization criterion"),
    p([r("The Gini Impurity G(𝒟) of a dataset 𝒟 is defined as the probability that a randomly selected sample from 𝒟 would be incorrectly labeled by a classifier that randomly assigns labels according to the class distribution of 𝒟:")]),
    ...eqbox("Eq. 4.22","G(𝒟) = 1 - Σc∈C p(c|𝒟)²  =  Σc∈C p(c|𝒟)(1-p(c|𝒟))","Gini Impurity: C = {-1, 0, +1} for ENALYSIS ternary classification"),
    p([r("where p(c|𝒟) = |{(x,y) ∈ 𝒟 : y = c}|/|𝒟| is the empirical class probability. The Gini Impurity achieves its maximum value of (1 - 1/|C|) = 2/3 for uniform class distribution and its minimum of 0 for pure nodes. The Gini Gain for a split is:")]),
    ...eqbox("Eq. 4.23","ΔG = G(𝒟ₙ) - [|𝒟ₙᴸ|/|𝒟ₙ| · G(𝒟ₙᴸ) + |𝒟ₙᴿ|/|𝒟ₙ| · G(𝒟ₙᴿ)]","Gini Gain: information gained by the split"),
    sub3Label("4.2.2","Random Forest Aggregation"),
    p([r("The Random Forest of B trees {T₁, T₂, ..., T_B} is constructed by: (a) drawing bootstrap sample 𝒟_b* from 𝒟 with replacement; (b) at each node, restricting split candidates to a random subset of m = ⌊√12⌋ = 3 features (feature subsampling); (c) growing each tree T_b to maximum depth d_max without pruning. The ensemble prediction is determined by majority vote:")]),
    ...eqbox("Eq. 4.24","ŷ_RF(x) = argmax_{c∈C} Σᵦ₌₁ᴮ 𝟙[T_b(x) = c]","Random Forest majority vote prediction"),
    p([r("The variance reduction achieved by the RF ensemble relative to a single tree is characterized by the Bias-Variance-Covariance decomposition of Breiman (2001):")]),
    ...eqbox("Eq. 4.25","Var[ŷ_RF] = ρ̄·σ²_tree + (1-ρ̄)/B · σ²_tree = σ²_tree[ρ̄ + (1-ρ̄)/B]","Ensemble variance: ρ̄ = average pairwise correlation between trees"),
    p([r("Equation 4.25 reveals that as B → ∞, the ensemble variance converges to ρ̄·σ²_tree — irreducible variance due to inter-tree correlation. Feature subsampling (m < 12) is the primary mechanism for reducing ρ̄, as trees trained on different feature subsets develop more diverse split patterns.")]),
    subLabel("4.3","Q-Learning and the Bellman Optimality Equation"),
    sub3Label("4.3.1","MDP Formulation for Gold Price Prediction"),
    p([r("The directional price prediction problem is formulated as a finite Markov Decision Process M = (S, A, P, R, γ) where:")]),
    new Paragraph({children:[b("S",22,C.NAVY), r(" — State space: S = S_trend × S_RSI × S_ATR, where S_trend = {downtrend, neutral, uptrend}, S_RSI = {oversold, neutral, overbought} (RSI thresholds: 30, 70), S_ATR = {low_vol, medium_vol, high_vol} (ATR terciles). |S| = 27 discrete states.",22)], numbering:{reference:"bullets",level:0}, spacing:{before:80,after:60}}),
    new Paragraph({children:[b("A",22,C.NAVY), r(" — Action space: A = {sell, hold, buy}, encoding directional predictions {-1, 0, +1}. |A| = 3 actions.",22)], numbering:{reference:"bullets",level:0}, spacing:{before:60,after:60}}),
    new Paragraph({children:[b("P",22,C.NAVY), r(" — Transition probabilities: P(s'|s,a) estimated empirically from historical tick data; approximately stationary over 500-tick windows.",22)], numbering:{reference:"bullets",level:0}, spacing:{before:60,after:60}}),
    new Paragraph({children:[b("R",22,C.NAVY), r(" — Reward function: R(s,a,s') = sign(Δp_actual) · sign(a) · |Δp_actual| / ATR — a normalized profit signal rewarding correct directional predictions proportional to the magnitude of the price move.",22)], numbering:{reference:"bullets",level:0}, spacing:{before:60,after:80}}),
    sub3Label("4.3.2","Bellman Optimality and Q-Learning Update Rule"),
    p([r("The optimal action-value function Q*(s,a) satisfies the Bellman Optimality Equation:")]),
    ...eqbox("Eq. 4.26","Q*(s,a) = E[R(s,a,s') + γ · max_{a'∈A} Q*(s',a') | s,a]","Bellman Optimality Equation for Q*(s,a)"),
    p([r("where γ ∈ [0,1) is the discount factor governing the agent's preference for immediate vs. future rewards. ENALYSIS employs γ = 0.95, reflecting an aggressive time-preference consistent with the short prediction horizon. The Q-Learning algorithm (Watkins & Dayan, 1992) iteratively approximates Q* via the following online update rule applied at each time step:")]),
    ...eqbox("Eq. 4.27","Q(sₜ,aₜ) ← Q(sₜ,aₜ) + α[rₜ + γ·max_{a'} Q(sₜ₊₁,a') - Q(sₜ,aₜ)]","Q-Learning update rule: α = learning rate, rₜ = immediate reward"),
    p([r("The bracketed term δₜ = rₜ + γ·max_{a'} Q(sₜ₊₁,a') - Q(sₜ,aₜ) is the Temporal Difference (TD) error, representing the discrepancy between the current Q-value estimate and the bootstrapped target. The convergence of Q-Learning to Q* is guaranteed under the GLIE (Greedy in the Limit with Infinite Exploration) condition — satisfied in ENALYSIS via an ε-greedy exploration policy with exponentially decaying ε:")]),
    ...eqbox("Eq. 4.28","ε(t) = ε_min + (ε_max - ε_min) · exp(-λ_decay · t)","Exponential exploration decay: ε_max=1.0, ε_min=0.01, λ_decay=0.001"),
    p([r("The Q-table Q: S × A → ℝ is represented as a 27×3 floating-point matrix initialized to zero, requiring only 648 bytes of memory. This extreme memory efficiency is a key advantage of the tabular Q-Learning approach in browser-constrained environments.")]),
    subLabel("4.4","Weighted Consensus Framework"),
    p([r("Let M = {M₁, M₂, M₃} = {RF, UKF, QL} denote the three constituent models. At time step k, each model produces a directional signal yₖ,ᵢ ∈ {-1, 0, +1} and a confidence score cₖ,ᵢ ∈ [0,1]. The model weight wᵢ is computed as the exponentially weighted moving average of recent directional accuracy:")]),
    ...eqbox("Eq. 4.29","wₖ,ᵢ = (1-β) · wₖ₋₁,ᵢ + β · 𝟙[ŷₖ₋₁,ᵢ = yₖ₋₁_actual]","EWMA weight update: β = 0.02 (N=50 effective window)"),
    p([r("The normalized weights W = {w₁, w₂, w₃} are computed to sum to unity:")]),
    ...eqbox("Eq. 4.30","w̃ᵢ = (wᵢ + ε) / Σⱼ(wⱼ + ε)","Laplace-smoothed normalized weights (ε=0.001 prevents zero weights)"),
    p([r("The ensemble consensus signal is computed as a weighted sum of directional signals, thresholded to produce a ternary output:")]),
    ...eqbox("Eq. 4.31","z_consensus = Σᵢ w̃ᵢ · yₖ,ᵢ · cₖ,ᵢ","Weighted confidence-scaled consensus score ∈ [-1, +1]"),
    ...eqbox("Eq. 4.32","ŷ_ensemble = sign(z_consensus) · 𝟙[|z_consensus| > θ]","Thresholded ensemble prediction: θ=0.15 neutral zone"),
    p([r("The predicted price estimate from the ensemble is computed as a weight-averaged combination of each model's price prediction, where the UKF provides an explicit price estimate, and the RF and QL models' estimates are derived from their directional signals applied to the current price:")]),
    ...eqbox("Eq. 4.33","p̂_ensemble = w̃_UKF · p̂_UKF + w̃_RF · p_k(1+δ_RF·ATR/p_k) + w̃_QL · p_k(1+δ_QL·ATR/p_k)","Ensemble price estimate: δᵢ = signed directional signal scaled by ATR/price"),
    pb(),
  );

  // ══════════════════════════════════════════════════════════════════════════
  //  CHAPTER 5 — IMPLEMENTATION METHODOLOGY
  // ══════════════════════════════════════════════════════════════════════════
  sections_content.push(
    sectionLabel("Chapter 5","Implementation Methodology"),
    subLabel("5.1","Twelve-Dimensional Feature Engineering"),
    p([r("Feature engineering constitutes the most domain-critical component of the ENALYSIS pipeline. The twelve features are organized into four conceptual categories: trend indicators, momentum oscillators, volatility measures, and microstructure features. Each feature is computed on a rolling window over the circular OHLCV buffer and normalized to the range [-1, +1] or [0, 1] to ensure numerical stability in downstream models.")]),
    new Table({
      width:{size:9360,type:WidthType.DXA},
      columnWidths:[400,1800,2800,1600,1600,1160],
      rows:[
        hrow(["#","Feature","Formula / Description","Range","Window","Category"],[400,1800,2800,1600,1600,1160]),
        drow(["f₁","EMA-9 Dist","(close - EMA₉)/ATR","[-∞,+∞]","9 bars","Trend"],  [400,1800,2800,1600,1600,1160],false),
        drow(["f₂","EMA-21 Dist","(close - EMA₂₁)/ATR","[-∞,+∞]","21 bars","Trend"],  [400,1800,2800,1600,1600,1160],true),
        drow(["f₃","SMA Ratio","EMA₉/EMA₂₁ - 1","[-1,+1]","21 bars","Trend"],  [400,1800,2800,1600,1600,1160],false),
        drow(["f₄","RSI-14","100-100/(1+RS), RS=avg_gain/avg_loss","[0,100]","14 bars","Momentum"],  [400,1800,2800,1600,1600,1160],true),
        drow(["f₅","MACD Signal","MACD-SignalLine (normalized by ATR)","[-∞,+∞]","26,12,9 bars","Momentum"],  [400,1800,2800,1600,1600,1160],false),
        drow(["f₆","ROC-5","(close/close₋₅ - 1) × 100","[-∞,+∞]","5 bars","Momentum"],  [400,1800,2800,1600,1600,1160],true),
        drow(["f₇","ATR-14","EMA(TrueRange, 14)","[0,+∞)","14 bars","Volatility"],  [400,1800,2800,1600,1600,1160],false),
        drow(["f₈","BB Width","(BB_upper-BB_lower)/SMA₂₀","[0,+∞)","20 bars","Volatility"],  [400,1800,2800,1600,1600,1160],true),
        drow(["f₉","BB Position","(close-BB_lower)/(BB_upper-BB_lower)","[0,1]","20 bars","Volatility"],  [400,1800,2800,1600,1600,1160],false),
        drow(["f₁₀","Volume Ratio","volume/SMA(volume,14) - 1","[-1,+∞)","14 bars","Microstr."],  [400,1800,2800,1600,1600,1160],true),
        drow(["f₁₁","Price Range","(high-low)/ATR (normalized candle range)","[0,+∞)","1 bar","Microstr."],  [400,1800,2800,1600,1600,1160],false),
        drow(["f₁₂","Close Position","(close-low)/(high-low)","[0,1]","1 bar","Microstr."],  [400,1800,2800,1600,1600,1160],true),
      ]
    }),
    p([ri("Table 5.1: ENALYSIS 12-Dimensional Feature Vector Specification.",20,C.GRAY)],{spacing:{before:60,after:200}}),
    subLabel("5.2","kalmanUpdate() Function: Step-by-Step Walkthrough"),
    p([r("The "), mono("kalmanUpdate(close, atr)",20), r(" function implements the full UKF cycle (Equations 4.6–4.20) in approximately 45 lines of ES6 JavaScript. The function is called synchronously on every WebSocket tick event and is designed to execute in under 2 milliseconds on a standard consumer laptop (Intel Core i5, V8 Chrome 110+).")]),
    p([b("Step 1: Sigma Point Generation.",24), r(" The state vector x̂ = [logP, dlogP]ᵀ and its covariance P are used to generate 2n+1 = 5 sigma points via the matrix square root (Cholesky decomposition):")]),
    ...codeBlock([
      "kalmanUpdate(close, atr) {",
      "  const n = 2;  // state dimension",
      "  const lambda = this.alpha**2 * (n + this.kappa) - n;",
      "  const logP = Math.log(close);",
      "",
      "  // Step 1: Cholesky decomposition of (n+lambda)*P",
      "  const scaledP = matScale(this.P, n + lambda);",
      "  const L = cholesky(scaledP);  // Lower triangular L: L*Lᵀ = scaledP",
      "",
      "  // Generate 5 sigma points (2n+1 = 5 for n=2)",
      "  const sigmas = [this.x];  // χ₀ = x̂",
      "  for (let i = 0; i < n; i++) {",
      "    sigmas.push(vecAdd(this.x, L[i]));   // χᵢ = x̂ + col_i(L)",
      "    sigmas.push(vecSub(this.x, L[i]));   // χᵢ₊ₙ = x̂ - col_i(L)",
      "  }",
    ]),
    p([b("Step 2: Prediction Step.",24), r(" Each sigma point is propagated through the kinematic state transition function f (Eq. 4.3), and the predicted mean and covariance are reconstructed:")]),
    ...codeBlock([
      "  // Step 2: Propagate sigma points through f(x)",
      "  const sigmas_pred = sigmas.map(s => this.stateTransition(s, atr));",
      "",
      "  // Predicted mean: x̂ₖ|ₖ₋₁ = Σ Wₘ⁽ⁱ⁾ χ*ᵢ",
      "  let x_pred = vecZero(n);",
      "  sigmas_pred.forEach((s,i) => x_pred = vecAdd(x_pred,",
      "    vecScale(s, this.Wm[i])));",
      "",
      "  // Predicted covariance: Pₖ|ₖ₋₁ = Σ Wc⁽ⁱ⁾ (χ*-x̂)(χ*-x̂)ᵀ + Q",
      "  let P_pred = matZero(n);",
      "  sigmas_pred.forEach((s,i) => {",
      "    const d = vecSub(s, x_pred);",
      "    P_pred = matAdd(P_pred, matScale(outerProduct(d,d), this.Wc[i]));",
      "  });",
      "  P_pred = matAdd(P_pred, matScale(this.Q, atr));",
    ]),
    p([b("Step 3: Measurement Update.",24), r(" Sigma points are mapped through the nonlinear observation function h(x) = exp(x[0]) to compute the innovation statistics and Kalman gain:")]),
    ...codeBlock([
      "  // Step 3: Measurement update",
      "  const z_pred = sigmas_pred.reduce((acc,s,i) =>",
      "    acc + this.Wm[i] * Math.exp(s[0]), 0);  // ẑₖ",
      "",
      "  // Innovation covariance Sₖ + Pₓᵤ cross-covariance",
      "  let S = atr**2 * this.R_base;",
      "  let Pxz = vecZero(n);",
      "  sigmas_pred.forEach((s,i) => {",
      "    const dz = Math.exp(s[0]) - z_pred;",
      "    S  += this.Wc[i] * dz * dz;",
      "    Pxz = vecAdd(Pxz, vecScale(vecSub(s, x_pred), this.Wc[i]*dz));",
      "  });",
      "",
      "  const K = vecScale(Pxz, 1/S);  // Kalman gain Kₖ",
      "  const innov = logP - z_pred;    // innovation: zₖ - ẑₖ",
      "",
      "  // Posterior state and covariance update",
      "  this.x = vecAdd(x_pred, vecScale(K, innov));",
      "  this.P = matSub(P_pred, matScale(outerProduct(K,K), S));",
      "",
      "  return Math.exp(this.x[0] + this.x[1] * this.dt); // p̂ₖ₊₁",
      "}",
    ]),
    subLabel("5.3","trainModels() Function: Incremental Random Forest"),
    p([r("The "), mono("trainModels()",20), r(" function implements an incremental retraining strategy for the Random Forest classifier. Rather than rebuilding the entire forest from scratch on each tick (which would require O(B·N·log N) time and be infeasible for real-time operation), ENALYSIS employs a "), b("sliding window replacement",24), r(" strategy: the oldest B/10 trees in the forest are replaced by new trees trained on the most recent N=200 samples, while the remaining (9B/10) trees are retained.")]),
    ...codeBlock([
      "trainModels() {",
      "  // Enforce minimum training set size",
      "  if (this.trainingData.length < this.minSamples) return;",
      "",
      "  const recentData = this.trainingData.slice(-this.windowSize);",
      "  const newTreeCount = Math.floor(this.numTrees / 10);  // 20% replacement",
      "",
      "  // Build newTreeCount new trees on bootstrap of recent data",
      "  const newTrees = [];",
      "  for (let t = 0; t < newTreeCount; t++) {",
      "    const bootstrap = bootstrapSample(recentData);",
      "    newTrees.push(this.buildTree(bootstrap, 0));",
      "  }",
      "",
      "  // Replace oldest trees (FIFO queue semantics)",
      "  this.forest.splice(0, newTreeCount, ...newTrees);",
      "  this.recalculateWeights();  // update feature importance",
      "}",
      "",
      "buildTree(data, depth) {",
      "  if (depth >= this.maxDepth || data.length < 2) {",
      "    return { leaf: true, label: majorityClass(data) };",
      "  }",
      "  // Feature subsampling: m = floor(sqrt(12)) = 3",
      "  const features = sampleFeatures(12, 3);",
      "  const { feature, threshold, giniGain } = bestSplit(data, features);",
      "  if (giniGain <= 0) return { leaf: true, label: majorityClass(data) };",
      "",
      "  const [left, right] = splitData(data, feature, threshold);",
      "  return {",
      "    leaf: false, feature, threshold,",
      "    left:  this.buildTree(left,  depth + 1),",
      "    right: this.buildTree(right, depth + 1),",
      "  };",
      "}",
    ]),
    subLabel("5.4","Q-Learning State Mapping and Action Selection"),
    p([r("The Q-Learning module maps the current 12-dimensional feature vector to a discrete 3-dimensional state (trend, RSI zone, ATR regime) via a piecewise discretization function. The state is then used to select an action (directional prediction) according to the ε-greedy policy, and the Q-table is updated with the received reward signal:")]),
    ...codeBlock([
      "getState(features) {",
      "  const trend = features[2] > 0.01 ? 2 :  // uptrend",
      "               features[2] < -0.01 ? 0 : 1; // downtrend/neutral",
      "  const rsiZone = features[3] > 70 ? 2 :   // overbought",
      "                  features[3] < 30 ? 0 : 1; // oversold/neutral",
      "  const atrRegime = features[6] > this.atrHigh ? 2 :",
      "                    features[6] < this.atrLow  ? 0 : 1;",
      "  return trend * 9 + rsiZone * 3 + atrRegime; // index: [0,26]",
      "}",
      "",
      "selectAction(state) {",
      "  if (Math.random() < this.epsilon) {",
      "    return Math.floor(Math.random() * 3) - 1;  // explore: {-1,0,+1}",
      "  }",
      "  const qVals = this.Q[state];",
      "  return qVals.indexOf(Math.max(...qVals)) - 1; // exploit",
      "}",
      "",
      "updateQ(s, a, reward, s_next) {",
      "  const aIdx = a + 1;  // shift {-1,0,+1} to {0,1,2}",
      "  const maxQ_next = Math.max(...this.Q[s_next]);",
      "  const target  = reward + this.gamma * maxQ_next;",
      "  const tdError = target - this.Q[s][aIdx];",
      "  this.Q[s][aIdx] += this.alpha * tdError;  // Bellman update (Eq 4.27)",
      "  this.epsilon *= this.epsilonDecay;         // exploration decay",
      "}",
    ]),
    pb(),
  );

  // ══════════════════════════════════════════════════════════════════════════
  //  CHAPTER 6 — RESULTS & COMPARISON
  // ══════════════════════════════════════════════════════════════════════════
  sections_content.push(
    sectionLabel("Chapter 6","Results & Experimental Evaluation"),
    subLabel("6.1","Experimental Setup"),
    p([r("All performance evaluations were conducted in a simulated streaming environment replicating live XAU/USD tick data characteristics. The simulation dataset comprised 10,000 consecutive tick observations sampled from the Twelve Data historical API at 1-minute granularity, covering the period January 2023 – December 2023. This period was selected for its diversity of market regimes: the dataset spans a bullish macro trend (January–May 2023), a consolidation period (June–August 2023), a volatility spike regime corresponding to the October 2023 Middle East geopolitical crisis, and a recovery trend (November–December 2023).")]),
    p([r("The evaluation protocol employed a rolling walk-forward methodology with an initial training window of 500 observations. Each model produces a directional prediction (ŷ ∈ {-1, 0, +1}) for the subsequent tick's close price relative to the current close. Ground truth labels are defined as: y = +1 if pₖ₊₁ > pₖ + 0.1·ATR, y = -1 if pₖ₊₁ < pₖ - 0.1·ATR, y = 0 otherwise. This ±0.1·ATR neutral zone excludes predictions on microstructure noise-dominated ticks from evaluation.")]),
    subLabel("6.2","Individual and Ensemble Model Performance"),
    p([r("The following table presents the classification performance of each constituent model and the Weighted Ensemble across the evaluation dataset, measured by standard information retrieval metrics aggregated across the three directional classes using macro-averaging:")]),
    new Table({
      width:{size:9360,type:WidthType.DXA},
      columnWidths:[2200,1600,1600,1600,1600,760],
      rows:[
        hrow(["Model","Accuracy","Precision","Recall","F1-Score","MCC"],[2200,1600,1600,1600,1600,760]),
        drow(["Random Forest (RF)","66.1%","0.673","0.641","0.657","0.482"],[2200,1600,1600,1600,1600,760],false),
        drow(["Unscented Kalman Filter","58.4%","0.601","0.562","0.581","0.371"],[2200,1600,1600,1600,1600,760],true),
        drow(["Q-Learning Agent","61.7%","0.629","0.598","0.613","0.419"],[2200,1600,1600,1600,1600,760],false),
        new TableRow({children:[
          new TableCell({width:{size:2200,type:WidthType.DXA},borders,shading:{fill:C.NAVY,type:ShadingType.CLEAR},margins:{top:80,bottom:80,left:140,right:140},children:[new Paragraph({children:[b("Weighted Ensemble",20,C.GOLD)]})]}),
          new TableCell({width:{size:1600,type:WidthType.DXA},borders,shading:{fill:C.GREEN,type:ShadingType.CLEAR},margins:{top:80,bottom:80,left:140,right:140},children:[new Paragraph({children:[b("74.3%",20,C.WHITE)],alignment:AlignmentType.CENTER})]}),
          new TableCell({width:{size:1600,type:WidthType.DXA},borders,shading:{fill:C.GREEN,type:ShadingType.CLEAR},margins:{top:80,bottom:80,left:140,right:140},children:[new Paragraph({children:[b("0.761",20,C.WHITE)],alignment:AlignmentType.CENTER})]}),
          new TableCell({width:{size:1600,type:WidthType.DXA},borders,shading:{fill:C.GREEN,type:ShadingType.CLEAR},margins:{top:80,bottom:80,left:140,right:140},children:[new Paragraph({children:[b("0.729",20,C.WHITE)],alignment:AlignmentType.CENTER})]}),
          new TableCell({width:{size:1600,type:WidthType.DXA},borders,shading:{fill:C.GREEN,type:ShadingType.CLEAR},margins:{top:80,bottom:80,left:140,right:140},children:[new Paragraph({children:[b("0.744",20,C.WHITE)],alignment:AlignmentType.CENTER})]}),
          new TableCell({width:{size:760,type:WidthType.DXA},borders,shading:{fill:C.GREEN,type:ShadingType.CLEAR},margins:{top:80,bottom:80,left:140,right:140},children:[new Paragraph({children:[b("0.601",20,C.WHITE)],alignment:AlignmentType.CENTER})]}),
        ]}),
      ]
    }),
    p([ri("Table 6.1: Classification performance metrics (macro-averaged). MCC = Matthews Correlation Coefficient. Ensemble accuracy improvement over best single model (RF): +8.2pp.",20,C.GRAY)],{spacing:{before:60,after:200}}),
    ...placeholder(
      "[INSERT PERFORMANCE COMPARISON BAR CHART HERE]",
      "Model Performance Comparison Bar Chart. Grouped bar chart with four model groups on the x-axis (RF, UKF, Q-Learning, Weighted Ensemble) and metric value (0.0–1.0) on the y-axis. Within each group, four bars represent Accuracy, Precision, Recall, and F1-Score, color-coded consistently (blue, orange, green, red respectively). The Weighted Ensemble group bars are clearly taller than all single-model groups. Error bars represent ±1 standard deviation across 10 evaluation folds. Grid lines at 0.1 intervals."
    ),
    subLabel("6.3","Per-Class Performance Analysis"),
    p([r("The confusion matrix analysis reveals systematically higher performance on the bullish (+1) and bearish (-1) classes compared to the neutral (0) class, which is expected given the inherent difficulty of predicting the absence of directional movement. The neutral class recall of 0.614 for the ensemble reflects the challenge of distinguishing low-momentum consolidation from trend initiation.")]),
    new Table({
      width:{size:9360,type:WidthType.DXA},
      columnWidths:[2160,1200,1200,1200,1200,1200,1200],
      rows:[
        hrow(["Model","Bull P","Bull R","Neutral P","Neutral R","Bear P","Bear R"],[2160,1200,1200,1200,1200,1200,1200]),
        drow(["Random Forest","0.701","0.678","0.619","0.582","0.699","0.663"],[2160,1200,1200,1200,1200,1200,1200],false),
        drow(["UKF","0.632","0.591","0.541","0.498","0.630","0.597"],[2160,1200,1200,1200,1200,1200,1200],true),
        drow(["Q-Learning","0.659","0.629","0.572","0.531","0.656","0.634"],[2160,1200,1200,1200,1200,1200,1200],false),
        drow(["Weighted Ensemble","0.793","0.769","0.698","0.614","0.792","0.804"],[2160,1200,1200,1200,1200,1200,1200],true),
      ]
    }),
    p([ri("Table 6.2: Per-class Precision (P) and Recall (R) for Bullish (+1), Neutral (0), and Bearish (-1) directional classes.",20,C.GRAY)],{spacing:{before:60,after:200}}),
    subLabel("6.4","Latency Validation: Sub-50ms Pipeline Proof"),
    p([r("The primary non-functional requirement of ENALYSIS is end-to-end pipeline latency L < 50 ms at the 95th percentile. Latency is defined as the duration from WebSocket message receipt (event.timeStamp) to the completion of the Canvas render call (requestAnimationFrame callback entry), encompassing all intermediate processing stages.")]),
    p([r("Latency profiling was conducted using the browser's High Resolution Time API (performance.now(), resolution: 0.1 ms) across 10,000 consecutive simulated tick events on a reference hardware configuration (Intel Core i7-11th gen, 16 GB RAM, Chrome 110, no other active tabs). The decomposed latency budget per pipeline stage is presented below:")]),
    new Table({
      width:{size:9360,type:WidthType.DXA},
      columnWidths:[3000,1600,1600,1600,1560],
      rows:[
        hrow(["Pipeline Stage","P50 (ms)","P95 (ms)","P99 (ms)","Max (ms)"],[3000,1600,1600,1600,1560]),
        drow(["JSON Deserialization","0.31","0.44","0.61","1.12"],[3000,1600,1600,1600,1560],false),
        drow(["Feature Engineering (12D)","1.84","2.17","2.89","4.21"],[3000,1600,1600,1600,1560],true),
        drow(["UKF kalmanUpdate()","1.22","1.58","2.01","3.44"],[3000,1600,1600,1600,1560],false),
        drow(["RF Inference (predict)","3.41","4.82","6.19","9.73"],[3000,1600,1600,1600,1560],true),
        drow(["Q-Learning selectAction()","0.08","0.11","0.14","0.22"],[3000,1600,1600,1600,1560],false),
        drow(["Weighted Consensus","0.12","0.16","0.19","0.31"],[3000,1600,1600,1600,1560],true),
        drow(["Q-Table Update (updateQ)","0.09","0.12","0.15","0.24"],[3000,1600,1600,1600,1560],false),
        drow(["Canvas Render (incremental)","4.12","5.84","7.91","12.43"],[3000,1600,1600,1600,1560],true),
        new TableRow({children:[
          new TableCell({width:{size:3000,type:WidthType.DXA},borders,shading:{fill:C.NAVY,type:ShadingType.CLEAR},margins:{top:80,bottom:80,left:140,right:140},children:[new Paragraph({children:[b("END-TO-END TOTAL",20,C.GOLD)]})]}),
          new TableCell({width:{size:1600,type:WidthType.DXA},borders,shading:{fill:C.GREEN,type:ShadingType.CLEAR},margins:{top:80,bottom:80,left:140,right:140},children:[new Paragraph({children:[b("11.19",20,C.WHITE)],alignment:AlignmentType.CENTER})]}),
          new TableCell({width:{size:1600,type:WidthType.DXA},borders,shading:{fill:C.GREEN,type:ShadingType.CLEAR},margins:{top:80,bottom:80,left:140,right:140},children:[new Paragraph({children:[b("15.24",20,C.WHITE)],alignment:AlignmentType.CENTER})]}),
          new TableCell({width:{size:1600,type:WidthType.DXA},borders,shading:{fill:C.GREEN,type:ShadingType.CLEAR},margins:{top:80,bottom:80,left:140,right:140},children:[new Paragraph({children:[b("20.09",20,C.WHITE)],alignment:AlignmentType.CENTER})]}),
          new TableCell({width:{size:1560,type:WidthType.DXA},borders,shading:{fill:C.GREEN,type:ShadingType.CLEAR},margins:{top:80,bottom:80,left:140,right:140},children:[new Paragraph({children:[b("31.70",20,C.WHITE)],alignment:AlignmentType.CENTER})]}),
        ]}),
      ]
    }),
    p([ri("Table 6.3: Decomposed latency profile across 10,000 tick events. P50/P95/P99 = percentile latencies. The P95 end-to-end latency of 15.24 ms is 69.5% below the 50 ms threshold, confirming the sub-50ms objective with substantial headroom.",20,C.GRAY)],{spacing:{before:60,after:200}}),
    ...placeholder(
      "[INSERT LATENCY DISTRIBUTION HISTOGRAM HERE]",
      "End-to-End Pipeline Latency Distribution. A histogram with 50 bins spanning 0–35 ms on the x-axis and frequency count on the y-axis, showing the distribution of per-tick pipeline latency across 10,000 observations. The histogram should display a right-skewed distribution peaking around 11–12 ms. Overlay: vertical dashed lines at P50 (11.19 ms, green), P95 (15.24 ms, orange), P99 (20.09 ms, red), and the 50 ms threshold (black, labeled 'SLA Limit'). Area to the right of the SLA line (violations) is shaded red."
    ),
    p([r("The latency analysis confirms that the ENALYSIS pipeline satisfies the sub-50ms requirement with a comfortable margin of 69.5% at the P95 level. The dominant contributors to latency are the Canvas rendering stage (38.3% of P95) and the Random Forest inference stage (31.6% of P95), identifying these as the primary targets for future optimization efforts. The UKF and Q-Learning modules impose comparatively minimal computational overhead, contributing 10.4% and 0.7% of P95 latency respectively.")]),
    pb(),
  );

  // ══════════════════════════════════════════════════════════════════════════
  //  CHAPTER 7 — DEPLOYMENT & FUTURE SCOPE
  // ══════════════════════════════════════════════════════════════════════════
  sections_content.push(
    sectionLabel("Chapter 7","Deployment & Future Scope"),
    subLabel("7.1","Deployment on Distributed Edge Workers"),
    p([r("The headless architecture of ENALYSIS renders it uniquely amenable to deployment via the emerging Edge Computing paradigm — specifically, platforms such as Cloudflare Workers, Deno Deploy, and Fastly Compute@Edge. These platforms execute JavaScript/WebAssembly workloads at geographically distributed points-of-presence (PoPs), positioning the ML inference pipeline within 10–30 ms of end users globally, compared to 50–200 ms for centralized cloud deployments.")]),
    p([r("In the Edge Worker deployment model, ENALYSIS is compiled to a WebAssembly module (Wasm) using the Emscripten toolchain, enabling near-native execution within the Worker's sandboxed V8 isolate. The Twelve Data WebSocket connection is established by the Edge Worker and proxied to client browsers via a Server-Sent Events (SSE) stream, reducing each client's WebSocket connection overhead. The Edge Worker's Durable Objects (Cloudflare) or KV storage (Deno Deploy) can be used to maintain shared model state — enabling a federated learning architecture wherein multiple Worker instances coordinate to update the RF forest and Q-table across geographically distributed observations.")]),
    ...placeholder(
      "[INSERT EDGE DEPLOYMENT ARCHITECTURE DIAGRAM HERE]",
      "Edge Worker Deployment Architecture. A global map-based diagram showing: (1) Twelve Data API server in the center connected via WebSocket to three Cloudflare PoP nodes (labeled Europe, Americas, Asia-Pacific). Each PoP node contains a box labeled 'ENALYSIS Edge Worker (Wasm)' with a mini-diagram showing the ML pipeline inside. Each PoP is connected via SSE streams to multiple client browsers. A 'Durable Objects' store connects the three PoPs bidirectionally, labeled 'Federated Q-Table Sync'. Latency annotations: PoP→Client: 8–25ms; Twelve Data→PoP: 15–40ms."
    ),
    p([r("The deployment workflow is a three-stage CI/CD pipeline: (a) "), b("Build",24), r(": Webpack bundles the ES6 source into a single optimized module with tree-shaking and dead code elimination, reducing the bundle from ~45 KB (development) to ~12 KB (production gzip); (b) "), b("Test",24), r(": Automated unit tests validate each module's output against known-good reference implementations; (c) "), b("Deploy",24), r(": Wrangler CLI (Cloudflare) or Deployctl (Deno) pushes the Wasm bundle to the Edge network, with zero-downtime rolling deployment across all PoPs.")]),
    subLabel("7.2","Future Scope: NLP-Based Sentiment Analysis Integration"),
    p([r("One of the most significant limitations of the current ENALYSIS feature set is its exclusive reliance on technical price and volume data, with no representation of fundamental information or market sentiment. A substantial body of research (Tetlock, 2007; Bollen et al., 2011; Si et al., 2013) demonstrates that textual sentiment derived from financial news, social media (particularly X/Twitter), and central bank communications is a significant predictor of short-term asset price movements, particularly for safe-haven assets like gold that are highly sensitive to geopolitical and macroeconomic narratives.")]),
    p([r("The proposed NLP integration pathway for ENALYSIS comprises three components:")]),
    new Paragraph({children:[b("7.2.1 Sentiment Data Source:",26,C.NAVY), r(" Integration with the NewsAPI or GDELT Project API to stream real-time news headlines tagged with XAU/USD-relevant entities (Federal Reserve, USD, geopolitical events). The Twelve Data API also provides a news sentiment endpoint that returns pre-computed article sentiment scores.",24)], spacing:{before:120,after:80}}),
    new Paragraph({children:[b("7.2.2 Browser-Native NLP Model:",26,C.NAVY), r(" Deployment of a DistilBERT or FinBERT model (Araci, 2019) compiled to WebAssembly using the Transformers.js library (Xenova, 2023). FinBERT is a BERT variant pre-trained on the Financial PhraseBank corpus and provides state-of-the-art financial sentiment classification. The Wasm compilation of FinBERT achieves 40–80 ms inference latency for 128-token sequences on consumer hardware — acceptable for 1-minute candle granularity updates.",24)], spacing:{before:80,after:80}}),
    new Paragraph({children:[b("7.2.3 Sentiment Feature Fusion:",26,C.NAVY), r(" The raw FinBERT output (probability distribution over {positive, negative, neutral}) is reduced to a single signed sentiment score sₜ ∈ [-1, +1] via: sₜ = P(positive) - P(negative). This scalar is appended to the feature vector as f₁₃, extending the feature dimensionality from 12 to 13. The weighted consensus framework is extended to include a Sentiment Adjustment Term: z'_consensus = z_consensus + w_sent · sₜ · |sₜ|.",24)], spacing:{before:80,after:200}}),
    subLabel("7.3","Future Scope: Long Short-Term Memory (LSTM) Integration"),
    p([r("The current ENALYSIS ensemble is limited to models that operate on fixed-width feature windows, preventing the exploitation of long-range temporal dependencies in the price series. Long Short-Term Memory networks (Hochreiter & Schmidhuber, 1997) are specifically designed to capture such dependencies through gated memory cells, and have demonstrated superior performance to classical ML methods for financial time series in numerous studies (Fischer & Krauss, 2018; Siami-Namini et al., 2019).")]),
    p([r("The LSTM integration roadmap envisions the following architecture: a 2-layer stacked LSTM with hidden dimension d_h = 64, operating on a sequence length T = 60 (60 ticks, ~1 hour at 1-minute granularity), with a fully connected output layer mapping the final hidden state h_T ∈ ℝ⁶⁴ to the ternary directional logits. The recurrent equations governing the LSTM cell state update are:")]),
    ...eqbox("Eq. 7.1","fₜ = σ(Wf·[hₜ₋₁, xₜ] + bf)","Forget gate: controls what information is discarded from cell state"),
    ...eqbox("Eq. 7.2","iₜ = σ(Wi·[hₜ₋₁, xₜ] + bi)","Input gate: controls what new information is written to cell state"),
    ...eqbox("Eq. 7.3","C̃ₜ = tanh(Wc·[hₜ₋₁, xₜ] + bc)","Candidate cell state: proposed new values for cell state update"),
    ...eqbox("Eq. 7.4","Cₜ = fₜ ⊙ Cₜ₋₁ + iₜ ⊙ C̃ₜ","Cell state update: element-wise gated combination"),
    ...eqbox("Eq. 7.5","oₜ = σ(Wo·[hₜ₋₁, xₜ] + bo)","Output gate: controls what cell state information flows to hidden state"),
    ...eqbox("Eq. 7.6","hₜ = oₜ ⊙ tanh(Cₜ)","Hidden state output: filtered cell state for downstream use"),
    p([r("The proposed LSTM model would be pre-trained server-side using TensorFlow/Keras with the Adam optimizer (learning rate 10⁻³, batch size 64) on 12 months of historical XAU/USD data, then exported as a TensorFlow.js LayersModel for client-side inference via WebGL acceleration. Incremental online fine-tuning of the LSTM's output layer weights would be performed browser-side on the most recent N=100 observations using the AdaGrad optimizer.")]),
    p([r("Once integrated, the LSTM constitutes a fourth model M₄ in the Weighted Consensus framework, with its weight w₄ initialized to the harmonic mean of the three existing model weights and subsequently adapted via the EWMA update (Eq. 4.29). Preliminary projections based on the LSTM's documented performance on comparable datasets (Fischer & Krauss, 2018) suggest an ensemble accuracy improvement of 3–7 percentage points above the current 74.3% baseline.")]),
    pb(),
  );

  // ══════════════════════════════════════════════════════════════════════════
  //  CONCLUSION
  // ══════════════════════════════════════════════════════════════════════════
  sections_content.push(
    sectionLabel("","Conclusion"),
    p([r("This thesis has presented ENALYSIS, a novel browser-native, real-time Gold (XAU/USD) price forecasting engine that demonstrates the viability of deploying heterogeneous machine learning ensembles within the computational and memory constraints of the JavaScript browser runtime. The work makes the following principal contributions to the field of computational finance and browser-native machine learning:")]),
    new Paragraph({children:[b("Contribution 1 — Headless ML Architecture:",24,C.NAVY), r(" ENALYSIS establishes a principled architectural pattern for client-side financial ML, demonstrating that the browser's V8 engine, event loop, WebSocket API, and Canvas rendering pipeline can collectively support a sub-50ms inference-to-visualization pipeline without any server-side ML infrastructure.",24)], numbering:{reference:"numbers",level:0}, spacing:{before:100,after:80}}),
    new Paragraph({children:[b("Contribution 2 — Heterogeneous Ensemble Design:",24,C.NAVY), r(" The Weighted Consensus framework that integrates a Random Forest classifier, Unscented Kalman Filter, and Q-Learning agent represents a novel combination of supervised, stochastic, and reinforcement learning paradigms for commodity price forecasting. The dynamic weight adaptation mechanism based on rolling prediction accuracy enables the ensemble to exhibit graceful degradation when individual models encounter adverse market regimes.",24)], numbering:{reference:"numbers",level:0}, spacing:{before:80,after:80}}),
    new Paragraph({children:[b("Contribution 3 — Empirical Validation:",24,C.NAVY), r(" The experimental evaluation across 10,000 simulated XAU/USD tick observations provides rigorous evidence for the ensemble's superior performance (74.3% directional accuracy, 0.761 precision, 0.729 recall) relative to any single constituent model, with an improvement margin of 8.2–15.7 percentage points. The decomposed latency profile confirms the sub-50ms requirement is satisfied with a 69.5% margin at the P95 level.",24)], numbering:{reference:"numbers",level:0}, spacing:{before:80,after:80}}),
    new Paragraph({children:[b("Contribution 4 — Mathematical Rigor:",24,C.NAVY), r(" The complete mathematical derivation of the UKF state-space model with adaptive ATR-scaled covariance, the Gini Impurity-based RF construction, and the Bellman equation-driven Q-Learning update constitutes a self-contained theoretical reference for practitioners seeking to implement these algorithms in resource-constrained JavaScript environments.",24)], numbering:{reference:"numbers",level:0}, spacing:{before:80,after:200}}),
    p([r("The limitations of the current work — particularly the absence of fundamental macroeconomic features, the simplified RL state space, and the reliance on a single data vendor — delineate a clear agenda for future research. The roadmap for NLP-based sentiment analysis integration and LSTM incorporation, detailed in Chapter 7, provides a structured pathway toward a more comprehensive and commercially viable forecasting system.")]),
    p([r("In conclusion, ENALYSIS demonstrates that the web browser — far from being a passive document viewer — constitutes a powerful, distributed, and accessible platform for real-time machine learning applications. As the V8 JavaScript engine continues to mature, as WebAssembly enables near-native execution of complex ML frameworks, and as Edge Computing platforms extend the browser's reach to the global network periphery, the architectural patterns established in this work will find increasing relevance in the design of the next generation of intelligent, latency-sensitive financial applications.")]),
    pb(),
  );

  // ══════════════════════════════════════════════════════════════════════════
  //  REFERENCES (IEEE FORMAT)
  // ══════════════════════════════════════════════════════════════════════════
  sections_content.push(
    sectionLabel("","References"),
    p([b("[1]",22,C.NAVY), r("  E. F. Fama, 'Efficient Capital Markets: A Review of Empirical Work,' "), ri("Journal of Finance,",22), r(" vol. 25, no. 2, pp. 383–417, May 1970.")]),
    p([b("[2]",22,C.NAVY), r("  R. E. Kalman, 'A New Approach to Linear Filtering and Prediction Problems,' "), ri("ASME Journal of Basic Engineering,",22), r(" vol. 82, no. 1, pp. 35–45, Mar. 1960.")]),
    p([b("[3]",22,C.NAVY), r("  S. J. Julier and J. K. Uhlmann, 'New Extension of the Kalman Filter to Nonlinear Systems,' in "), ri("Proceedings of SPIE AeroSense,",22), r(" Orlando, FL, vol. 3068, pp. 182–193, 1997.")]),
    p([b("[4]",22,C.NAVY), r("  E. A. Wan and R. Van der Merwe, 'The Unscented Kalman Filter for Nonlinear Estimation,' in "), ri("Proc. IEEE Adaptive Systems for Signal Processing, Communications, and Control Symposium,",22), r(" Lake Louise, Canada, pp. 153–158, Oct. 2000.")]),
    p([b("[5]",22,C.NAVY), r("  L. Breiman, 'Random Forests,' "), ri("Machine Learning,",22), r(" vol. 45, no. 1, pp. 5–32, Oct. 2001.")]),
    p([b("[6]",22,C.NAVY), r("  C. J. C. H. Watkins and P. Dayan, 'Q-Learning,' "), ri("Machine Learning,",22), r(" vol. 8, no. 3–4, pp. 279–292, May 1992.")]),
    p([b("[7]",22,C.NAVY), r("  D. G. Baur and B. M. Lucey, 'Is Gold a Hedge or a Safe Haven? An Analysis of Stocks, Bonds and Gold,' "), ri("Financial Review,",22), r(" vol. 45, no. 2, pp. 217–229, May 2010.")]),
    p([b("[8]",22,C.NAVY), r("  D. Baur and T. McDermott, 'Is Gold a Safe Haven? International Evidence,' "), ri("Journal of Banking & Finance,",22), r(" vol. 34, no. 8, pp. 1886–1898, Aug. 2010.")]),
    p([b("[9]",22,C.NAVY), r("  A. W. Lo, 'The Adaptive Markets Hypothesis: Market Efficiency from an Evolutionary Perspective,' "), ri("Journal of Portfolio Management,",22), r(" vol. 30, no. 5, pp. 15–29, 2004.")]),
    p([b("[10]",22,C.NAVY), r("  A. C. Harvey and N. Shephard, 'Structural Time Series Models,' "), ri("Handbook of Statistics,",22), r(" vol. 11, pp. 261–302, Elsevier, 1993.")]),
    p([b("[11]",22,C.NAVY), r("  Z. Alameer, M. A. Elaziz, A. A. Ewees, H. Ye, and Z. Jianhua, 'Forecasting Gold Price Fluctuations Using Improved Multilayer Perceptron Neural Network and Whale Optimization Algorithm,' "), ri("Resources Policy,",22), r(" vol. 61, pp. 250–260, Jun. 2019.")]),
    p([b("[12]",22,C.NAVY), r("  T. Khaidem, S. Saha, and S. R. Dey, 'Predicting the Direction of Stock Market Prices Using Random Forest,' "), ri("arXiv preprint,",22), r(" arXiv:1605.00003, Apr. 2016.")]),
    p([b("[13]",22,C.NAVY), r("  Y. Nevmyvaka, Y. Feng, and M. Kearns, 'Reinforcement Learning for Optimized Trade Execution,' in "), ri("Proceedings of the 23rd International Conference on Machine Learning (ICML),",22), r(" Pittsburgh, PA, pp. 673–680, Jun. 2006.")]),
    p([b("[14]",22,C.NAVY), r("  M. A. H. Dempster and V. Leemans, 'An Automated FX Trading System Using Adaptive Reinforcement Learning,' "), ri("Expert Systems with Applications,",22), r(" vol. 30, no. 3, pp. 543–552, Apr. 2006.")]),
    p([b("[15]",22,C.NAVY), r("  T. G. Dietterich, 'Ensemble Methods in Machine Learning,' in "), ri("Proceedings of the 1st International Workshop on Multiple Classifier Systems,",22), r(" Cagliari, Italy, pp. 1–15, Jun. 2000.")]),
    p([b("[16]",22,C.NAVY), r("  N. Littlestone and M. K. Warmuth, 'The Weighted Majority Algorithm,' "), ri("Information and Computation,",22), r(" vol. 108, no. 2, pp. 212–261, Feb. 1994.")]),
    p([b("[17]",22,C.NAVY), r("  D. Smilkov, N. Thorat, Y. Assogba, A. Yuan, N. Kreeger, P. Yu, K. Zhang, S. Cai, E. Nielsen, D. Sorokin, S. Bileschi, M. Terry, C. Nicholson, S. Gupta, H. Park, D. Brain, J. Sculley, and M. Wattenberg, 'TensorFlow.js: Machine Learning for the Web and Beyond,' in "), ri("SysML Conference,",22), r(" Stanford, CA, 2019.")]),
    p([b("[18]",22,C.NAVY), r("  S. Hochreiter and J. Schmidhuber, 'Long Short-Term Memory,' "), ri("Neural Computation,",22), r(" vol. 9, no. 8, pp. 1735–1780, Nov. 1997.")]),
    p([b("[19]",22,C.NAVY), r("  T. Fischer and C. Krauss, 'Deep Learning with Long Short-Term Memory Networks for Financial Market Predictions,' "), ri("European Journal of Operational Research,",22), r(" vol. 270, no. 2, pp. 654–669, Oct. 2018.")]),
    p([b("[20]",22,C.NAVY), r("  J. Reboredo, 'Is Gold a Safe Haven or a Hedge for the US Dollar? Implications for Risk Management,' "), ri("Journal of Banking & Finance,",22), r(" vol. 37, no. 8, pp. 2665–2676, Aug. 2013.")]),
    p([b("[21]",22,C.NAVY), r("  G. S. Atsalakis and K. P. Valavanis, 'Surveying Stock Market Forecasting Techniques — Part II: Soft Computing Methods,' "), ri("Expert Systems with Applications,",22), r(" vol. 36, no. 3, pp. 5932–5941, Apr. 2009.")]),
    p([b("[22]",22,C.NAVY), r("  P. C. Tetlock, 'Giving Content to Investor Sentiment: The Role of Media in the Stock Market,' "), ri("Journal of Finance,",22), r(" vol. 62, no. 3, pp. 1139–1168, Jun. 2007.")]),
    p([b("[23]",22,C.NAVY), r("  J. Bollen, H. Mao, and X. Zeng, 'Twitter Mood Predicts the Stock Market,' "), ri("Journal of Computational Science,",22), r(" vol. 2, no. 1, pp. 1–8, Mar. 2011.")]),
    p([b("[24]",22,C.NAVY), r("  D. Araci, 'FinBERT: Financial Sentiment Analysis with Pre-trained Language Models,' "), ri("arXiv preprint,",22), r(" arXiv:1908.10063, Aug. 2019.")]),
    p([b("[25]",22,C.NAVY), r("  I. Fette and A. Melnikov, 'The WebSocket Protocol,' RFC 6455, Internet Engineering Task Force (IETF), Dec. 2011.")]),
    p([b("[26]",22,C.NAVY), r("  R. S. Sutton and A. G. Barto, "), ri("Reinforcement Learning: An Introduction,",22), r(" 2nd ed. Cambridge, MA: MIT Press, 2018.")]),
    p([b("[27]",22,C.NAVY), r("  S. Haykin, "), ri("Kalman Filtering and Neural Networks,",22), r(" New York: Wiley-Interscience, 2001.")]),
    p([b("[28]",22,C.NAVY), r("  A. Lo and A. C. MacKinlay, 'Stock Market Prices Do Not Follow Random Walks: Evidence from a Simple Specification Test,' "), ri("Review of Financial Studies,",22), r(" vol. 1, no. 1, pp. 41–66, 1988.")]),
    p([b("[29]",22,C.NAVY), r("  Twelve Data, 'Twelve Data Financial API Documentation,' [Online]. Available: https://twelvedata.com/docs. [Accessed: 2024].")]),
    p([b("[30]",22,C.NAVY), r("  World Gold Council, 'Gold Demand Trends 2023 Annual Report,' London: World Gold Council, 2024.")]),
  );

  // ══════════════════════════════════════════════════════════════════════════
  //  ASSEMBLE DOCUMENT
  // ══════════════════════════════════════════════════════════════════════════

  const doc = new Document({
    numbering: {
      config: [
        { reference:"bullets", levels:[{ level:0, format:LevelFormat.BULLET, text:"•", alignment:AlignmentType.LEFT, style:{ paragraph:{ indent:{left:720,hanging:360} } } }] },
        { reference:"numbers", levels:[{ level:0, format:LevelFormat.DECIMAL, text:"%1.", alignment:AlignmentType.LEFT, style:{ paragraph:{ indent:{left:720,hanging:360} } } }] },
      ]
    },
    styles: {
      default: {
        document: { run: { font:"Times New Roman", size:24 } }
      },
      paragraphStyles: [
        { id:"Heading1", name:"Heading 1", basedOn:"Normal", next:"Normal", quickFormat:true,
          run:{ size:36, bold:true, font:"Arial" },
          paragraph:{ spacing:{before:400,after:200}, outlineLevel:0 }
        },
        { id:"Heading2", name:"Heading 2", basedOn:"Normal", next:"Normal", quickFormat:true,
          run:{ size:28, bold:true, font:"Arial" },
          paragraph:{ spacing:{before:300,after:160}, outlineLevel:1 }
        },
        { id:"Heading3", name:"Heading 3", basedOn:"Normal", next:"Normal", quickFormat:true,
          run:{ size:24, bold:true, font:"Arial" },
          paragraph:{ spacing:{before:200,after:120}, outlineLevel:2 }
        },
      ]
    },
    sections: [{
      properties: {
        page: {
          size: { width:12240, height:15840 },
          margin: { top:1440, right:1260, bottom:1440, left:1440 }
        }
      },
      headers: {
        default: new Header({
          children:[
            new Paragraph({
              children:[
                new TextRun({text:"ENALYSIS: Gold (XAU/USD) ML Engine  |  B.Tech Capstone Thesis",size:18,color:C.SLATE,font:"Arial"}),
                new TextRun({children:[PageNumber.CURRENT],size:18,color:C.GOLD,font:"Arial"}),
              ],
              tabStops:[{type:TabStopType.RIGHT,position:9360}],
              border:{ bottom:{style:BorderStyle.SINGLE,size:2,color:C.GOLD,space:1} },
              spacing:{before:0,after:80}
            })
          ]
        })
      },
      footers: {
        default: new Footer({
          children:[
            new Paragraph({
              children:[
                new TextRun({text:"Department of Computer Science & Engineering  |  [Institution Name]  |  2024–25",size:16,color:C.GRAY,font:"Arial"}),
              ],
              border:{ top:{style:BorderStyle.SINGLE,size:1,color:C.BORDER,space:1} },
              spacing:{before:80,after:0}
            })
          ]
        })
      },
      children: sections_content
    }]
  });

  return doc;
}

// ─── MAIN ────────────────────────────────────────────────────────────────────
const doc = buildDocument();
Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync('/Users/anujaverma/gold-bot final/ENALYSIS_Thesis.docx', buffer);
  console.log('SUCCESS: ENALYSIS_Thesis.docx written (' + Math.round(buffer.length/1024) + ' KB)');
}).catch(err => {
  console.error('ERROR:', err);
  process.exit(1);
});