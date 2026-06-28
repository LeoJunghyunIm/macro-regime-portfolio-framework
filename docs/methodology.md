# Methodology

## Research objective

The framework converts macro and market conditions into transparent
multi-asset portfolio stances. It is intended to discipline the reasoning
process: identify the active risk channels, classify the policy and inflation
backdrop, and translate those states into positioning across equity,
duration, cash, and TIPS.

## Frequency and alignment

All series are converted to week-ending Friday observations. The workflow
uses the latest completed Friday, takes the last available observation within
each week, and forward-fills after weekly resampling. Signal calculations are
then performed on the aligned weekly dataset.

Two common transformations are used:

- **Rolling P80:** the 80th percentile over the latest 52 weekly observations.
- **Four-week change:** current weekly value minus the value four weekly
  observations earlier.

## Risk Appetite Engine

### Credit stress

High Yield OAS is flagged when either:

- the spread is above its rolling P80, or
- its four-week change is positive.

The reason code distinguishes `Level`, `Momentum`, `Both`, and `None`.

### Financial conditions

NFCI is flagged only when its level is above rolling P80. A momentum rule is
not used because the index already summarizes broad financial conditions and
the public framework favors a compact, non-duplicative score.

### Volatility

VIX is flagged only when its level is above rolling P80. Low volatility does
not subtract from the score or create an additional risk-on signal.

### Labor-market stress

Initial Claims is flagged only when its level is above rolling P80. A weekly
momentum trigger is excluded because claims data can be noisy.

### Score and classification

The four binary flags are summed without fractional weights:

```text
Risk Appetite Score = Credit + NFCI + VIX + Initial Claims
```

| Score | Classification |
|---:|---|
| 0 | Risk-On |
| 1-2 | Neutral |
| 3-4 | Risk-Off |

Risk Appetite Trend compares the current score with the previous valid week.
An unchanged score is marked deteriorating when the credit reason code still
contains momentum.

## Rates & Inflation Engine

### Rate Pressure

The 10-year Treasury yield activates Rate Pressure when its level is above
rolling P80 or its four-week change is positive.

### Inflation Concern

The 10-year breakeven inflation rate activates Inflation Concern when its
level is above rolling P80 or its four-week change is positive.

### Two-by-two regime map

| Rate Pressure | Inflation Concern | Regime |
|---|---|---|
| ON | ON | Regime 1 |
| ON | OFF | Regime 2 |
| OFF | ON | Regime 3 |
| OFF | OFF | Regime 4 |

### Real-Yield Constraint

The 10-year real yield activates Real Yield Tightening when its level is above
rolling P80 or its four-week change is positive. It is not a separate regime
axis. Instead, it constrains the Portfolio Router by capping equity and
duration overweight stances at Neutral.

The 2-year Treasury yield is retained as potential curve context but does not
enter the current public decision rules.

## Portfolio Router

The router combines three Risk Appetite classifications with four Rates &
Inflation regimes. The matrix in `config.yaml` produces one of three stances
for each asset class:

- Overweight
- Neutral
- Underweight

The Portfolio Router produces qualitative directional stances across equity,
nominal Treasury duration, cash, and TIPS or inflation protection. These
stances are not optimized portfolio weights or allocation bands.

## Data quality

Pre-export checks cover:

- missing latest observations
- source freshness
- sufficient rolling history
- aligned as-of dates
- availability of Portfolio Router outputs
- validity of the latest regime signals

Data Confidence is High when all checks pass, Medium when warnings exist
without a failure, and Low when any required check fails. Data Confidence
measures output reliability, not investment conviction.

## Historical diagnostics

### Historical Regime Recognition

Selected historical windows are used to evaluate whether the framework
recognized relevant stress conditions during known historical episodes. The
diagnostic reports recognition status, recognition timing, peak Risk Appetite
Score, and peak Rates & Inflation regime.

Recognition timing of zero may reflect stress that was already active at the
beginning of the selected event window. Recognition does not imply prediction,
and the diagnostic does not evaluate portfolio returns.

### Regime Co-occurrence

The co-occurrence matrix reports the historical share of observations in each
Risk Appetite and Rates & Inflation combination. It helps distinguish broad
financial stress from rate/inflation pressure and shows why similar risk
scores can have different portfolio implications.

## Interpretation boundaries

The framework is primarily a coincident regime monitor and portfolio
decision-support tool. It does not estimate expected returns, optimize a
portfolio, predict the timing of market shocks or specific turning points, or
claim causal identification.

At each weekly date, signal calculations use no observations carrying dates
after that date. Rolling windows are trailing and uncentered. This
computational treatment does not constitute a complete real-time or
point-in-time data reconstruction.

Historical diagnostics use currently available FRED histories rather than a
complete real-time-vintage database. Publication lags and subsequent data
revisions may therefore affect historical classifications.

The framework is not a complete business-cycle classifier. Its scope is risk
appetite, financial conditions, rates, inflation compensation, real yields,
and their portfolio implications. Thresholds are transparent and intentionally
compact, but rolling rules may lag abrupt turning points and selected
historical diagnostic windows involve judgment.
