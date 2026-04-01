# set-esg-analyzer
AI-Powered ESG Screening System for SET50 Using Python + Claude API + CFA Framework
# AI-Powered ESG Screening System for SET50

Senior Project — KBS KMITL
Bachelor of Economics (Business Economics and Management)

## Overview
Automated ESG analysis system for SET50 stocks
using Python, Claude API, and CFA Framework

## Features
- ESG Score analysis (E/S/G 0-100)
- DCF Valuation for each stock
- Investment Decision Matrix
- Supports 10+ SET50 stocks

## Tech Stack
- Python, yfinance, pandas, matplotlib
- Anthropic Claude API (claude-sonnet-4-5)
- Google Colab

## Results
| Stock  | ESG Score | DCF Valuation | Recommendation |
|--------|-----------|---------------|----------------|
| KBANK  | 75        | Undervalued   | BUY            |
| CPALL  | 75        | Fairly Valued | BUY            |
| SCB    | 75        | Fairly Valued | BUY            |
| PTT    | 68        | Fairly Valued | HOLD           |

## Key Finding
KBANK is the only Sweet Spot stock with
high ESG score AND undervalued DCF

## Files
- esg_analysis.py — Main analysis code
- esg_ranking.png — ESG Ranking chart
- investment_matrix.png — Decision Matrix
