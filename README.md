# Portuguese Tourism Time Series Analysis and Forecasting

A comprehensive time series analysis and forecasting project for monthly tourist arrivals in Portugal's specialized accommodation sector using advanced statistical modeling techniques.

## 📊 Project Overview

This project analyzes monthly guest arrivals in Portugal's specialized accommodation sector from 2017-2023, developing robust forecasting models to support tourism industry planning and decision-making. The analysis successfully navigates the challenges of COVID-19 impact while providing actionable insights for stakeholders.

### Key Highlights

- **Dataset**: 84 monthly observations (2017-2023) from Portugal's National Institute of Statistics
- **Methodology**: Box & Jenkins SARIMA modeling with comprehensive model comparison
- **Best Model**: SARIMA(1,0,1)(1,1,0)[12] with superior statistical diagnostics
- **Forecast Accuracy**: 4.4% projected growth for 2024 with uncertainty quantification
- **Impact Analysis**: Detailed COVID-19 disruption assessment and recovery patterns

## 🎯 Objectives

1. **Analyze seasonal patterns** and structural trends in Portuguese tourism data
2. **Develop robust forecasting models** using SARIMA, ETS, and GARCH methodologies
3. **Assess COVID-19 impact** and recovery trajectories in tourism demand
4. **Generate actionable forecasts** with uncertainty quantification for 2024-2026
5. **Provide scenario analysis** for strategic planning purposes

## 📈 Key Findings

### Statistical Results
- **Seasonal Patterns**: Strong seasonality with 3.4:1 peak-to-trough ratio (August/January)
- **COVID-19 Impact**: 69.1% decline in 2020, partial recovery by 2021 (-29.1% vs 2019)
- **Model Performance**: SARIMA model achieves excellent residual diagnostics (Ljung-Box p=0.806)
- **Forecast Horizon**: 2024 projection of 33.9 million guests (+4.4% vs 2023)

### Business Insights
- **Peak Season**: August remains the strongest month (5.2M guests projected)
- **Recovery Trajectory**: Moderate growth expected with high uncertainty
- **Planning Scenarios**: Conservative (-30.6%), Expected (+4.4%), Optimistic (+39.5%)
- **Risk Assessment**: Wide confidence intervals emphasize need for flexible capacity planning

## 📚 References

1. **Shumway, R.H. & Stoffer, D.S.** (2017). *Time Series Analysis and Its Applications: With R Examples* (4th ed.). Springer.

2. **Hyndman, R.J. & Athanasopoulos, G.** (2021). *Forecasting: Principles and Practice* (3rd ed.). OTexts.

3. **Cryer, J.D. & Chan, K.-S.** (2008). *Time Series Analysis: With Applications in R* (2nd ed.). Springer.

4. **Instituto Nacional de Estatística** (2024). *Tourism Statistics - Accommodation*. Retrieved from https://www.ine.pt/

5. **Box, G.E.P. & Jenkins, G.M.** (1970). *Time Series Analysis: Forecasting and Control*. Holden-Day.
