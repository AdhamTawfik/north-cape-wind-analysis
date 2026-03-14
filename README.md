# Wind Speed Data Analysis – North Cape, PEI

Analysis of hourly wind speed data from the North Cape meteorological station (Environment Canada) using Python.

## Dataset
- **Source:** Environment Canada Historical Climate Data
- **Station:** North Cape, Prince Edward Island (Station ID: 8300516)
- **Period:** January 2025 – February 2026
- **Resolution:** Hourly

## Features
- Combines and cleans 14 months of raw hourly CSV data
- Removes missing values and sorts by timestamp
- Computes key statistics: average, max, min, standard deviation, % calm hours, and % of hours above turbine cut-in speed (14 km/h)
- Monthly average breakdown across the full dataset period
- Exports a formatted Excel report with three sheets: Cleaned Data, Summary Statistics, and Monthly Averages
- Generates three charts saved as PNG files:
  - Wind speed distribution histogram
  - Average monthly wind speed bar chart
  - Wind speed time series

## Output Files
| File | Description |
|------|-------------|
| `wind_analysis_report.xlsx` | Excel report with cleaned data, stats, and monthly averages |
| `wind_speed_histogram.png` | Distribution of wind speeds |
| `monthly_avg_wind.png` | Average wind speed by month |
| `wind_speed_timeseries.png` | Wind speed over time |

## Requirements
```
pandas
matplotlib
openpyxl
```
Install with: `pip install pandas matplotlib openpyxl`

## Usage
```
python wind_data_reader.py
```
Place all monthly CSV files in the `data_csv/` folder before running.
