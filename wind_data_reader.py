import pandas as pd
import matplotlib.pyplot as plt
import glob

# Load and combine all monthly CSV files
files = glob.glob("data_csv/wind_data_*.csv") + glob.glob("data_csv/wind_speed_*.csv")
df = pd.concat([pd.read_csv(f) for f in sorted(files)], ignore_index=True)

# Data Cleaning
df['Date/Time (LST)'] = pd.to_datetime(df['Date/Time (LST)'])
df = df.dropna(subset=['Wind Spd (km/h)'])
df = df.sort_values('Date/Time (LST)').reset_index(drop=True)

# Statistics
avg_wind = df['Wind Spd (km/h)'].mean()
max_wind = df['Wind Spd (km/h)'].max()
min_wind = df['Wind Spd (km/h)'].min()
std_wind = df['Wind Spd (km/h)'].std()
calm_pct = (df['Wind Spd (km/h)'] == 0).sum() / len(df) * 100
above_cutin_pct = (df['Wind Spd (km/h)'] >= 14).sum() / len(df) * 100

summary = pd.DataFrame({
    "Metric": ["Average Wind Speed (km/h)", "Max Wind Speed (km/h)",
               "Min Wind Speed (km/h)", "Std Deviation (km/h)",
               "% Calm Hours", "% Hours Above Cut-in Speed (14 km/h)"],
    "Value": [avg_wind, max_wind, min_wind, std_wind, calm_pct, above_cutin_pct]
})

monthly = df.groupby(df['Date/Time (LST)'].dt.to_period('M'))['Wind Spd (km/h)'].mean().reset_index()
monthly.columns = ['Month', 'Avg Wind Speed (km/h)']
monthly['Month'] = monthly['Month'].astype(str)

# Visualization
plt.figure()
plt.hist(df['Wind Spd (km/h)'], bins=20, edgecolor='black')
plt.title('Wind Speed Distribution - North Cape')
plt.xlabel('Wind Speed (km/h)')
plt.ylabel('Frequency')
plt.savefig('wind_speed_histogram.png')
plt.close()

plt.figure()
plt.bar(monthly['Month'], monthly['Avg Wind Speed (km/h)'])
plt.title('Average Monthly Wind Speed - North Cape')
plt.xlabel('Month')
plt.ylabel('Avg Wind Speed (km/h)')
plt.savefig('monthly_avg_wind.png')
plt.close()

plt.figure()
plt.plot(df['Date/Time (LST)'], df['Wind Spd (km/h)'], linewidth=0.5)
plt.title('Wind Speed Over Time - North Cape')
plt.xlabel('Date')
plt.ylabel('Wind Speed (km/h)')
plt.savefig('wind_speed_timeseries.png')
plt.close()

# Save processed data + summary to Excel
with pd.ExcelWriter("wind_analysis_report.xlsx", engine="openpyxl") as writer:
    df.to_excel(writer, sheet_name="Cleaned Data", index=False)
    summary.to_excel(writer, sheet_name="Summary Statistics", index=False)
    monthly.to_excel(writer, sheet_name="Monthly Averages", index=False)

    for sheet_name in writer.sheets:
        worksheet = writer.sheets[sheet_name]
        for col in worksheet.columns:
            max_length = max(len(str(cell.value)) if cell.value else 0 for cell in col)
            worksheet.column_dimensions[col[0].column_letter].width = max_length + 4

print("Excel report generated.")
               

