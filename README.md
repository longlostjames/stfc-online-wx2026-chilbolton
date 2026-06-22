# Chilbolton Climatology — Work Expericne Project

Jupyter notebook worksheets and introduction materials for the
**Chilbolton Atmospheric Observatory Climatology** project,
part of the STFC Online WX2026.

## Contents

| File | Purpose |
|---|---|
| `intro_chilbolton.ipynb` | Observatory overview: location, instruments, dataset summary |
| `intro_netcdf.ipynb` | Introduction to NetCDF files and the Python `netCDF4` library |
| `intro_statistics.ipynb` | Statistical methods: PDFs, exceedance curves, regression, R² |
| `student1_rainfall.ipynb` | Worksheet: Rainfall |
| `student2_temperature_humidity.ipynb` | Worksheet: Temperature & Relative Humidity |
| `student3_pressure.ipynb` | Worksheet: Atmospheric Pressure |
| `student4_wind.ipynb` | Worksheet: Wind Speed & Direction |

## Getting Started

1. Read through the introduction notebooks in order:
   `intro_chilbolton` → `intro_netcdf` → `intro_statistics`
2. Open your assigned student worksheet (`student1_rainfall.ipynb` etc.)
3. Run the **Setup** cell first — it defines all helper functions
4. Work through the tasks in order; each builds on the previous one

## Data

Precipitation data are stored as NetCDF files in:
```
/datawexp/cwalden/precipitation/
```
The data are already configured in the setup cells of each notebook.

## Requirements

```
netCDF4
numpy
pandas
matplotlib
scipy
cftime
```

