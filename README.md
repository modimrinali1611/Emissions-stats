#### Emissions Stats Analysis

R script to calculate sustained statistics for chemical species from emissions data.

**Sustained emissions** are defined as periods during which the measured concentration of a species remains relatively stable within a specified tolerance.
In this script:
- Stable periods are identified per operational configuration and water wash stage.
- A stable period must last a **minimum duration** (default 3 hours) or a percentage of the total time in that block (default 10%).
- Concentrations within a stable period must not deviate by more than a **percent tolerance** (default 10%) from the mean of that period.
- These metrics help identify consistent emission trends, removing short-term fluctuations that could bias analysis.
- For each species, we calculate:
  - **Sustained_Min** – the lowest mean concentration across stable periods
  - **Sustained_Max** – the highest mean concentration across stable periods
  - **Sustained_Avg** – the average of the means across all stable periods


#### Usage

1. Load your emissions data from Excel.
2. Define the species list.
3. Run the functions `get_ww_stage_block_summary()` to get summary statistics.

#### Dependencies

- dplyr
- readxl
- openxlsx
- lubridate
- stringr
- forcats
- viridis
- purrr
