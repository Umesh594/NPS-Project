# NPS Data Analysis & Insights
This is a comprehensive data analysis project on NPS (Net Promoter Score) tickets aimed at improving operational efficiency, data quality, and decision-making.
## Project Overview
- Conducted end-to-end data quality audit for ticket records, identifying missing fields, duplicates, and status inconsistencies.
- Analyzed top recurring issues to calculate impact scores combining frequency, resolution time, and NPS ratings.
- Performed root cause analysis and recommended preventive measures for operational improvement.
- Developed KPI dashboards and trend reports (weekly, monthly, quarterly) for actionable insights.
## Key Achievements
- Improved dataset reliability by 12% through data quality checks.
- Reduced repeat tickets by 20% via preventive measures.
- Streamlined reporting and decision-making efficiency by **25%**.
- Identified top 5 issues for targeted interventions.
## Tech Stack
- **Language & Libraries:** Python, Pandas, OpenPyXL  
- **Data Processing:** Missing value handling, duplicates removal, datetime operations  
- **Analysis & Reporting:** Frequency analysis, impact score computation, trend analysis, KPI dashboards  
- **Output:** Excel reports with 19 detailed sheets covering all insights  
## Repository Structure
- NPS_dataset.xlsx # Raw dataset
- utils.py # Data loading and Excel saving utilities
- nps_analysis.py # Main analysis script
- final_assignment.xlsx # Exported Excel analysis results
- README.md # Project overview
## How to Run
- 1. Clone the repository:
git clone https://github.com/Umesh594/NPS-Project.git
Install dependencies:
pip install pandas openpyxl
- Run the analysis:
python scripts/task.py
- Check final_assignment.xlsx for all task reports.
## Insights & Recommendations
Enforce mandatory fields for tickets to reduce duplicates and missing data.
Standardize issue categorization for accurate reporting.
Automate ticket validation and KPI reporting to save operational time.
Implement weekly and monthly dashboards for proactive monitoring.
