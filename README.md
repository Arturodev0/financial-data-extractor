# Financial Data Extractor & Analyzer

A robust Python ETL (Extract, Transform, Load) tool designed to process financial Excel reports. It automates the extraction of specific financial indicators (like "Income Statement" or "COGS"), cleans the data, and generates granular CSV reports for further analysis.

##  Features

- **Automated Extraction:** Reads large Excel datasets and filters by specific fiscal years.
- **Data Cleaning:** Handles date formatting and missing values automatically.
- **Granular Reporting:** - Generates a high-level summary grouped by Parent/Class.
  - Generates a "Zoom-in" detailed report for specific cost centers (e.g., COGS) broken down by source.
- **Configurable:** Fully agnostic design. You can define column names, target years, and file paths in the configuration section without touching the logic.

## 📋 Prerequisites

You need **Python 3.x** installed. The project relies on pandas for data manipulation and openpyxl for Excel reading.

## 🛠️ Installation

1. Clone this repository:
   ```bash
   git clone [https://github.com/your-username/financial-data-extractor.git](https://github.com/your-username/financial-data-extractor.git)
