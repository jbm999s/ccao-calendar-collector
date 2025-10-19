<img width="1536" height="1024" alt="calendar collector" src="https://github.com/user-attachments/assets/178f5707-1b3b-4478-95eb-ce037766602b" />

# CCAO Calendar Collector
Tracking and Visualizing Cook County Assessment Deadlines  
**By [JustinMcClelland.com](https://www.justinmcclelland.com)**

---

## Overview

**CCAO Calendar Collector** is a lightweight Python utility that retrieves official publication and appeal deadline data from the [Cook County Assessor’s Assessment Calendar](https://www.cookcountyassessor.com/assessment-calendar-and-deadlines) and saves it as a clean Excel workbook.  
Just run it locally and get structured data, ready for analysis.

---

## Quick Start

```bash
git clone https://github.com/jbm999s/ccao-calendar-collector.git
cd ccao-calendar-collector
python3 -m venv venv
source venv/bin/activate  # (Windows: venv\Scripts\activate)
pip install -r requirements.txt
python ccao_calendar_collector.py

