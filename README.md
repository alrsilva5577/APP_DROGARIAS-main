# 💊 APP_Drogarias (Smart Sanitary Inspection System)

An enterprise-grade, cloud-native dynamic inspection and reporting application built for the **Health Surveillance Agency (Vigilância Sanitária)**. This system replaces cumbersome static forms with an intelligent, data-driven workflow tailored for drugstores and pharmacy regulations.

## 🚀 The Real-World Problem
Sanitary inspectors traditionally rely on extensive, monolithic forms (like Google Forms) to conduct on-site audits. This creates severe friction:
* **Overwhelming Navigation:** Inspectors spend valuable time scrolling through irrelevant sections.
* **Double Work:** Findings must be manually transcribed into official Word (`.docx`) reports post-inspection.
* **Data Silos:** Compliance history, infrastructure metrics, and pending legal requirements are stored across isolated files, making data tracking impossible.

## ✨ Solution & Key Features
**VisaInspecon** turns the inspection process into an agile, modular, and interconnected web ecosystem accessible from any tablet or mobile device in the field.

* **Dynamic Inspection Form:** Features conditional logic rendering. Selecting non-compliance flags automatically generates targeted observation boxes and stores corresponding regulatory infraction keys.
* **Automated Official Document Engine:** Generates legally-compliant, structurally-sound `.docx` inspection reports on the fly with custom string-formatting parameters.
* **Centralized Cloud Architecture:** Uses synchronized cloud data pipelines to logging historical records, making metrics instantly available across different users/inspectors.
* **Automated Pending Requirements Sheet:** Analyzes the database to extract the last inspection for a specific license request, cross-references infraction tables, and outputs an itemized official legal requirement document.

## 🛠️ Tech Stack & Architecture
* **Frontend/Interface:** Python with [Streamlit](https://streamlit.io) for high-performance reactive web rendering.
* **Data Pipeline & Analytics:** [Pandas](https://pydata.org) for loading, indexing, and processing multi-table regulatory databases.
* **Document Automation:** `python-docx` for background XML manipulation and dynamic text injection.
* **Cloud Storage & Database:** Cloud Data integration mimicking relational schemas (`INSPECOES_DB` & `INFRACOES_DB`) for seamless multiple-user concurrency.

## 📁 Repository Structure
```text
├── .streamlit/
│   └── secrets.toml          # Encrypted Cloud Database credentials
├── app.py                     # Main application flow and reactive routing engine
├── INFRACOES_DB.xlsx          # Normalized multi-column legal regulations database
└── requirements.txt           # Declared system dependencies
```

## ⚙️ Core Engineering Design (How it works under the hood)
The engine utilizes a **Template String / Key-Value Mapping design**. Legal frameworks (e.g., RDC 44/2009, Portaria 344/98) are normalized in a database. 

When an inspector marks an item as **"Non-Compliant" (NC)**:
1. The app flags the primary key (e.g., `RDC44_3`).
2. The reporting engine extracts the matching `phrase_non_compliant` from the dataset.
3. The server pushes the infraction key into the tracking database to enable targeted follow-up actions (such as generating enforcement sheets).

## 🚀 Getting Started

### Prerequisites
* Python 3.10 or higher installed.

### Installation
1. Clone the repository:
   ```bash
   git clone https://github.com
   cd VisaInspecon-Drogarias
   ```

2. Install the production dependencies:
   ```bash
   pip install -r requirements.txt
   ```

3. Configure your local environment secrets in `.streamlit/secrets.toml`.

4. Run the local development server:
   ```bash
   streamlit run app.py
   ```

---
*Developed as a highly scalable architecture to modernize public health monitoring infrastructure.*

