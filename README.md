# Who's Who Document Generator

This repository hosts a Streamlit application designed to generate professional presentations from structured data. The application merges and cleans data from various sources, applying a pre-defined template to produce a tailored presentation.

## Features

- **Interactive Web Interface:** Utilizes Streamlit for an easy-to-navigate GUI.
- **Data Integration:** Combines data from different Excel or CSV files.
- **Automated Presentation Creation:** Generates presentations using a predefined PowerPoint template.
- **Downloadable Outputs:** Allows users to download the final presentation in PowerPoint format.

## Prerequisites

Ensure you have Python installed on your system. The application requires the following Python packages:

```bash
pip install streamlit pandas python-pptx
```
## Installation
Clone the repository:
```bash
git clone https://github.com/your-username/whos-who-document-generator.git
```
## Navigate to the project directory:
```bash
cd whos-who-document-generator
```
## Usage
Run the Streamlit application using the following command:

```bash
streamlit run app.py
```

Follow the steps in the application's GUI to upload the required data files and generate the presentation.

## Application Workflow

Upload Job Nomenclature: Upload an Excel or CSV file containing job nomenclature data.
Upload Odoo Extraction: Upload an Excel or CSV file with data extracted from Odoo.
Generate Presentation: Click the button to merge, clean, and transform the data into a PowerPoint presentation based on a template.

## Contributing
Contributions to this project are welcome. Please fork the repository, make your changes, and submit a pull request.
