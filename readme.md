# 🏏 Cricket Centuries Analysis - Automated PPT Generation 📊

## Overview

This project showcases my ability to automate the generation of PowerPoint presentations from web-scraped data. I've developed a Python-based pipeline that extracts cricket statistics, processes it, and creates visually appealing PPTs with data visualizations.

## Key Skills Demonstrated

* **Web Scraping:** Data extraction from websites.
* **Data Manipulation:** Pandas for data cleaning and transformation.
* **PPT Automation:** `python-pptx` for dynamic PPT creation.
* **Data Visualization:** `Matplotlib` for generating charts.
* **Interactive Elements:** HTML and Plotly for interactive graphs (Note: Implementation details in separate HTML files).
* **Object-Oriented Programming:** Code organization into classes for better structure.
* **Workflow Automation:** End-to-end automation from data to presentation.

## Project Details

### 1. Data Acquisition

* I used web scraping techniques to gather data on cricket players and their century records from various online sources.
* The data includes player details (name, date of birth, place of birth, family information), career statistics (total centuries, centuries per year), and match information.

### 2. Data Processing

* I organized the scraped data into Pandas DataFrames.
* Data cleaning and transformation were performed to handle missing values and ensure data consistency.
* The DataFrames were stored as Excel files for intermediate persistence.

### 3. PPT Generation

* I developed a Python script that uses the `python-pptx` library to automate the creation of PowerPoint presentations.
* The script reads data from the Excel files and generates slides for each player, including:
  * Player information
  * Century statistics
  * Tables summarizing key data
* I used `Matplotlib` to create visualizations of century trends over the years (bar and line plots).

### 4. Interactive Visualizations (HTML/Plotly)

* To enhance interactivity, I created HTML pages with Plotly graphs.
* These graphs provide dynamic exploration of the data (e.g., zooming, tooltips).
* *(Note: The PPTs themselves are static. The interactive graphs are in separate HTML files, and the intended integration would involve linking from the PPT or exporting the PPT to HTML.)*

### 5. Code Structure

The project code is organized as follows:

* `runner.ipynb`:  Notebook to orchestrate the PPT creation process.
* `prepare_data.ipynb`:  Notebook to orchestrate the data gathering and web scraping process.
* `ppt_generator.py` (Class):  Handles data transformation and preparation.
* `custom_presentation.py` (Class):  Manages PPT styling, slide creation, and content population.

## PPT Analysis Summary

I've created two PPTs: one for male players ("player\_Male.pptx") and one for female players ("player\_Female.pptx").

## Technical Considerations

* Error handling is implemented in the web scraping and data processing stages.
* The code is designed to be modular and maintainable.
* Further improvements can include:
  * Dynamic data updates
  * External configuration
  * More granular modularity
  * Logging
  * Unit testing

## Conclusion

This project demonstrates my ability to combine various technical skills to create an automated data-to-presentation workflow. I am eager to apply and expand these skills in a professional setting.

---
