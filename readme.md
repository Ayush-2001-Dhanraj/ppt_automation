<a id="readme-top"></a>
<div align="center">

[![Web Scraping](https://img.shields.io/badge/Web_Scraping-informational?style=for-the-badge&logoColor=white)](https://en.wikipedia.org/wiki/Web_scraping)
[![Pandas](https://img.shields.io/badge/Pandas-150458?style=for-the-badge&logo=pandas&logoColor=white)](https://pandas.pydata.org/)
[![python-pptx](https://img.shields.io/badge/python--pptx-informational?style=for-the-badge&logoColor=white)](https://python-pptx.readthedocs.io/en/latest/)
[![Matplotlib](https://img.shields.io/badge/Matplotlib-%23F37B7D.svg?style=for-the-badge&logo=Matplotlib&logoColor=white)](https://matplotlib.org/)
[![HTML5](https://img.shields.io/badge/HTML5-%23E34F26.svg?style=for-the-badge&logo=html5&logoColor=white)](https://developer.mozilla.org/en-US/docs/Web/HTML)
[![Plotly](https://img.shields.io/badge/Plotly-2C3E50?style=for-the-badge&logo=plotly&logoColor=white)](https://plotly.com/)
[![Python](https://img.shields.io/badge/Python-3776AB?style=for-the-badge&logo=python&logoColor=white)](https://www.python.org/)
[![OOP](https://img.shields.io/badge/OOP-informational?style=for-the-badge&logoColor=white)](https://en.wikipedia.org/wiki/Object-oriented_programming)
[![Automation](https://img.shields.io/badge/Automation-informational?style=for-the-badge&logoColor=white)](https://en.wikipedia.org/wiki/Automation)

<br />

[![Email](https://img.shields.io/badge/dhanrajaayush123%40gmail.com-important?style=for-the-badge)](mailto:dhanrajaayush123@gmail.com)
[![Email](https://img.shields.io/badge/ayushdhanraj.work%40gmail.com-important?style=for-the-badge)](mailto:ayushdhanraj.work@gmail.com)
[![LinkedIn](https://img.shields.io/badge/LinkedIn-0077B5?style=for-the-badge&logo=linkedin&logoColor=white)](https://www.linkedin.com/in/ayush-d-1759461a1)
[![GitHub](https://img.shields.io/badge/GitHub-100000?style=for-the-badge&logo=github&logoColor=white)](https://github.com/Ayush-2001-Dhanraj)

</div>

<br />
<div align="center">
  <img src="assets/logo.png" alt="Project Logo" width="150">

  <h3 align="center">🏏 Cricket Centuries Analysis - Automated PPT Generation 📊</h3>

  <p align="center">
    A project demonstrating the automated generation of PowerPoint presentations from web-scraped cricket data, showcasing data analysis and visualization skills.
    <br />
    <a href="https://github.com/Ayush-2001-Dhanraj/ppt_automation/blob/main/readme.md"><strong>View on GitHub</strong></a>
  </p>
</div>

<details>
  <summary>Table of Contents</summary>
  <ol>
    <li><a href="#overview">Overview</a></li>
    <li><a href="#key-skills-demonstrated">Key Skills Demonstrated</a></li>
    <li><a href="#project-details">Project Details</a>
      <ol>
        <li><a href="#1-data-acquisition">Data Acquisition</a></li>
        <li><a href="#2-data-processing">Data Processing</a></li>
        <li><a href="#3-ppt-generation">PPT Generation</a></li>
        <li><a href="#4-interactive-visualizations-htmlplotly">Interactive Visualizations (HTML/Plotly)</a></li>
        <li><a href="#5-code-structure">Code Structure</a>
          <ol>
            <li><a href="#51-how-to-add-changes">How to Add Changes?</a></li>
          </ol>
        </li>
        <li><a href="#6-how-to-run">How to Run?</a></li>
      </ol>
    </li>
    <li><a href="#ppt-analysis-summary">PPT Analysis Summary</a></li>
    <li><a href="#dependencies">Dependencies</a></li>
    <li><a href="#technical-considerations">Technical Considerations</a></li>
    <li><a href="#conclusion">Conclusion</a></li>
    <li><a href="#contact">Contact</a></li>
  </ol>
</details>

## Overview

This project showcases my ability to automate the generation of PowerPoint presentations from web-scraped data. I've developed a Python-based pipeline that extracts cricket statistics, processes it, and creates visually appealing PPTs with data visualizations.

<p align="right">(<a href="#readme-top">back to top</a>)</p>

## Key Skills Demonstrated

* [![Web Scraping](https://img.shields.io/badge/Web_Scraping-informational?style=for-the-badge&logoColor=white)](https://en.wikipedia.org/wiki/Web_scraping)
* [![Pandas](https://img.shields.io/badge/Pandas-150458?style=for-the-badge&logo=pandas&logoColor=white)](https://pandas.pydata.org/)
* [![python-pptx](https://img.shields.io/badge/python--pptx-informational?style=for-the-badge&logoColor=white)](https://python-pptx.readthedocs.io/en/latest/)
* [![Matplotlib](https://img.shields.io/badge/Matplotlib-%23F37B7D.svg?style=for-the-badge&logo=Matplotlib&logoColor=white)](https://matplotlib.org/)
* [![HTML5](https://img.shields.io/badge/HTML5-%23E34F26.svg?style=for-the-badge&logo=html5&logoColor=white)](https://developer.mozilla.org/en-US/docs/Web/HTML)
* [![Plotly](https://img.shields.io/badge/Plotly-2C3E50?style=for-the-badge&logo=plotly&logoColor=white)](https://plotly.com/)
* [![Python](https://img.shields.io/badge/Python-3776AB?style=for-the-badge&logo=python&logoColor=white)](https://www.python.org/)
* [![OOP](https://img.shields.io/badge/OOP-informational?style=for-the-badge&logoColor=white)](https://en.wikipedia.org/wiki/Object-oriented_programming)
* [![Automation](https://img.shields.io/badge/Automation-informational?style=for-the-badge&logoColor=white)](https://en.wikipedia.org/wiki/Automation)

<p align="right">(<a href="#readme-top">back to top</a>)</p>

## Project Details

### 1. Data Acquisition

* I employed web scraping techniques to gather data on cricket players and their century records from various online sources.
* The collected data encompasses player details (name, date of birth, place of birth, family information), career statistics (total centuries, centuries per year), and match information.

<p align="right">(<a href="#readme-top">back to top</a>)</p>

### 2. Data Processing

* The scraped data was organized into efficient Pandas DataFrames for streamlined manipulation.
* Data cleaning and transformation steps were performed to address missing values and ensure data integrity.
* The processed DataFrames were saved as Excel files (`personal_data.xlsx` and `processed_data.xlsx`) for intermediate storage.

<p align="right">(<a href="#readme-top">back to top</a>)</p>

### 3. PPT Generation

* I developed a Python script leveraging the `python-pptx` library to automate the creation of visually informative PowerPoint presentations.
* The script dynamically reads data from the generated Excel files to create individual slides for each player, including:
    * Comprehensive player information
    * Detailed century statistics
    * Well-structured tables summarizing key data points
* `Matplotlib` was utilized to generate insightful visualizations of century trends over the years, presented as clear bar and line plots within the PPT.

<p align="right">(<a href="#readme-top">back to top</a>)</p>

### 4. Interactive Visualizations (HTML/Plotly)

* To provide a more engaging data exploration experience, I created interactive visualizations using Plotly and embedded them in separate HTML pages.
* These dynamic graphs allow for features like zooming and tooltips, enabling deeper data analysis.
* *(Note: The generated PPTs are static. The interactive graphs reside in separate HTML files. Future integration could involve linking from the PPT or exporting the PPT to an HTML format.)*

<p align="right">(<a href="#readme-top">back to top</a>)</p>

### 5. Code Structure

The project's codebase is thoughtfully structured for clarity and maintainability:

* `runner.ipynb`: This Jupyter Notebook orchestrates the entire PPT creation process, acting as the main execution point.
* `prepare_data.ipynb`: This notebook manages the data gathering and web scraping phases of the project.
* `ppt_generator.py` (Class): This Python class is responsible for data transformation, the generation of static graphs (for the PPT), and the creation of interactive HTML versions of these graphs.
* `custom_presentation.py` (Class): This class handles the styling of the PowerPoint presentation, the creation of individual slides, and the population of these slides with text, tables, and images.

#### 5.1 How to Add Changes?

* **`runner.ipynb`**: Modify the `PPT_DATA` variable within this file to adjust the filters applied when generating the PPTs (e.g., specific player groups or data ranges).
* **`prepare_data.ipynb`**: Update this file to modify the data sources or the web scraping logic to work with different or updated cricket statistics.
* **`ppt_generator.py` (Class)**: Alter the data filtering and transformation logic within this class. You can also customize the appearance of the static graphs (for the PPT) and the interactive Plotly graphs (in HTML) here.
* **`custom_presentation.py` (Class)**: Modify this file to change the overall style of the generated PowerPoint presentations, including the logo, color scheme, slide layouts, and font styles.

<p align="right">(<a href="#readme-top">back to top</a>)</p>

### 6. How to Run?

1.  **`prepare_data.ipynb`**: Execute this notebook first to fetch and process the latest cricket data. This step will generate or update the `personal_data.xlsx` and `processed_data.xlsx` files.
2.  **`runner.ipynb or main.py`**: After successfully running `prepare_data.ipynb` (or if you already have the `personal_data.xlsx` and `processed_data.xlsx` files), execute this notebook or the python file. This will trigger the PPT generation process, creating the output PowerPoint files.

<p align="right">(<a href="#readme-top">back to top</a>)</p>

## PPT Analysis Summary

The project currently generates two distinct PowerPoint presentations based on player gender: "player\_Male.pptx" containing analysis for male cricket players and "player\_Female.pptx" for female players.

<p align="right">(<a href="#readme-top">back to top</a>)</p>

## Dependencies

This project relies on the following Python libraries. Ensure they are installed in your environment:

* [![Python](https://img.shields.io/badge/Python-3.11-blue?style=for-the-badge&logo=python&logoColor=white)](https://www.python.org/)
* [![Matplotlib](https://img.shields.io/badge/Matplotlib-%23F37B7D.svg?style=for-the-badge&logo=Matplotlib&logoColor=white)](https://matplotlib.org/)
* [![python-pptx](https://img.shields.io/badge/python--pptx-informational?style=for-the-badge&logoColor=white)](https://python-pptx.readthedocs.io/en/latest/)
* [![Pandas](https://img.shields.io/badge/Pandas-150458?style=for-the-badge&logo=pandas&logoColor=white)](https://pandas.pydata.org/)
* [![Plotly](https://img.shields.io/badge/Plotly-2C3E50?style=for-the-badge&logo=plotly&logoColor=white)](https://plotly.com/)
* [![SciPy](https://img.shields.io/badge/SciPy-%230C52A5.svg?style=for-the-badge&logo=scipy&logoColor=white)](https://scipy.org/)

You can install these dependencies using pip:
```sh
pip install matplotlib python-pptx pandas plotly scipy
```
The application will open a window displaying the webcam feed with detected cards and the identified poker hand.

You can also adapt the capture variable in the script to process a video file instead of a live webcam feed.

<p align="right">(<a href="#readme-top">back to top</a>)</p>

## Technical Considerations

* Robust error handling has been implemented within the web scraping and data processing stages to ensure the pipeline's stability.
* The codebase is designed with a modular architecture to enhance maintainability and facilitate future extensions.
* Potential future improvements include:
  * Implementing dynamic data updates to keep the presentations current.
  * Utilizing external configuration files for easier customization.
  * Further breaking down the code into smaller, more specialized modules.
  * Adding comprehensive logging for better monitoring and debugging.
  * Incorporating unit tests to ensure the reliability of individual components.

<p align="right">(<a href="#readme-top">back to top</a>)</p>

## Conclusion
This project effectively demonstrates my ability to integrate diverse technical skills to create a fully automated workflow, from extracting raw data to generating insightful and visually appealing presentations. I am enthusiastic about leveraging and further developing these skills in a professional environment.

<p align="right">(<a href="#readme-top">back to top</a>)</p>

## Contact

Feel free to reach out if you have any questions, suggestions, or would like to collaborate!

* [![Name](https://img.shields.io/badge/Ayush%20Dhanraj-informational?style=for-the-badge)](https://www.linkedin.com/in/ayush-d-1759461a1)
* [![Email](https://img.shields.io/badge/dhanrajaayush123%40gmail.com-important?style=for-the-badge)](mailto:dhanrajaayush123@gmail.com)
* [![Email](https://img.shields.io/badge/ayushdhanraj.work%40gmail.com-important?style=for-the-badge)](mailto:ayushdhanraj.work@gmail.com)
* [![LinkedIn](https://img.shields.io/badge/LinkedIn-0077B5?style=for-the-badge&logo=linkedin&logoColor=white)](https://www.linkedin.com/in/ayush-d-1759461a1)
* [![GitHub](https://img.shields.io/badge/GitHub-100000?style=for-the-badge&logo=github&logoColor=white)](https://github.com/Ayush-2001-Dhanraj)

<p align="right">(<a href="#readme-top">back to top</a>)</p>
