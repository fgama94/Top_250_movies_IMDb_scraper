# IMDb Top 250 Movies Scraper

This Python script fetches and extracts information about the top 250 movies from IMDb, using Selenium for web scraping and asyncio for asynchronous operations. It extracts details like the original movie title, year, duration, and rating, and exports them to an Excel file.



# Features

- Web Scraping: Utilizes Selenium WebDriver to fetch dynamic content from IMDb's top movies page.

- Asynchronous Processing: Uses asyncio to handle concurrent requests for faster data extraction.

- Logging: Logs errors and information during execution, storing logs in a dedicated directory.

- Export to Excel: Generates an Excel file (Top250MoviesIMDB_<timestamp>.xlsx) containing extracted movie data.



# Setup

Prerequisites:

- Python 3.x
  
- Chrome WebDriver (chromedriver.exe included in the repository) - might need to update the chromedriver (https://googlechromelabs.github.io/chrome-for-testing/)
  

Required Python packages:
  
    asyncio
    
    urllib
    
    openpyxl
    
    beautifulsoup4
    
    selenium



# Installation

Clone the repository:

    git clone https://github.com/your-username/imdb-top-250-scraper.git
  
    cd imdb-top-250-scraper


Install dependencies if not already installed:

    pip install asyncio urllib3 openpyxl beautifulsoup4 selenium
  

Ensure chromedriver.exe is in the project directory.



# Usage

Run the script imdb_top_250_scraper.py:

    python imdb_top_250_scraper.py
  

The script will fetch IMDb's top 250 movies data, process it, and generate an Excel file (Top250MoviesIMDB_<timestamp>.xlsx) in the Top 250 Movies IMDB directory.



# Logs

Logs are stored in the Logs directory, named imdb_<timestamp>.log, capturing execution details and errors encountered during scraping.
