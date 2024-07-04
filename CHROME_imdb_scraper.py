
import asyncio
import urllib.request, urllib.parse, urllib.error
import logging
import os
import openpyxl
import datetime
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Get the directory of the current script
script_dir = os.path.dirname(os.path.abspath(__file__))

# Set up logging
log_dir = os.path.join(script_dir, 'Logs')
os.makedirs(log_dir, exist_ok=True)
log_file = os.path.join(log_dir, f'imdb_{datetime.datetime.now().strftime("%Y%m%d_%H%M%S")}.log')
logging.basicConfig(
    filename=log_file,
    level=logging.INFO, 
    format='%(asctime)s - %(levelname)s - %(message)s'
)

# Initialize Chrome driver instance
chromedriver_path = os.path.join(script_dir, 'chromedriver.exe')
service = Service(chromedriver_path)
driver = webdriver.Chrome(service=service)

async def fetch_html(url):
    try:
        # Open the URL
        driver.get(url)

        # Wait for the page to load completely
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, 'body')))

        # Get the HTML content of the page
        html_content = driver.page_source

        return html_content
    
    except Exception as e:
        # Log any errors encountered during fetching HTML content
        logging.error(f"Error fetching HTML from {url}: {e}")
        return None


def extract_ranking(href):
    parts = href.split('_')
    ranking = parts[-1]
    return ranking


async def extract_movies_data(html, url_1):
    soup = BeautifulSoup(html, 'html.parser')
    movies_data = []    # List to store tuples of ranking and title

    batch_size = 10
    num_movies = 250
    num_processed = 0

    # Extract both poster and title links and combine them
    poster_links = soup.find_all('a', class_='ipc-lockup-overlay ipc-focusable')
    title_links = soup.find_all('a', class_='ipc-title-link-wrapper')
    combined_links = list(zip(poster_links, title_links))
    
    # Process movie links in batches
    for batch_start in range(1, num_movies + 1, batch_size):
        batch_end = min(batch_start + batch_size, num_movies + 1)
        batch_combined_links = combined_links[batch_start - 1:batch_end - 1]

        tasks = []
        for poster_link, title_link in batch_combined_links:
            ranking = extract_ranking(title_link['href'])
            url_2 = urllib.parse.urljoin(url_1, title_link['href'])
            logging.info(f"Retrieving data from {url_2}")
            print(f"Retrieving data from {url_2}")
            
            # Fetch HTML content and extract movie info concurrently
            tasks.append(asyncio.create_task(fetch_and_extract_movie_data(url_2, ranking)))
        
        # Wait for all tasks to complete
        if tasks:
            movies_info = await asyncio.gather(*tasks)
            movies_info = [info for info in movies_info if info is not None]  # Filter out skipped movies
            movies_data.extend(movies_info)
            num_processed += len(movies_info)

        logging.info('Finished processing batch.')
        print('Finished processing batch.')

    logging.info(f"Total movies extracted: {num_processed}")
    print(f"Total movies extracted: {num_processed}")

    return movies_data


async def fetch_and_extract_movie_data(url, ranking):
    try:
        html = await fetch_html(url)
        if html is None:
            raise Exception("HTML content is None")

        return await extract_movie_info(html, ranking)
    except Exception as e:
        logging.warning(f"Error processing URL {url}: {e}")
        print(f"Error processing URL {url}: {e}")
        return {
            'ranking': ranking,
            'title': None,
            'year': None,
            'duration': None,
            'rating': None
        }
    

async def extract_movie_info(html, ranking):
    soup = BeautifulSoup(html, 'html.parser')

    skipped_movies = []  # List to store tuples of ranking and title of skipped movies

    # Check for the presence of 'sc-d8941411-1 fTeJrK' class
    title = soup.find(class_='sc-d8941411-1 fTeJrK')

    # If the 'sc-d8941411-1 fTeJrK' class is not found, search for 'hero__primary-text' class
    if not title:
        title = soup.find(class_='hero__primary-text')
        if not title:
            # Return None if movie is skipped
            return None

    title_text = title.get_text(strip=True)
    title_text = title_text.replace("Original title: ", "").replace("Título original: ", "")
    # Extract year, duration, and rating
    year_elements = soup.find_all('a', href=lambda href: href and '/releaseinfo' in href)
    duration_elements = soup.find_all('li', class_='ipc-inline-list__item')
    rating_elements = soup.find_all('span', class_='sc-bde20123-1 cMEQkK')
    year = year_elements[0].get_text(strip=True) if year_elements else None
    duration = next((element.get_text(strip=True) for element in duration_elements if 'h' in element.get_text(strip=True) or 'm' in element.get_text(strip=True)), None)
    rating = rating_elements[0].get_text(strip=True) if rating_elements else None
    
    return {
        'ranking': ranking,
        'title': title_text,
        'year': year,
        'duration': duration,
        'rating': rating
    }


def export_to_excel(movies_data):
    # Create a new Excel workbook
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = 'Movies'

    # Write headers
    sheet['A1'] = 'Ranking'
    sheet['B1'] = 'Title'
    sheet['C1'] = 'Year'
    sheet['D1'] = 'Duration'
    sheet['E1'] = 'Rating'

    # Write data
    for row, movie_data in enumerate(movies_data, start=2):
        sheet[f'A{row}'] = movie_data['ranking']
        sheet[f'B{row}'] = movie_data['title']
        sheet[f'C{row}'] = movie_data['year']
        sheet[f'D{row}'] = movie_data['duration']
        sheet[f'E{row}'] = movie_data['rating']

    # Check if the directory exists, if not, create it
    directory = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'Top 250 Movies IMDB')
    if not os.path.exists(directory):
        os.makedirs(directory)

    # Save the workbook inside the directory with the filename appended with date
    file_name = os.path.join(directory, f'Top250MoviesIMDB_{datetime.datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx')
    workbook.save(file_name)

    logging.info("Excel file created successfully.")
    print("Excel file created successfully.")

    return file_name


async def main():
    logging.info('Starting...')
    print('Starting...')

    try:
        # Fetch HTML content from IMDb top movies page
        url_1 = 'https://www.imdb.com/chart/top/'
        html_1 = await fetch_html(url_1)

        # Extract movies from the BeautifulSoup object and store them in a list
        movies_task = extract_movies_data(html_1, url_1)

        # Wait for both tasks to complete
        movies = await movies_task

        # Export data to Excel
        exported_file_path = export_to_excel(movies)

        # Log movies extracted and skipped movies
        logging.info(f"Movies extracted: {movies}")
        skipped_movies = [movie for movie in movies if movie is None]
        logging.info(f"Skipped movies: {skipped_movies}")

        return exported_file_path
    
    except Exception as e:
        # Log any unexpected errors encountered during main execution
        logging.error(f"Error in main(): {e}")
        return None

if __name__ == "__main__":
    asyncio.run(main())
