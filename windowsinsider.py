from datetime import datetime, timedelta
import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from sumy.parsers.plaintext import PlaintextParser
from sumy.nlp.tokenizers import Tokenizer
from sumy.summarizers.lsa import LsaSummarizer
from deep_translator import GoogleTranslator
import nltk
import time

nltk.download('punkt') 

headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
    }
# Function to find div elements with class "c-card__date"
def find_date_divs(url):
    

    with requests.Session() as session:
        response = session.get(url, headers=headers)
        time.sleep(1)  # Wait for 1 second between requests

        if response.status_code == 200:
            soup = BeautifulSoup(response.text, 'lxml')
            date_elements = soup.find_all('div', class_=lambda value: value and 'c-card__date' in value)
            return date_elements
        else:
            print(f"Failed to retrieve data. Status code: {response.status_code}")
            return []

# Function to get links for yesterday
def get_links_for_yesterday(url):
    yesterday = datetime.now() - timedelta(days=1)
    yesterday_date = yesterday.strftime('%B %d, %Y')

    date_elements = find_date_divs(url)
    links_yesterday = []

    for date_element in date_elements:
        date_string = date_element.get_text(strip=True)

        try:
            date = datetime.strptime(date_string, '%B %d, %Y')
        except ValueError:
            continue

        if date.strftime('%B %d, %Y') == yesterday_date:
            link = date_element.find_parent('li', class_='s-list__item').find('a', class_='c-card__link')
            if link and 'href' in link.attrs:
                links_yesterday.append(link['href'])

    return links_yesterday

# Function to download title and summary of an article
def get_article_info(url):
    article_info = {
        'link': url,
        'title': '',
        'summary': ''
    }
    
    try:
        article_html = requests.get(url, headers=headers).text
        article_soup = BeautifulSoup(article_html, 'lxml')
        
       
        article_info['title'] = article_soup.find('h1', class_='h1').text.strip()
        article_content = article_soup.find('div', class_='item-single__content t-content')
        if article_content:
            article_text = article_content.get_text(separator='\n').strip()

            LANGUAGE = "english"
            parser = PlaintextParser.from_string(article_text, Tokenizer(LANGUAGE))
            summarizer = LsaSummarizer()
            summary = summarizer(parser.document, 3)

            article_info['summary'] = ' '.join(map(str, summary))

            translator_title = GoogleTranslator(source='en', target='pl')
            translated_title = translator_title.translate(article_info['title'])
            article_info['title'] = translated_title
            
            translator = GoogleTranslator(source='en', target='pl')
            translated_summary = translator.translate(article_info['summary'])
            article_info['summary'] = translated_summary
            

    except Exception as e:
        print(f"An error occurred: {e}")

    return article_info

# Function to save data to an Excel file
def save_to_excel(articles, date):
    file_name = f'{date}_windowsblogs.xlsx'
    wb = Workbook()
    ws = wb.active
    ws.append(['Link', 'Tytuł artykułu', 'Streszczenie'])

    for article in articles:
        ws.append([article['link'], article['title'], article['summary']])

    wb.save(file_name)

# Function to get yesterday's date
def get_yesterday_date():
    yesterday = datetime.now() - timedelta(days=1)
    return yesterday.strftime('%Y-%m-%d')

if __name__ == "__main__":
    url = 'https://blogs.windows.com/'
    
    # Get links from yesterday
    links = get_links_for_yesterday(url)

    # Get info about articles and summary
    articles_info = [get_article_info(link) for link in links]

    # Get yesterday's date for the Excel file
    date = get_yesterday_date()

    # Save data to an Excel file
    save_to_excel(articles_info, date)
