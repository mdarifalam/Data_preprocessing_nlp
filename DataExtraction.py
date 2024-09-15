import pandas as pd
import requests
from bs4 import BeautifulSoup
import os

# Read the input Excel file
input_path = 'Input/Input.xlsx'
df = pd.read_excel(input_path)


# Defining a function to fetch and parse article content
def fetch_article_content(url):
    try:
        response = requests.get(url)
        response.raise_for_status()

        soup = BeautifulSoup(response.content, 'html.parser')

        # Assuming the article title is within <h1> tags and the main content is within <article> tags
        title = soup.find('h1').get_text(strip=True) if soup.find('h1') else ''
        article = soup.find('article')

        # If the article tag is not found, try a common fallback 
        if not article:
            article = soup.find('div', {'class': 'article-content'}) 

        if article:
            paragraphs = article.find_all('p')
            article_text = '\n\n'.join([para.get_text(strip=True) for para in paragraphs])
        else:
            article_text = ''

        return title, article_text
    except Exception as e:
        print(f"Error fetching content from {url}: {e}")
        return '', ''


# Directory name articles to save the text files
output_dir = 'articles'
os.makedirs(output_dir, exist_ok=True)


for index, row in df.iterrows():
    url_id = row['URL_ID']
    url = row['URL']

    title, article_text = fetch_article_content(url)

    # Combine title and article text
    content = title + '\n\n' + article_text if title and article_text else article_text

    # Save the content to a text file
    if content:
        file_path = os.path.join(output_dir, f"{url_id}.txt")
        with open(file_path, 'w', encoding='utf-8') as file:
            file.write(content)

print("Article extraction and saving completed.")
