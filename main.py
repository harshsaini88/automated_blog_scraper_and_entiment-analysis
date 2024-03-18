import os
import re
import pandas as pd
import requests
from bs4 import BeautifulSoup

def read_input_file(input_file):
    """Read URLs from input Excel file."""
    try:
        df = pd.read_excel(input_file)
        urls = df['URL'].tolist()
        return urls
    except Exception as e:
        print(f"Error reading input file: {e}")

def load_stop_words(stop_words_dir):
    """Load stop words from multiple files and merge them into one set."""
    stop_words = set()
    for file_name in os.listdir(stop_words_dir):
        with open(os.path.join(stop_words_dir, file_name), 'r') as file:
            stop_words.update(file.read().splitlines())
    return stop_words

def clean_text(text, stop_words):
    """Clean the text by removing stop words and punctuation."""
    cleaned_text = re.sub(r'[^\w\s]', '', text.lower())  # Remove punctuation and convert to lowercase
    cleaned_text = ' '.join(word for word in cleaned_text.split() if word not in stop_words)  # Remove stop words
    return cleaned_text

def calculate_scores(cleaned_text, positive_dict, negative_dict):
    """Calculate positive score, negative score, and polarity score."""
    positive_score = sum(1 for word in cleaned_text.split() if word in positive_dict)
    negative_score = sum(1 for word in cleaned_text.split() if word in negative_dict)
    polarity_score = (positive_score - negative_score) / ((positive_score + negative_score) + 0.000001)
    return positive_score, negative_score, polarity_score

def calculate_subjectivity_score(cleaned_text, positive_score, negative_score):
    """Calculate subjectivity score."""
    total_words = len(cleaned_text.split())
    subjectivity_score = (positive_score + negative_score) / (total_words + 0.000001)
    return subjectivity_score

def calculate_complex_words(cleaned_text):
    """Calculate the percentage of complex words."""
    words = cleaned_text.split()
    total_words = len(words)
    complex_words = sum(1 for word in words if len(word) > 2)  # Assuming words with length > 2 are complex
    percentage_complex_words = (complex_words / total_words) * 100
    return percentage_complex_words

def calculate_fog_index(average_sentence_length, percentage_complex_words):
    """Calculate the Fog Index."""
    fog_index = 0.4 * (average_sentence_length + percentage_complex_words)
    return fog_index

def calculate_average_word_length(cleaned_text):
    """Calculate the average word length."""
    words = cleaned_text.split()
    total_characters = sum(len(word) for word in words)
    average_word_length = total_characters / len(words)
    return average_word_length

def calculate_personal_pronouns(text):
    """Calculate the count of personal pronouns."""
    personal_pronouns = re.findall(r'\b(I|we|my|ours|us)\b', text)
    return len(personal_pronouns)

def analyze_article(article_text, stop_words, positive_dict, negative_dict):
    """Analyze sentiment and derive metrics for a single article."""
    cleaned_text = clean_text(article_text, stop_words)
    positive_score, negative_score, polarity_score = calculate_scores(cleaned_text, positive_dict, negative_dict)
    subjectivity_score = calculate_subjectivity_score(cleaned_text, positive_score, negative_score)
    average_sentence_length = len(cleaned_text.split('.'))  # Assuming each sentence ends with a period
    percentage_complex_words = calculate_complex_words(cleaned_text)
    fog_index = calculate_fog_index(average_sentence_length, percentage_complex_words)
    average_word_length = calculate_average_word_length(cleaned_text)
    personal_pronouns_count = calculate_personal_pronouns(article_text)
    return {
        "Positive Score": positive_score,
        "Negative Score": negative_score,
        "Polarity Score": polarity_score,
        "Subjectivity Score": subjectivity_score,
        "Average Sentence Length": average_sentence_length,
        "Percentage of Complex Words": percentage_complex_words,
        "Fog Index": fog_index,
        "Average Word Length": average_word_length,
        "Personal Pronouns Count": personal_pronouns_count
    }

def scrape_website(url):
    """Scrape content from a website."""
    try:
        response = requests.get(url)
        if response.status_code == 200:
            soup = BeautifulSoup(response.content, 'html.parser')
            
            # Extract content from both div classes
            article_content_div1 = soup.find('div', class_='td-post-content tagdiv-type')
            article_content_div2 = soup.find('div', class_='vc_column tdi_120 wpb_column vc_column_container tdc-column td-pb-span8')
            
            article_content = ""

            # Extract text from paragraphs in the first div
            if article_content_div1:
                paragraphs = article_content_div1.find_all('p')
                article_content += '\n'.join(paragraph.get_text() for paragraph in paragraphs) + "\n"

            # Extract text from any text elements in the second div
            if article_content_div2:
                text_elements = article_content_div2.find_all(text=True)
                article_content += '\n'.join(element.strip() for element in text_elements if element.strip()) + "\n"

            if article_content:
                return article_content
            else:
                print(f"No article content found on {url}")
                return None
        else:
            print(f"Failed to scrape {url}: Status code {response.status_code}")
            return None
    except Exception as e:
        print(f"Error scraping {url}: {e}")
        return None

def main():
    # Directory containing stop words files
    stop_words_dir = 'stop_words'
    # Directory containing positive and negative dictionaries
    sentiment_dicts_dir = 'sentiment_dicts'

    # Load stop words
    stop_words = load_stop_words(stop_words_dir)

    # Load positive and negative dictionaries
    with open(os.path.join(sentiment_dicts_dir, 'positive-words.txt'), 'r') as file:
        positive_dict = set(file.read().splitlines())
    with open(os.path.join(sentiment_dicts_dir, 'negative-words.txt'), 'r') as file:
        negative_dict = set(file.read().splitlines())

    # Read URLs from input Excel file
    input_file = 'input.xlsx'  # Path to input Excel file
    urls = read_input_file(input_file)

    # Directory to store scraped content
    output_dir = 'processed_data'
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # Results dictionary to store metrics for each article
    results = {}

    # Scrape and analyze each article
    for idx, url in enumerate(urls, start=1):
        print(f"Scraping website {idx}/{len(urls)}: {url}")
        content = scrape_website(url)
        if content:
            results[f'article_{idx}.txt'] = analyze_article(content, stop_words, positive_dict, negative_dict)
            output_file = os.path.join(output_dir, f'article_{idx}.txt')
            with open(output_file, 'w', encoding='utf-8') as f:
                f.write(content)
            print(f"Content saved to {output_file}")
        else:
            print(f"No content scraped from {url}")

    # Convert results to DataFrame
    df_results = pd.DataFrame(results).T

    # Sort the results DataFrame to match the order of articles in the input Excel file
    df_input = pd.read_excel(input_file)
    df_results_sorted = df_results.reindex([f'article_{idx}.txt' for idx in range(1, len(urls)+1)])

    # Merge input data with analysis results
    df_output = pd.concat([df_input.reset_index(drop=True), df_results_sorted.reset_index(drop=True)], axis=1)

    # Write output to Excel file
    output_file = 'output.xlsx'
    df_output.to_excel(output_file, index=False)
    print(f"Output saved to {output_file}")

if __name__ == "__main__":
    main()
