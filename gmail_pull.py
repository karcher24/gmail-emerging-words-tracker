import os
import base64
import json
import re
import pickle
from datetime import datetime
from googleapiclient.discovery import build
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from bs4 import BeautifulSoup

# If modifying the scope, delete the token.pickle file.
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']

# Stopwords to ignore for single-word analysis
STOPWORDS = set([
    'the', 'and', 'for', 'that', 'with', 'you', 'are', 'was', 'this', 'but',
    'not', 'have', 'from', 'they', 'will', 'your', 'all', 'our', 'has', 'can',
    'their', 'about', 'who', 'what', 'when', 'where', 'which', 'would', 'there',
    'been', 'were', 'more', 'had', 'she', 'him', 'her', 'his', 'out', 'over', 'into',
    'also', 'as', 'of', 'in', 'on', 'an', 'is', 'to', 'at', 'by', 'it', 'be', 'a'
])

def authenticate_gmail():
    """Authenticate the Gmail API."""
    creds = None
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)
    return build('gmail', 'v1', credentials=creds)

def clean_html(raw_html):
    """Removes HTML tags from a string."""
    soup = BeautifulSoup(raw_html, "html.parser")
    return soup.get_text()

def analyze_words_and_bigrams(text):
    """Analyze frequency of words (excluding stopwords) and bigrams (all allowed) in text."""
    words = re.findall(r'\b\w+\b', text.lower())
    word_counts = {}
    bigram_counts = {}

    for i, word in enumerate(words):
        if len(word) > 2:  # Ignore very short words
            # Only count word if not a stopword
            if word not in STOPWORDS:
                word_counts[word] = word_counts.get(word, 0) + 1

            # Always count bigrams (even if stopwords)
            if i < len(words) - 1:
                next_word = words[i + 1]
                if len(next_word) > 2:
                    bigram = f"{word} {next_word}"
                    bigram_counts[bigram] = bigram_counts.get(bigram, 0) + 1

    return word_counts, bigram_counts

def save_counts(word_counts, bigram_counts, filename='word_counts.json'):
    """Save updated word and bigram counts."""
    if os.path.exists(filename):
        with open(filename, 'r') as f:
            saved = json.load(f)
            old_words = saved.get('words', {})
            old_bigrams = saved.get('bigrams', {})
    else:
        old_words = {}
        old_bigrams = {}

    for word, count in word_counts.items():
        old_words[word] = old_words.get(word, 0) + count

    for bigram, count in bigram_counts.items():
        old_bigrams[bigram] = old_bigrams.get(bigram, 0) + count

    with open(filename, 'w') as f:
        json.dump({'words': old_words, 'bigrams': old_bigrams}, f, indent=2)

def load_counts(filename='word_counts.json'):
    """Load saved word and bigram counts."""
    if os.path.exists(filename):
        with open(filename, 'r') as f:
            saved = json.load(f)
            return saved.get('words', {}), saved.get('bigrams', {})
    return {}, {}

def find_new_or_emerging(old_counts, new_counts, threshold=5):
    """Find new or spiking words or bigrams."""
    new_items = {}
    for item, count in new_counts.items():
        old_count = old_counts.get(item, 0)
        if old_count == 0:
            new_items[item] = count
        elif count > old_count * threshold:
            new_items[item] = count
    return new_items

def save_emerging_today(words, bigrams):
    """Save today's emerging words/bigrams into a date-stamped JSON file."""
    today = datetime.now().strftime('%Y-%m-%d')
    filename = f'emerging_{today}.json'
    with open(filename, 'w') as f:
        json.dump({'emerging_words': words, 'emerging_bigrams': bigrams}, f, indent=2)
    print(f"\nToday's emerging results saved to: {filename}")

def get_emails():
    """Main function to pull emails, extract text, and find emerging words and bigrams."""
    service = authenticate_gmail()
    results = service.users().messages().list(userId='me', labelIds=['INBOX'], q="is:unread").execute()
    messages = results.get('messages', [])

    all_text = ""

    if not messages:
        print('No new messages.')
    else:
        print('Pulling messages...')
        for message in messages:
            msg = service.users().messages().get(userId='me', id=message['id']).execute()
            payload = msg['payload']
            body_text = None

            if 'parts' in payload:
                for part in payload['parts']:
                    mime_type = part.get('mimeType')
                    if mime_type == 'text/plain':
                        data = part['body'].get('data')
                        if data:
                            byte_code = base64.urlsafe_b64decode(data)
                            body_text = byte_code.decode('utf-8')
                            break
                    elif mime_type == 'text/html' and not body_text:
                        data = part['body'].get('data')
                        if data:
                            byte_code = base64.urlsafe_b64decode(data)
                            body_text = byte_code.decode('utf-8')
            else:
                data = payload['body'].get('data')
                if data:
                    byte_code = base64.urlsafe_b64decode(data)
                    body_text = byte_code.decode('utf-8')

            if body_text:
                if '<html' in body_text.lower():
                    body_text = clean_html(body_text)

                all_text += body_text + "\n"

    if all_text:
        # Analyze words and bigrams
        new_word_counts, new_bigram_counts = analyze_words_and_bigrams(all_text)
        print(f"\nWords found: {len(new_word_counts)} unique words.")
        print(f"Bigrams found: {len(new_bigram_counts)} unique bigrams.")

        # Load old counts
        old_word_counts, old_bigram_counts = load_counts()

        # Find new/emerging words and bigrams
        emerging_words = find_new_or_emerging(old_word_counts, new_word_counts)
        emerging_bigrams = find_new_or_emerging(old_bigram_counts, new_bigram_counts)

        print("\nEmerging/New Words:")
        for word, count in sorted(emerging_words.items(), key=lambda x: -x[1]):
            print(f"{word}: {count}")

        print("\nEmerging/New Bigrams:")
        for bigram, count in sorted(emerging_bigrams.items(), key=lambda x: -x[1]):
            print(f"{bigram}: {count}")

        # Save updated counts
        save_counts(new_word_counts, new_bigram_counts)

        # Save today's emerging results
        save_emerging_today(emerging_words, emerging_bigrams)

if __name__ == '__main__':
    get_emails()
