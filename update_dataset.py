import requests
import pandas as pd
from datetime import datetime

def get_metadata_and_citations(doi):
    base_url = 'https://api.crossref.org/works/'
    url = f'{base_url}{doi}'
    
    try:
        response = requests.get(url)
        response.raise_for_status()  # Check for any HTTP errors

        data = response.json()
        metadata = data['message']
        
        # Fetch citation count
        citation_count = metadata.get('is-referenced-by-count', 0)
        
        # Fetch volume
        volume = metadata.get('volume', '')
        
        # Fetch publication year
        publication_year = metadata.get('created', {}).get('date-parts', [[None]])[0][0]
        
        # Convert publication date to desired format
        publication_date_parts = metadata.get('published-print', {}).get('date-parts', [[]])[0]
        if len(publication_date_parts) >= 2:
            publication_date = datetime(year=publication_date_parts[0], month=publication_date_parts[1], day=1).strftime('%b %Y')
        else:
            publication_date = ''
        
        # Create metadata dictionary
        metadata_dict = {
            'Title': metadata.get('title', ''),
            'Authors': [author['given'] + ' ' + author['family'] for author in metadata.get('author', [])],
            'Corporate Authors': metadata.get('publisher', ''),
            'Publication Date': publication_date,
            'Publication Year': publication_year,
            'Total Citations': citation_count,
            'Average per Year': citation_count / ((datetime.now().year +1) - publication_year) if publication_year else None,
            'DOI': metadata.get('DOI', ''),
        }
        
        return metadata_dict
    except requests.exceptions.RequestException as e:
        print(f"Error retrieving metadata: {e}")
        return None


# Initialize an empty list to store metadata
metadata_list = []

# Fetch and accumulate metadata for each DOI
combined_df = pd.read_excel('dataset.xlsx')
for doi in combined_df['DOI']:
    metadata = get_metadata_and_citations(doi)
    if metadata:
        metadata_list.append(metadata)

# Create a DataFrame from the accumulated metadata
if metadata_list:
    updated_df = pd.DataFrame(metadata_list)
    
    # Remove brackets from the title column
    updated_df['Title'] = updated_df['Title'].astype('string').str[2:-2]
    updated_df['Authors'] = updated_df['Title'].astype('string').str[1:-1]
else:
    print("No metadata found or retrieval failed.")
    
updated_df.to_excel('updated_dataset.xlsx', index=False)


