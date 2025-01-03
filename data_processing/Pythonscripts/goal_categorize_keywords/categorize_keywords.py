import spacy
import os
import pandas as pd
from collections import defaultdict
from sklearn.metrics.pairwise import cosine_similarity

# Load spaCy model
nlp = spacy.load("en_core_web_md")  

# Define categories with seed words
categories = {
    "Geographic Regions": ["Mediterranean", "Europe", "Middle East", "Caucasus"],
    "Political Terms": ["democracy", "election", "policy", "government"],
    "Politicians": ["Donald Trump", "Trump", "Joe Biden", "Biden", "Macron", "Emmanuel Macron", "Nikol Pashinyan", "Pashinyan", "Ilham Aliyev", "Putin"],
    "Alliances": ["EU", "EEU", "CSTO", "NATO"],
    "Organizations": ["ARS", "AGBU", "ARF", "AYF", "ANCA"],
    "Food": ["manti", "cucumber", "apricot", "liquor", "bread", "cheese", "lettuce", "recipe"],
    "Education": ["school", "learn", "literacy", "education", "writing", "reading", "language"],
    "Nationalities": ["Armenian", "Azeri", "Azerbaijani", "Israeli", "Russian", "French", "German"],
    "Cities": ["Gyumri", "Yerevan", "Baku", "Paris", "Berlin", "Moscow", "Toronto"],
    "Countries": ["Armenia", "Azerbaijan", "Canada", "France", "Russia", "China", "USA", "America"],
    "Genocide": ["genocide", "1915", "recognition", "Talaat Pasha"],
    "Religion": ["church", "diocese", "archbishop", "prayer", "Christianity", "Islam", "Christian", "Jew"],
    "Culture": ["poetry", "art", "culture", "literature", "dance", "music", "song", "book"],
    "Blockade": ["blockade", "corridor"],
    "Social Issues": ["violence", "charity", "humanitarian", "aid"],
    "Social Services": ["healthcare", "education", "public transit", "welfare", "subsidies"],
    "Economic Terms": ["tax", "economic", "jobs", "spending", "budget", "dollar"],
    "Women's Issues": ["menstrual", "domestic", "pregnancy", "maternity", "feminine", "woman", "mother"],
}

# Preprocess seed words to compute category vectors
category_vectors = {cat: nlp(" ".join(words)).vector for cat, words in categories.items()}

def categorize_keywords(keywords, category_vectors):
    categorized = defaultdict(list)
    for keyword in keywords:
        keyword_vector = nlp(keyword).vector.reshape(1, -1)  # Ensure it's 2D
        similarities = {
            cat: cosine_similarity(keyword_vector, category_vector.reshape(1, -1))[0][0]
            for cat, category_vector in category_vectors.items()
        }
        best_match = max(similarities, key=similarities.get)
        categorized[best_match].append(keyword)
    return categorized


# Load input spreadsheet
input_file = "Normalized_Armenian_Womens_Articles.xlsx"
df = pd.read_excel(input_file)

# Check for "Keyword" column
if "Keywords" not in df.columns:
    raise ValueError("Input spreadsheet must contain a 'Keywords' column.")

keywords = df["Keywords"].dropna().tolist()  # Remove any NaN values and convert to list

# Categorize keywords
categorized_keywords = categorize_keywords(keywords, category_vectors)

# Define output directory and ensure it exists
output_dir = "./output/"
if not os.path.exists(output_dir):
    os.makedirs(output_dir)

# Save each category as a separate spreadsheet
for category, words in categorized_keywords.items():
    # Find rows in the original DataFrame matching the keywords
    filtered_rows = df[df["Keywords"].isin(words)]
    
    # Prepare output file name
    sanitized_category = category.replace(' ', '_').replace('/', '_')
    output_file = os.path.join(output_dir, f"{sanitized_category}.xlsx")
    
    # Save the filtered rows
    try:
        filtered_rows.to_excel(output_file, index=False)
        print(f"Category '{category}' saved to {output_file}")
    except Exception as e:
        print(f"Failed to save category '{category}' to {output_file}: {e}")

