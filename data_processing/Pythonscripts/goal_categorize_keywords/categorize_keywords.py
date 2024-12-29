import spacy
import pandas as pd
from collections import defaultdict
from sklearn.metrics.pairwise import cosine_similarity

# Load spaCy model
nlp = spacy.load("en_core_web_md")  # Use a medium model for better accuracy

# Define categories with seed words
categories = {
    "Geographic Regions": ["Paris", "USA", "Europe", "Asia"],
    "Political Terms": ["democracy", "election", "policy", "government"],
    "Politicians": ["Trump", "Biden", "Obama", "Macron"],
    "Alliances": ["Trump", "Biden", "Obama", "Macron"],
    "Food": ["Trump", "Biden", "Obama", "Macron"],
    "Education": ["Trump", "Biden", "Obama", "Macron"],
    "Social Issues": ["Trump", "Biden", "Obama", "Macron"],
    "Cities": ["Trump", "Biden", "Obama", "Macron"],
    "Countries": ["Trump", "Biden", "Obama", "Macron"],
    "Genocide": ["Trump", "Biden", "Obama", "Macron"],
    "Religion": ["Trump", "Biden", "Obama", "Macron"],
    "Culture": ["Trump", "Biden", "Obama", "Macron"],
    "Blockade": ["Trump", "Biden", "Obama", "Macron"],
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
input_file = "input.xlsx"
df = pd.read_excel(input_file)

# Check for "Keyword" column
if "Keyword" not in df.columns:
    raise ValueError("Input spreadsheet must contain a 'Keyword' column.")

keywords = df["Keyword"].dropna().tolist()  # Remove any NaN values and convert to list

# Categorize keywords
categorized_keywords = categorize_keywords(keywords, category_vectors)

# Save each category as a separate spreadsheet
output_dir = "./output/"
for category, words in categorized_keywords.items():
    output_file = f"{output_dir}{category.replace(' ', '_')}.xlsx"
    category_df = pd.DataFrame(words, columns=["Keyword"])
    category_df.to_excel(output_file, index=False)
    print(f"Category '{category}' saved to {output_file}")
