import os
import pandas as pd
from collections import defaultdict
from sklearn.metrics.pairwise import cosine_similarity
from sentence_transformers import SentenceTransformer

# Load transformer model for embeddings
model = SentenceTransformer("sentence-transformers/all-MiniLM-L6-v2")

# Define categories with seed words
categories = {
    "Geographic Regions": ["Mediterranean", "sevan", "mediterranean", "Europe", "Middle East", "karabakh", "Caucasus", "caucasus", "nagorno-karabakh", "west", "crimea"],
    "Political Terms": ["democracy", "election", "policy", "security", "minister", "normalization", "territorial", "ministry", "corruption", "rights", "civic", "state", "government", "vote", "independence", "revolution", "protest", "justice", "marches"],
    "Politicians": ["Donald Trump", "Trump", "trump", "biden", "putin", "Joe Biden", "Biden", "Macron", "Emmanuel Macron", "Nikol Pashinyan", "Pashinyan", "pashinyan", "macron", "Ilham Aliyev", "Putin", "aliyev", "ilham", "Kocharyan", "erdogan"],
    "Alliances": ["EU", "EEU", "CSTO", "NATO"],
    "Organizations": ["ARS", "AGBU", "ARF", "AYF", "ANCA", "homenetmen", "hamazkayin"],
    "Food": ["manti", "cucumber", "lavash", "kibbeh", "bakery","apricot", "liquor", "vodka", "cognac", "wine", "bread", "dinner", "lunch", "breakfast", "cheese", "lettuce", "recipe", "spices", "pilaf", "rice", "sarma", "dolma", "peach", "cook"],
    "Education": ["school", "graduation", "curriculum", "youth", "graduates", "children", "learn", "literacy", "education", "writing", "reading", "language", "teacher", "math", "teach", "students", "study", "university"],
    "Nationalities": ["armenian", "lebanese", "polish", "spanish", "slavic", "ukrainian", "turkish", "turkic", "syrian", "azeri", "azerbaijani", "israeli", "palestinian", "russian", "assyrians", "french", "german", "greek", "european", "american", "iranian", "swedish", "swiss"],
    "Cities": ["gyumri", "boston", "damascus", "chicago", "ani", "van", "beirut", "aleppo", "yerevan", "Baku", "paris", "berlin", "moscow", "toronto", "watertown", "tel aviv", "jerusalem", "haifa", "nyc", "los angeles", "glendale"],
    "Countries": ["armenia", "spain", "poland", "lebanon", "ukraine", "syria", "turkey", "azerbaijan", "Azerbaijan", "canada", "australia", "israel", "palestine", "germany", "france", "russia", "china", "usa", "america", "artsakh", "georgia", "sweden", "greece", "abkhazia", "kosovo"],
    "Genocide": ["genocide", "deportees", "1915", "recognition", "Talaat Pasha", "massacres", "martyrs", "holocaust"],
    "Religion": ["church", "faith", "catholicos", "holiness", "prophet", "orthodox", "secular", "atheist", "lent", "diocese", "archbishop", "prayer", "Christianity", "Islam", "Christian", "Jew", "ministry", "Catholic", "apostolic", "Jesus", "Christ", "cross", "missionaries", "spiritual", "holy", "buddhist", "biblical", "worship", "gospel", "sermon", "reverend", "friar", "pastor", "evangelical","Christmas", "Easter", "Saint"],
    "Culture": ["poetry", "art", "exhibits", "heritage", "cultural", "books", "musuem", "painting", "sculpture", "culture", "literature", "dance", "music", "song", "book", "comedy", "performance", "humor"],
    "Blockade and Conflict": ["blockade", "trauma", "aggression", "ceasefire", "munitions", "bombs", "missile", "corridor", "military", "troops", "soldiers", "weapons", "conflict", "war", "battle", "peace", "shootings", "wounded", "starvation", "refugees"],
    "Social Issues": ["charity", "humanitarian", "aid", "volunteer", "fundraiser", "organizes", "awareness", "activism", "lobbying", "activists", "organizing", "lobbyist"],
    "Social Services": ["healthcare", "housing", "education", "public transit", "welfare", "subsidies", "childcare", "fares", "transporatation", "roads"],
    "Economic Terms": ["tax", "labor", "industry", "tech", "technology", "tourism", "manufacturing", "layoffs", "recession", "cuts", "employees", "unemployment", "unions", "workers", "gdp", "debt", "bank", "dram", "economic", "jobs", "employment", "spending", "budget", "dollar", "revenue", "productivity", "fund", "market", "economy", "economist", "economics", "trade"],
    "Gender Issues": ["menstrual", "girls", "family", "disability", "adoption", "infants", "babies", "teenagers", "gay", "lgbtq", "homophobia", "queer", "domestic", "pregnancy", "maternity", "feminine", "woman", "mother", "motherinlaw", "infertility"],
}

# Precompute category embeddings
category_vectors = {
    cat: model.encode(" ".join(words), convert_to_tensor=True)
    for cat, words in categories.items()
}

# Similarity threshold for categorizing a keyword
SIMILARITY_THRESHOLD = 0.34  

def categorize_keywords(keywords, category_vectors, similarity_threshold=SIMILARITY_THRESHOLD):
    """
    Categorizes keywords based on semantic similarity with category vectors.
    If similarity is below the threshold, the keyword will be categorized as "Uncategorized".
    """
    categorized = defaultdict(list)
    for keyword in keywords:
        keyword_vector = model.encode(keyword, convert_to_tensor=True)
        similarities = {
            cat: cosine_similarity(keyword_vector.cpu().numpy().reshape(1, -1),
                                   category_vector.cpu().numpy().reshape(1, -1))[0][0]
            for cat, category_vector in category_vectors.items()
        }
        best_match = max(similarities, key=similarities.get)
        best_similarity = similarities[best_match]
        
        # If the similarity is below the threshold, categorize as "Uncategorized"
        if best_similarity < similarity_threshold:
            categorized["Uncategorized"].append(keyword)
        else:
            categorized[best_match].append(keyword)
    
    return categorized

# Load input spreadsheet
input_file = "Normalized_Armenian_Womens_Articles.xlsx"
df = pd.read_excel(input_file)

# Check for "Keyword" column
if "Keywords" not in df.columns:
    raise ValueError("Input spreadsheet must contain a 'Keywords' column.")

keywords = df["Keywords"].dropna().tolist()  # Remove NaN values and convert to list

# Categorize keywords
categorized_keywords = categorize_keywords(keywords, category_vectors)

# Define output directory and ensure it exists
output_dir = "./output4/"
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
