# This is new_component_similarity_2024-12-13.py
import pandas as pd
from sklearn.feature_extraction.text import CountVectorizer
from sklearn.metrics.pairwise import cosine_similarity

# Input and output file paths
input_file = "ivntest.xlsx"
output_file = "filtered_new_components_similarity.xlsx"

# 1. Load the existing dataset
df = pd.read_excel(input_file)

# Display column names to guide the user
print("Column names in the Excel file:")
for col in df.columns:
    print(f"- {col}")

# Mandatory columns for comparison
enabling_columns = ["Enabling Component", "Enabling Component Description", "Enabling Source",
                    "Enabling Component URL", "Enabling Source Agency"]
dependent_columns = ["Dependent Component", "Dependent Component Description", "Dependent Source",
                     "Dependent Component URL", "Dependent Source Agency"]

# Check if required columns exist
missing_columns = set(enabling_columns + dependent_columns) - set(df.columns)
if missing_columns:
    raise ValueError(f"The input file is missing required columns: {', '.join(missing_columns)}")

# Fill missing values with an empty string
df = df.fillna("")

# Extract Enabling and Dependent components
enabling_df = df[enabling_columns].drop_duplicates()
dependent_df = df[dependent_columns].drop_duplicates()

# 2. Define new components with field names
new_components = [
    {
        "Source": "EO 14115 - Imposing Certain Sanctions on Persons Undermining Peace, Security, and Stability in the West Bank",
        "Component": "EO 14115 Section 1",
        "Component Description": "All property and interests in property that are in the United States, that hereafter come within the United States, or that are or hereafter come within the possession or control of any United States person...",
        "Component URL": "https://www.federalregister.gov/executive-order/14115",
        "Source Agency": "EOP"
    },
    {
        "Source": "EO 14116 - Amending Regulations Relating to the Safeguarding of Vessels, Harbors, Ports, and Waterfront Facilities of the United States",
        "Component": "EO 14116 Section 1",
        "Component Description": "This order amends regulations to strengthen safeguards for United States ports, harbors, and waterfront facilities...",
        "Component URL": "https://www.federalregister.gov/executive-order/14116",
        "Source Agency": "EOP"
    }
]

# Function to calculate cosine similarity
def calculate_similarity(text1, text2):
    if not text1 or not text2:  # Avoid processing empty text
        return 0
    vectorizer = CountVectorizer().fit_transform([text1, text2])
    vectors = vectorizer.toarray()
    return cosine_similarity(vectors)[0, 1]

# 3. Compare new components with existing Enabling and Dependent components
filtered_data = []

for new_component in new_components:
    new_source = new_component["Source"]
    new_component_name = new_component["Component"]
    new_description = new_component["Component Description"]
    new_url = new_component["Component URL"]
    new_agency = new_component["Source Agency"]

    # Compare as a Dependent Component against Enabling Components
    for enabling in enabling_df.itertuples(index=False):
        similarity = calculate_similarity(new_description, enabling[1])  # Compare descriptions
        if similarity >= 0.6:  # Threshold for filtering
            filtered_data.append({
                "Enabling Source": enabling[2],
                "Enabling Component": enabling[0],
                "Enabling Component Description": enabling[1],
                "Dependent Component": new_component_name,
                "Dependent Component Description": new_description,
                "Dependent Source": new_source,
                "Linkage mandated by what US Code or OMB policy?": "",
                "Enabling Component URL": enabling[3],
                "Dependent Component URL": new_url,
                "Enabling Source Agency": enabling[4],
                "Dependent Source Agency": new_agency,
                "Similarity": similarity
            })

    # Compare as an Enabling Component against Dependent Components
    for dependent in dependent_df.itertuples(index=False):
        similarity = calculate_similarity(new_description, dependent[1])  # Compare descriptions
        if similarity >= 0.6:  # Threshold for filtering
            filtered_data.append({
                "Enabling Source": new_source,
                "Enabling Component": new_component_name,
                "Enabling Component Description": new_description,
                "Dependent Component": dependent[0],
                "Dependent Component Description": dependent[1],
                "Dependent Source": dependent[2],
                "Linkage mandated by what US Code or OMB policy?": "",
                "Enabling Component URL": new_url,
                "Dependent Component URL": dependent[3],
                "Enabling Source Agency": new_agency,
                "Dependent Source Agency": dependent[4],
                "Similarity": similarity
            })

# 4. Convert results to a DataFrame and save to Excel
if filtered_data:
    output_columns = [
        "Enabling Source", "Enabling Component", "Enabling Component Description",
        "Dependent Component", "Dependent Component Description", "Dependent Source",
        "Linkage mandated by what US Code or OMB policy?",
        "Enabling Component URL", "Dependent Component URL",
        "Enabling Source Agency", "Dependent Source Agency", "Similarity"
    ]
    filtered_df = pd.DataFrame(filtered_data)

    # Ensure all required output columns exist
    for col in output_columns:
        if col not in filtered_df.columns:
            filtered_df[col] = ""  # Add empty column if missing

    # Reorder and save to file
    filtered_df = filtered_df[output_columns]
    filtered_df.to_excel(output_file, index=False)
    print(f"Filtered data has been saved to {output_file}")
else:
    print("No matches found with the specified similarity threshold.")



