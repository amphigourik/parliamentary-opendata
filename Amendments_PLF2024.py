import requests
import openpyxl
import html
import re
import concurrent.futures
from tqdm import tqdm

# Function to fetch JSON data for an article
def fetch_article_data(article_number):
    url = f"https://www.assemblee-nationale.fr/dyn/opendata/AMANR5L16PO791932B1680P1D1N{article_number:06}.json"
    response = requests.get(url)
    if response.status_code == 200:
        return article_number, response.json()
    else:
        return article_number, None

# Use concurrent.futures for parallelization
def main(end_article_number, start_article_number=1):
    uid_list = [n for n in range(start_article_number, end_article_number + 1)]

    json_data = {}

    with concurrent.futures.ThreadPoolExecutor() as executor:
        futures = {executor.submit(fetch_article_data, article_number): article_number for article_number in uid_list}

        # Create a progress bar with tqdm
        with tqdm(total=len(uid_list), desc="Fetching JSON Data") as pbar:
            for future in concurrent.futures.as_completed(futures):
                article_number, data = future.result()
                if data is not None:
                    json_data[article_number] = data
                pbar.update(1)

    return json_data

def get_value(data, key, default="N/A"):
    try:
        value = data[key]
        return re.sub('<.*?>', '', html.unescape(value)) if value else default
    except KeyError:
        return default

# Define a dictionary to map party IDs to party names
party_mapping = {
    "PO793087": "Non inscrit",
    "PO800484": "Démocrate (MoDem et Indépendants)",
    "PO800490": "La France insoumise - Nouvelle Union Populaire écologique et sociale",
    "PO800496": "Socialistes et apparentés (membre de l’intergroupe NUPES)",
    "PO800502": "Gauche démocrate et républicaine - NUPES",
    "PO800508": "Les Républicains",
    "PO800514": "Horizons et apparentés",
    "PO800520": "Rassemblement National",
    "PO800526": "Écologiste - NUPES",
    "PO800532": "Libertés, Indépendants, Outre-mer et Territoires",
    "PO800538": "Renaissance",
    # Add more party mappings if needed
}

# Function to extract and clean data
def extract_and_clean_data(json_data):
    data_list = []

    for article_number, data in json_data.items():
        id = get_value(data['identification'], 'numeroOrdreDepot')
        author = get_value(data['signataires'], 'libelle')

        # Replace party ID with party name using the party_mapping dictionary
        party_id = get_value(data['signataires']['auteur'], 'groupePolitiqueRef')
        party = party_mapping.get(party_id, "Party not found")  # Use "Party not found" as the default value
        
        article = get_value(data['pointeurFragmentTexte']['division'], 'articleDesignationCourte')
        arrangement = get_value(data['corps']['contenuAuteur'], 'dispositif')
        briefStatement = get_value(data['corps']['contenuAuteur'], 'exposeSommaire')

        # Check if "assurance" is in the "Dispositif" and create a boolean column
        hasAssurance = "assurance" in arrangement or "Assurance" in arrangement or "assurances" in arrangement or "Assurances" in arrangement or "assurance" in briefStatement or "Assurance" in briefStatement or "assurances" in briefStatement or "Assurances" in briefStatement

        data_list.append([id, author, party, article, hasAssurance, arrangement, briefStatement])

    return data_list


# Function to save data to an Excel file
def save_to_excel(data_list):
    # Sort the data by the "Amendement" number
    sorted_data = sorted(data_list, key=lambda x: int(x[0]))

    workbook = openpyxl.Workbook()
    worksheet = workbook.active

    # Add headers to the Excel sheet
    worksheet.append(["Amendement", "Auteur", "Parti", "Article", "Assurance", "Dispositif", "ExposéSommaire"])

    # Add sorted data rows to the Excel sheet
    for row in sorted_data:
        worksheet.append(row)

    # Save the Excel file with the collected data
    workbook.save("amendements_data.xlsx")

# Main execution
if __name__ == "__main__":
    # Define the range of article numbers
    start_article_number = 1
    end_article_number = 5116

    json_data = main(start_article_number=start_article_number, end_article_number=end_article_number)
    
    if json_data:
        data_list = extract_and_clean_data(json_data)
        save_to_excel(data_list)
        print("Data has been collected and saved to article_data.xlsx")
    else:
        print("Failed to retrieve JSON data for some articles. Check error handling for details.")
