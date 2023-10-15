import requests
import openpyxl
import html
import re
import concurrent.futures
from tqdm import tqdm

# Define the range of article numbers
start_article_number = 1
end_article_number = 5116

# Function to fetch JSON data for an article
def fetch_article_data(article_number):
    url = f"https://www.assemblee-nationale.fr/dyn/opendata/AMANR5L16PO791932B1680P1D1N{article_number:06}.json"
    response = requests.get(url)
    if response.status_code == 200:
        return article_number, response.json()
    else:
        return article_number, None

# Use concurrent.futures for parallelization
def main():
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

# Function to extract and clean data
def extract_and_clean_data(json_data):
    data_list = []

    for article_number, data in json_data.items():
        id = get_value(data['identification'], 'numeroOrdreDepot')
        author = get_value(data['signataires'], 'libelle')
        party = get_value(data['signataires']['auteur'], 'groupePolitiqueRef')
        title = get_value(data['pointeurFragmentTexte']['division'], 'articleDesignationCourte')
        type = get_value(data['pointeurFragmentTexte']['division'], 'type')
        newArticle = get_value(data['pointeurFragmentTexte']['division'], 'articleAdditionnel')
        newChapter = get_value(data['pointeurFragmentTexte']['division'], 'chapitreAdditionnel')
        arrangement = get_value(data['corps']['contenuAuteur'], 'dispositif')
        briefStatement = get_value(data['corps']['contenuAuteur'], 'exposeSommaire')

        data_list.append([id, author, party, title, type, newArticle, newChapter, arrangement, briefStatement])

    return data_list


# Function to save data to an Excel file
def save_to_excel(data_list):
    workbook = openpyxl.Workbook()
    worksheet = workbook.active

    # Add headers to the Excel sheet
    worksheet.append(["Amendement", "Auteur", "Parti", "Titre", "Type", "ArticleAdditionnel", "ChapitreAdditionnel", "Dispositif", "ExposéSommaire"])

    # Add data rows to the Excel sheet
    for row in data_list:
        worksheet.append(row)

    # Save the Excel file with the collected data
    workbook.save("article_data.xlsx")

# Main execution
if __name__ == "__main__":
    json_data = main()
    
    if json_data:
        data_list = extract_and_clean_data(json_data)
        save_to_excel(data_list)
        print("Data has been collected and saved to article_data.xlsx")
    else:
        print("Failed to retrieve JSON data for some articles. Check error handling for details.")
