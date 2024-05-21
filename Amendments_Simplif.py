import requests
import openpyxl
import concurrent.futures
from tqdm import tqdm
from bs4 import BeautifulSoup, Comment
from re import search
from json import loads

# Function to extract a number from a URL
def extract_number_from_url(url):
    match = search(r'\d+', url)
    return int(match.group()) if match else 0

def extract_data_from_html(html_doc):
    soup = BeautifulSoup(html_doc, 'html.parser')

    fields = ['signataires', 'accordGouv', 'subdivision', 'dispositif', 'objet']

    data = {}

    for field in fields:
        # Find the debut and fin comments for the current field
        debut_comment = soup.find(text=lambda text: isinstance(text, Comment) and f'debut_{field}' in text)
        fin_comment = soup.find(text=lambda text: isinstance(text, Comment) and f'fin_{field}' in text)

        # Check if both comments were found
        if debut_comment and fin_comment:
            # Extract the data between the comments
            data_between = []
            sibling = debut_comment.next_sibling
            while sibling != fin_comment:
                data_between.append(str(sibling))
                sibling = sibling.next_sibling
            data[field] = ''.join(data_between)

    return data

def fetch_senators():
    url = "https://www.senat.fr/api-senat/senateurs.json"
    response = requests.get(url)
    if response.status_code == 200:
        data = loads(response.text)
        return data

# Function to fetch JSON data for an article
def fetch_article_data(amdt_url):
    # url = f"https://www.senat.fr/encommission/2023-2024/550/Amdt_COM-{article_id}.html"
    url = f"https://www.senat.fr/encommission/2023-2024/550/{amdt_url}"
    response = requests.get(url)
    if response.status_code == 200:
        data = extract_data_from_html(response.text)
        return amdt_url, data.get('signataires'), data.get('accordGouv'), data.get('subdivision'), data.get('dispositif'), data.get('objet')
    else:
        return amdt_url, None, None, None, None, None

# Use concurrent.futures for parallelization
def main(json_data):
    amdt_list = [amendment['urlAmdt'] for subdivision in json_data['Subdivisions'] for amendment in subdivision['Amendements']]
    sign_list = [amendment['urlAuteur'].split('.')[0][-6:].upper() for subdivision in json_data['Subdivisions'] for amendment in subdivision['Amendements']]
    senators = fetch_senators()
    
    html_data = {}

    with concurrent.futures.ThreadPoolExecutor() as executor:
        futures = {executor.submit(fetch_article_data, article_id): article_id for article_id in amdt_list}

        # Create a progress bar with tqdm
        with tqdm(total=len(amdt_list), desc="Fetching HTML Data") as pbar:
            for future in concurrent.futures.as_completed(futures):
                article_id, signataires, accordGouv, subdivision, dispositif, objet = future.result()
                
                for senator in senators:
                    if senator['matricule'] == sign_list[amdt_list.index(article_id)]:
                        party = senator['groupe']['libelle']
                        break

                # html_data[article_id] = (signataires, accordGouv, subdivision, dispositif, objet)
                soup = BeautifulSoup(dispositif, 'html.parser')
                dispositif_text = soup.get_text()
                
                soup = BeautifulSoup(objet, 'html.parser')
                objet_text = soup.get_text()
                
                html_data[article_id] = (signataires, accordGouv, subdivision, dispositif_text, objet_text, party)

    return html_data

# Function to save data to an Excel file
def save_to_excel(data_list):
    # Sort the data by the "Amendement" number
    sorted_data = sorted(data_list, key=lambda x: extract_number_from_url(x[0]))
    # add a boolean flag if article 14
    for row in sorted_data:
        if row[2] == "Article 14":
            row.append(True)
        else:
            row.append(False)
    # add a boolean flag if assurance is in the dispositif or objet
    for row in sorted_data:
        if "assurance" in row[3].lower() or "assurances" in row[3].lower() or "assurance" in row[4].lower() or "assurances" in row[4].lower():
            row.append(True)
        else:
            row.append(False)

    workbook = openpyxl.Workbook()
    worksheet = workbook.active

    # Add headers to the Excel sheet
    worksheet.append(["Amendement", "Auteur", "Article", "Signataires", "AccordGouv", "Subdivision", "Dispositif", "Objet", "Groupe", "ConcerneArticle14", "ContientAssurance"])

    # Add sorted data rows to the Excel sheet
    for row in sorted_data:
        worksheet.append(row)

    # Save the Excel file with the collected data
    workbook.save("simplif_data.xlsx")

# Main execution
if __name__ == "__main__":
    url = "https://www.senat.fr/encommission/2023-2024/550/liste_discussion.json"
    response = requests.get(url)
    json_data = response.json()

    if json_data:
        html_data = main(json_data)
        data_list = [[amendment['urlAmdt'], amendment['auteur'], subdivision['libelle_subdivision'], *html_data[amendment['urlAmdt']]] for subdivision in json_data['Subdivisions'] for amendment in subdivision['Amendements'] if amendment['urlAmdt'] in html_data]
       

        save_to_excel(data_list)
        print("Data has been collected and saved to amendements_data.xlsx")
    else:
        print("Failed to retrieve JSON data. Check error handling for details.")