import scholarly
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd

def fetch_paper_references(title):
    search_query = scholarly.search_pubs_query(title)
    paper = next(search_query)
    return paper

def create_google_sheet(paper):
    scope = ["https://spreadsheets.google.com/feeds",
             "https://www.googleapis.com/auth/drive"]
    credentials = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
    client = gspread.authorize(credentials)

    spreadsheet = client.create('Research Paper References')
    sheet = spreadsheet.sheet1

    headers = ['Title', 'Authors', 'Year', 'Abstract', 'Citations', 'URL']
    sheet.append_row(headers)

    row_data = [paper.bib.get('title', ''), paper.bib.get('author', ''), paper.bib.get('year', ''),
                paper.bib.get('abstract', ''), paper.citedby, paper.bib.get('url', '')]
    sheet.append_row(row_data)

    return spreadsheet.url

def export_to_excel(spreadsheet_url):
    scope = ["https://spreadsheets.google.com/feeds",
             "https://www.googleapis.com/auth/drive"]
    credentials = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
    client = gspread.authorize(credentials)

    spreadsheet_key = spreadsheet_url.split('/')[-1]
    spreadsheet = client.open_by_key(spreadsheet_key)
    sheet = spreadsheet.sheet1

    data = sheet.get_all_values()
    df = pd.DataFrame(data[1:], columns=data[0])

    excel_file = 'research_paper_references.xlsx'
    df.to_excel(excel_file, index=False)

    print(f'Research paper references exported to "{excel_file}" successfully.')

if __name__ == '__main__':
    title = input('Enter the title of the research paper: ')
    paper = fetch_paper_references(title)
    if paper:
        spreadsheet_url = create_google_sheet(paper)
        export_to_excel(spreadsheet_url)
    else:
        print('No research paper found for the given title.')