import pandas as pd
import xmltodict
from metapub import PubMedFetcher
import unicodedata 
from fuzzywuzzy import fuzz
import datetime
import re
import streamlit as st
from io import BytesIO
from pyxlsb import open_workbook as open_xlsb
import os
import time

import shutil
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from bs4 import BeautifulSoup
from urllib.parse import quote
from selenium.webdriver.chrome.options import Options



# ------------------ Global Variables ------------------ #

prod = pd.DataFrame(columns=[
                            'pmid','title', 'authors','journal','year','pagination',
                            'volume','issue', 'day_when_published','month_when_published', 
                            'authors_full_name','affiliations', 'corresponging_author_email', 'epub', 'ciber', 'if_actual', 'quantile_actual',
                            'if_when_published', 'quantile_when_published', 'type', 'doi'
                            ]
                    )
# add last, fist, corresppnding (if any are last first corresponding)

names_df = pd.DataFrame()
jcr = pd.DataFrame()
articles = {}
articles_xml = {}
author_list = {}
pubdate = {}
epubdate = {}

today = time.strftime("%d/%m/%Y")
username = "milagros.mejia@vhir.org"
password = "12424Car!"

# ------------------ Functions ------------------ #
fetch = PubMedFetcher()

def intro():
    st.image('https://upload.wikimedia.org/wikipedia/commons/thumb/9/90/Vhir_logo.svg/1200px-Vhir_logo.svg.png', width=200)
    st.write("# Hola! 👋")
    st.sidebar.success("Escoge qué quieres hacer")
    st.markdown(
        """
        Esta aplicación permite extraer información de artículos científicos a partir de sus PMIDs y generar un registro de publicaciones y actualizar los valors de IF de un año concreto.
    """
    )
def upload_clicked():
    st.session_state.clicked = not st.session_state.clicked

def extract_articles_from_pmids(pmids):
    progress_text = "Extrayendo informacion de artículos"
    valid_pmids = []
    my_bar = st.progress(0, text=progress_text)
    percent_complete = 0
    for i, pmid in enumerate(pmids):
        my_bar.progress(percent_complete, text=progress_text)
        start_time = time.time()
        timer_expired = False  # Flag to indicate if timer has expired
        #extract articles
        try:
            articles[pmid] = fetch.article_by_pmid(pmid)
            timer_duration = 40  # 40 seconds

        except:
            while True:
                try:
                    articles[pmid] = fetch.article_by_pmid(pmid)
                    break
                except:
                    print("fetching ", pmid, end='\r')
                      # Check if the timer has expired
                    elapsed_time = time.time() - start_time
                    if elapsed_time >= timer_duration:
                        print("expired!")
                        st.write(f'Error al extraer el artículo con pmid {pmid}, por favor comprueba que el pmid es correcto')
                        timer_expired = True
                        break
        if timer_expired:
            # Timer expired continue to the next pmid
            continue

        #process extracted articles
        valid_pmids.append(pmid)
        articles_xml[pmid] = xmltodict.parse(articles[pmid].xml)
        authors =  articles_xml[pmid]['PubmedArticleSet']['PubmedArticle']['MedlineCitation']['Article'][ 'AuthorList']['Author']
        pubdate[pmid] = articles_xml[pmid]['PubmedArticleSet']['PubmedArticle']['MedlineCitation']['Article']['Journal']['JournalIssue']['PubDate']
        epubdate[pmid] = articles_xml[pmid]['PubmedArticleSet']['PubmedArticle']['MedlineCitation']['Article']
        if type(authors) == list: # indicates multiple authors
            author_list[pmid] = authors
        else: # only one author
            author_list[pmid] = [authors] #set to a 1 entry list
        
        percent_complete = int((i+1)*100/len(pmids))
    my_bar.progress(100, text=progress_text)
    my_bar.empty()
    return valid_pmids


def create_dataframe_from_articles(pmids):
    global current_year
    prod.pmid = pmids
    prod.title = [articles[pmid].title for pmid in pmids]
    prod.day_when_published = [pubdate[pmid]['Day'] if 'Day' in pubdate[pmid].keys() else None for pmid in pmids]
    prod.month_when_published = [pubdate[pmid]['Month'] if 'Month' in pubdate[pmid].keys() else None for pmid in pmids]
    prod.year = [articles[pmid].year for pmid in pmids]
    prod.journal = [articles[pmid].journal for pmid in pmids]
    prod.authors = ['; '.join([str(author['LastName'] + ' ' + author['Initials']) if 'ForeName' in author.keys() and 'Initials' in author.keys() else '' for author in author_list[pmid]]) for pmid in pmids]
    prod.authors_full_name = ['; '.join([str(author['LastName'] + ' ' + author['ForeName']) if 'ForeName' in author.keys() and 'LastName' in author.keys() else '' for author in author_list[pmid]]) for pmid in pmids]
    affiliations_list = [[author['AffiliationInfo'] if 'AffiliationInfo' in author.keys() else '' for author in author_list[pmid]] for pmid in pmids]
    affiliations_list = [list(filter(lambda x: x != '', affiliation)) for affiliation in affiliations_list]
    affiliations = [[affiliation['Affiliation'] if type(affiliation)==dict else '; '.join([a['Affiliation'] for a in affiliation]) for affiliation in affiliations] for affiliations in affiliations_list]
    prod.affiliations = ['; '.join(affiliation) for affiliation in affiliations]
    prod.doi = [articles[pmid].doi for pmid in pmids]
    pubtypes = [list(articles[pmid].publication_types.values()) for pmid in pmids]
    prod.type = ['; '.join(types) for types in pubtypes]
    prod.pagination = [str(articles[pmid].pages) for pmid in pmids]
    prod.issue = [articles[pmid].issue for pmid in pmids]
    prod.volume = [articles[pmid].volume for pmid in pmids]
    epub_year  = [epubdate[pmid]['ArticleDate']['Year'] if 'ArticleDate' in epubdate[pmid].keys() else None for pmid in pmids]
    epub_month = [epubdate[pmid]['ArticleDate']['Month'] if 'ArticleDate' in epubdate[pmid].keys() else None for pmid in pmids]
    epub_day = [epubdate[pmid]['ArticleDate']['Day'] if 'ArticleDate' in epubdate[pmid].keys() else None for pmid in pmids]
    epub = [list(date) for date in zip(epub_year, epub_month, epub_day)]
    prod.epub = ['-'.join(e) if e[0] != None and e[1] != None and e[2] != None else None for e in epub]
    prod.ciber = [1 if 'CIBER' in affiliation else 0 for affiliation in prod.affiliations]
    prod.if_when_published = [jcr[jcr['Revista'] == journal]['IF' + str(year)].values[0] if 'IF' + str(year) in jcr.columns and len(jcr[jcr['Revista'] == journal]['IF' + str(year)].values) > 0 else None for journal, year in zip(prod.journal, prod.year)]
    prod.quantile_when_published = [jcr[jcr['Revista'] == journal]['Q' + str(year)].values[0] if 'Q' + str(year) in jcr.columns and len(jcr[jcr['Revista'] == journal]['Q' + str(year)].values) > 0 else None for journal, year in zip(prod.journal, prod.year)]
    prod.if_actual = [jcr[jcr['Revista'] == journal]['IF' + str(current_year)].values[0] if 'IF' + str(current_year) in jcr.columns and len(jcr[jcr['Revista'] == journal]['IF' + str(current_year)].values) > 0 else None for journal in prod.journal]
    prod.quantile_actual = [jcr[jcr['Revista'] == journal]['Q' + str(current_year)].values[0] if 'Q' + str(current_year) in jcr.columns and len(jcr[jcr['Revista'] == journal]['Q' + str(current_year)].values) > 0 else None for journal in prod.journal]
    prod.corresponging_author_email = prod.affiliations.apply(lambda x: [get_email(text) for text in x.split('; ') if '@' in text])

def get_email(text):
    email = re.findall(r'[\w\.-]+@[\w-]+[\.\w-]+[^\.]', text)
    return email

def whose_email(email):
    for _, row in names_df.iterrows():
        if email in row['email']:
            return row['author_col']
def corresponding_author(emails):
    authors_list = []
    for email in emails: #[[], []]
        for e in email: # [.., ..]
            authors_list.append(whose_email(e))
    authors_list = list(set(authors_list))
    return authors_list

def strip_accents(string): 
    return "".join(c for c in unicodedata.normalize("NFD", string) if not unicodedata.combining(c)) 

def fuzzy_match_author(author_name, authors): #authors is the actual list of authors as it appears in the article
    for position, actual_author in enumerate(authors):
        ratio = fuzz.token_sort_ratio(author_name, actual_author)
        if ratio >= 70:  # Adjust the threshold as needed
            if position == 0:
                return 'first'
            if position == len(authors) - 1:
                return 'last'
            return '1'
    return '0'

def check_ciberesp(row):
    for _, name in names_df.iterrows():
        if name.ciberesp == 1:
            col = strip_accents(name.author_name).lower().replace(' ', '_')
            if row[col] == 1:
                return True
    return False
def check_cibercv(row):
    for _, name in names_df.iterrows():
        if name.cibercv == 1:
            col = strip_accents(name.author_name).lower().replace(' ', '_')
            if row[col] == 1:
                return True

def create_authors_columns(pmids):
    global prod
    names_df.email = names_df.email.apply(lambda row: row.replace(' ', '').split(',') if type(row) == str else [])
    names = [strip_accents(name).lower() for name in names_df.author_name.values] # full author names from VHIR
    name_cols = [name.replace(' ', '_') for name in names]
    names_df['author_col'] = name_cols
    cols = ['pmid', 'corresponding_authors', 'authors_full_name_normalized', 'ciberesp', 'cibercv', 'any_first', 'any_last', 'any_corresponding'] + name_cols
    authors_df = pd.DataFrame(columns = cols)
    authors_df.pmid = pmids
    authors_df.ciberesp = 0
    authors_df.cibercv = 0
    authors_df.authors_full_name_normalized = prod.authors_full_name
    authors_df.authors_full_name_normalized = authors_df.authors_full_name_normalized.apply(lambda row: strip_accents(row).lower().split('; '))
    authors_df.corresponding_authors = prod.corresponging_author_email.apply(lambda row: corresponding_author(row) if len(row) > 0 else [])
    for index , row in authors_df.iterrows():
        corresponding_authors = row.corresponding_authors
        for name, column in zip(names, name_cols):
            if len(corresponding_authors)> 0 and column in corresponding_authors:
                authors_df.at[index, column] = 'corresponding'
            else:
                result = fuzzy_match_author(name, row.authors_full_name_normalized) #VHIR author vs articles authors
                authors_df.at[index, column] = result
    authors_df.any_first = [1 if any([row[name] == 'first' for name in name_cols]) else 0 for _, row in authors_df.iterrows()]
    authors_df.any_last = [1 if any([row[name] == 'last' for name in name_cols]) else 0 for _, row in authors_df.iterrows()]
    authors_df.any_corresponding = [1 if any([row[name] == 'corresponding' for name in name_cols]) else 0 for _, row in authors_df.iterrows()]
    prod = prod.merge(authors_df, on='pmid') # merge authors columns with articles dataframe

def check_ciber():
    prod.ciberesp = prod.apply(lambda row: 1
                           if row.ciber == 1 and check_ciberesp(row) == True
                           else 0, axis=1)
    prod.cibercv= prod.apply(lambda row: 1
                            if row.ciber == 1 and check_cibercv(row) == True
                            else 0, axis=1)

@st.cache_data
def convert_pub(df):
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name='registros')
    workbook = writer.book
    worksheet = writer.sheets['registros']
    format1 = workbook.add_format({'num_format': '0'}) 
    worksheet.set_column('A:A', None, format1)  
    writer.close()
    processed_data = output.getvalue()
    return processed_data

@st.cache_data
def convert_if(if_df):
    global if_year
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        if_df.to_excel(writer, sheet_name=f"IF {if_year}", index=False)
        # data_frame2.to_excel(writer, sheet_name="Vegetables", index=False)
        # data_frame3.to_excel(writer, sheet_name="Baked Items", index=False)
    workbook = writer.book
    worksheet = writer.sheets[f"IF {if_year}"]
    format1 = workbook.add_format({'num_format': '0'}) 
    worksheet.set_column('A:A', None, format1)  
    writer.close()
    processed_data = output.getvalue()
    return processed_data

def downloaded():
    upload_clicked()

def save_results_publications():
    prod.drop(columns=['authors_full_name_normalized','authors_full_name','ciber', 'corresponding_authors'], inplace=True)
    xlm = convert_pub(prod)
    st.write('Resultados creados correctamente')
    st.download_button(
        label="Descargar resultados en Excel",
        data=xlm,
        file_name=f'registro_publicaciones_{time.strftime("%d")}-{time.strftime("%m")}-{time.strftime("%Y")}.xlsx',
        mime='text/xlsx',
        on_click = downloaded()
    )

def save_results_if(df):
    global if_year
    xlm = convert_if(df)
    st.write('Resultados creados correctamente')
    st.download_button(
        label="Descargar resultados en Excel",
        data= xlm,
        file_name=f'Impact_Factor_{if_year}.xlsx',
        mime='text/xlsx',
        on_click = downloaded()
    )

def create_dataframe(pmids_file, authors_file, jcr_file):
# READ FILES
    global jcr, names_df
    df = pd.read_excel(pmids_file, header = None)
    jcr = pd.DataFrame(pd.read_excel(jcr_file))
    names_df = pd.DataFrame(pd.read_excel(authors_file))

    #check if the first column contains data or not
    df.columns = ['pmids']
    if str(df.pmids.iloc[0]).lower() == 'pmids':
        df.drop(df.index[0], inplace=True)

    #convert to int and drop duplicates of pmids
    df = df.dropna(subset=['pmids']).drop_duplicates()
    df.pmids = df.apply(lambda row: int(row['pmids']), axis=1)
    pmids = [int(pmid) for pmid in df.pmids.values]

    valid_pmids = extract_articles_from_pmids(pmids)
    valid_pmids = list(set(valid_pmids))

    create_dataframe_from_articles(valid_pmids)

    create_authors_columns(valid_pmids) 

    check_ciber()

@st.cache_resource(show_spinner=False)

def login_to_website(username, password):
    base_url = "https://jcr.clarivate.com/jcr-jp/journal-profile"
    login_url = f"{base_url}/login"  # Update with the actual login page URL
  
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument("--headless")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--disable-dev-shm-usage")
    #chrome_options.add_argument("--window-size=1920,1080")
    
    # Initialize ChromeDriverManager with the desired version
    #chrome_driver_path = ChromeDriverManager(chrome_type='google').install()
    
    # Initialize the service with the path to ChromeDriver executable
    #service = Service(chrome_driver_path)
    
    # Initialize WebDriver with the provided options and service
    driver = webdriver.Chrome(options=chrome_options)#,service=service, )
    driver.get(login_url)
  
    # Wait for the login page to load, you may need to adjust the sleep time
    time.sleep(5)
    # Find the login form elements and enter credentials
    username_input = driver.find_element("css selector", "input[name='email']")  # replace with the actual CSS selector of the username field
    password_input = driver.find_element("css selector", "input[name='password']")  # replace with the actual CSS selector of the password field
    login_button = driver.find_element("css selector", "button[type='submit']")  # replace with the actual CSS selector of the login button

    username_input.send_keys(username)
    password_input.send_keys(password)
    login_button.click()

    # Wait for the login to complete, you may need to adjust the sleep time
    time.sleep(5)

    return driver  # Return the driver with the authenticated session

def get_impact_factor(driver, journal_name,if_year):
    journal_name_encoded = journal_name.upper()
    # print(journal_name_encoded)
    base_url = "https://jcr.clarivate.com/jcr-jp/journal-profile"
    # Search for the journal name
    search_url = f"{base_url}?journal={journal_name_encoded}&year={if_year}"
    save_url = f"{base_url}?journal={journal_name_encoded}&year=All%20years"

    # Use the existing driver (with authenticated session) to load the page
    driver.get(search_url)

    # Wait for the page to load, you may need to adjust the sleep time
    time.sleep(5)

    # Get the page source after JavaScript execution
    page_source = driver.page_source

    # Parse the HTML content
    soup = BeautifulSoup(page_source, 'html.parser')

    # Find the elements containing the impact factor and quantile
    impact_factor_element = soup.find('div', class_='col-sm-5 col-md-5 col-lg-5 jif-values')
    quantile_element = soup.find('tr', class_='tr-highlight ng-star-inserted')

    if impact_factor_element and quantile_element:
        impact_factor = impact_factor_element.find('p', class_='value').text.strip()
        quantile = quantile_element.find('td', class_='rbj-quartile')
        if not quantile:
            quantile = quantile_element.find('td', class_='indicator-quartile')
        quantile = quantile.text.strip()

        save_url = search_url #if found correct year, else all years
        return impact_factor, quantile
    else:
        return today, None
    
def run_scrapping(if_xlm):
    global if_year
    if_xlm = pd.read_excel(if_xlm, sheet_name=f'IF {if_year}')
    # link_xlm = pd.read_excel(if_xlm, sheet_name=f'LINK')
    authenticated_driver = login_to_website(username, password)
    if f'IF{if_year}' not in if_xlm.columns:
        if_xlm[f'IF{if_year}'] = None
        if_xlm[f'Q{if_year}'] = None
    progress_text = "Buscando artículos...."
    my_bar = st.progress(0, text=progress_text)
    percent_complete = 0
    for idx, row in if_xlm.iterrows():
        my_bar.progress(percent_complete, text=progress_text)
        if pd.isna(row[f'IF{if_year}']):
            impact_factor, quantile = get_impact_factor(authenticated_driver, row.Revista,if_year)
            
            colif = 'IF{}'.format(str(if_year))
            colq = 'Q{}'.format(str(if_year))
            if_xlm.at[idx, colif] = impact_factor
            if_xlm.at[idx, colq] = quantile
        percent_complete = int((idx+1)*100/len(if_xlm))
    my_bar.progress(100, text=progress_text)
    my_bar.empty()
    return if_xlm

def registro_publicaciones ():
    st.write("# Generar Registro de Publicaciones")
    global current_year
    current_year = st.selectbox(
        "Cual es el año de IF actual?",
        ("2030","2029","2028","2027","2026","2025","2024","2023","2022", "2021", "2020"),
        index=None,
        placeholder="Escoge un año",
    )
    if current_year:
        uploaded_file_pmids = st.file_uploader("PMIDS")
        uploaded_file_authors= st.file_uploader("NOMBRES DE LOS AUTORES")
        uploaded_file_jcr= st.file_uploader("IMPACT FACTOR")

        if 'clicked' not in st.session_state:
            st.session_state.clicked = False

        st.button('Extraer Información', on_click=upload_clicked)

        if st.session_state.clicked and uploaded_file_pmids is not None and uploaded_file_authors is not None and uploaded_file_jcr is not None:
            st.write('Archivos cargados correctamente')
            create_dataframe(uploaded_file_pmids,uploaded_file_authors,uploaded_file_jcr) 
            save_results_publications()

def actualizar_if():
    st.write("# Actualizar Impact Factor")
    global if_year
    if_year = st.selectbox(
        "Cual es el año que quieres añadir?",
        ("2030","2029","2028","2027","2026","2025","2024","2023","2022", "2021", "2020"),
        index=None,
        placeholder="Escoge un año",
    )
    if  if_year:
        uploaded_file_if = st.file_uploader("IMPACT FACTOR")
        if uploaded_file_if:
            if 'clicked' not in st.session_state:
                st.session_state.clicked = False
            st.button('Actualizar', on_click=upload_clicked)
            if st.session_state.clicked:
                st.write('Archivo cargado correctamente')
                new_if = run_scrapping(uploaded_file_if) 
                save_results_if(new_if)


page_names_to_funcs = {
    "Inicio": intro,
    "Actualizar Impact Factor": actualizar_if,
    "Generar Registro de Publicaciones": registro_publicaciones
}

demo_name = st.sidebar.selectbox("Choose a demo", page_names_to_funcs.keys())
page_names_to_funcs[demo_name]()
