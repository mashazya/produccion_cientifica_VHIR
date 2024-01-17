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
import time


# ------------------ Global Variables ------------------ #

prod = pd.DataFrame(columns=[
                            'pmid','title','day','month', 'year', 'journal','authors', 
                            'authors_full_name','affiliations', 'email', 'doi','type','pagination',
                            'volume','issue','group', 'epub', 'ciber', 'if_actual',
                            'if_when_published', 'quantile_actual', 'quantile_when_published'
                            ]
                    )
names_df = pd.DataFrame()
jcr = pd.DataFrame()
articles = {}
articles_xml = {}
author_list = {}
pubdate = {}
epubdate = {}

today = datetime.date.today()
current_year = str(today.year)
current_year = '2022' # CHANGE !!!!!!

# ------------------ Functions ------------------ #
fetch = PubMedFetcher()

def upload_clicked():
    st.session_state.clicked = True

def extract_articles_from_pmids(pmids):
    progress_text = "Extrayendo informacion de artículos"
    my_bar = st.progress(0, text=progress_text)
    percent_complete = 0
    for i, pmid in enumerate(pmids):
        my_bar.progress(percent_complete, text=progress_text)
        #extract articles
        try:
            articles[pmid] = fetch.article_by_pmid(pmid)
        except:
            while True:
                try:
                    articles[pmid] = fetch.article_by_pmid(pmid)
                    break
                except:
                    print("fetching ", pmid, end='\r')
        #process extracted articles
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
    st.write('Artículos extraidos correctamente')


def create_dataframe_from_articles(pmids):
    prod.pmid = pmids
    prod.title = [articles[pmid].title for pmid in pmids]
    prod.day = [pubdate[pmid]['Day'] if 'Day' in pubdate[pmid].keys() else None for pmid in pmids]
    prod.month = [pubdate[pmid]['Month'] if 'Month' in pubdate[pmid].keys() else None for pmid in pmids]
    prod.year = [articles[pmid].year for pmid in pmids]
    prod.journal = [articles[pmid].journal for pmid in pmids]
    prod.authors = ['; '.join([str(author['LastName'] + ' ' + author['Initials']) if 'ForeName' in author.keys() and 'Initials' in author.keys() else '' for author in author_list[pmid]]) for pmid in pmids]
    prod.authors_full_name = ['; '.join([str(author['LastName'] + ' ' + author['ForeName']) if 'ForeName' in author.keys() and 'LastName' in author.keys() else '' for author in author_list[pmid]]) for pmid in pmids]
    affiliations = [[author['AffiliationInfo']['Affiliation'] if 'AffiliationInfo' in author.keys() else '' for author in author_list[pmid]] for pmid in pmids]
    affiliations = [list(filter(lambda x: x != '', affiliation)) for affiliation in affiliations]
    prod.affiliations = ['; '.join(affiliation) for affiliation in affiliations]
    prod.doi = [articles[pmid].doi for pmid in pmids]
    pubtypes = [list(articles[pmid].publication_types.values()) for pmid in pmids]
    prod.type = ['; '.join(types) for types in pubtypes]
    prod.pagination = [str(articles[pmid].pages) for pmid in pmids]
    prod.issue = [articles[pmid].issue for pmid in pmids]
    prod.volume = [articles[pmid].volume for pmid in pmids]
    groups = [list(filter(lambda entry: 'CollectiveName' in entry.keys(), author_list[pmid])) for pmid in pmids]
    prod.group = ['; '.join([entry['CollectiveName'] for entry in group]) for group in groups]
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
    prod.email = prod.affiliations.apply(lambda x: [get_email(text) for text in x.split('; ') if '@' in text])

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
    st.write('Identificando autores')
    names_df.email = names_df.email.apply(lambda row: row.replace(' ', '').split(',') if type(row) == str else [])
    names = [strip_accents(name).lower() for name in names_df.author_name.values] # full author names from VHIR
    name_cols = [name.replace(' ', '_') for name in names]
    names_df['author_col'] = name_cols
    cols = ['pmid', 'corresponding_authors', 'authors_full_name_normalized', 'ciberesp', 'cibercv'] + name_cols
    authors_df = pd.DataFrame(columns = cols)
    authors_df.pmid = pmids
    authors_df.ciberesp = 0
    authors_df.cibercv = 0
    authors_df.authors_full_name_normalized = prod.authors_full_name
    authors_df.authors_full_name_normalized = authors_df.authors_full_name_normalized.apply(lambda row: strip_accents(row).lower().split('; '))
    authors_df.corresponding_authors = prod.email.apply(lambda row: corresponding_author(row) if len(row) > 0 else [])
    for index , row in authors_df.iterrows():
        corresponding_authors = row.corresponding_authors
        for name, column in zip(names, name_cols):
            if len(corresponding_authors)> 0 and column in corresponding_authors:
                authors_df.at[index, column] = 'corresponding'
            else:
                result = fuzzy_match_author(name, row.authors_full_name_normalized) #VHIR author vs articles authors
                authors_df.at[index, column] = result

    prod = prod.merge(authors_df, on='pmid') # merge authors columns with articles dataframe
    prod.corresponding_authors = prod.corresponding_authors.apply(lambda row: 1 if len(row) > 0 else 0)

def check_ciber():
    prod.ciberesp = prod.apply(lambda row: 1
                           if row.ciber == 1 and check_ciberesp(row) == True
                           else 0, axis=1)
    prod.cibercv= prod.apply(lambda row: 1
                            if row.ciber == 1 and check_cibercv(row) == True
                            else 0, axis=1)

@st.cache
def convert_df(df):
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name='registros')
    workbook = writer.book
    worksheet = writer.sheets['registros']
    format1 = workbook.add_format({'num_format': '0.00'}) 
    worksheet.set_column('A:A', None, format1)  
    writer.close()
    processed_data = output.getvalue()
    return processed_data

def save_results():
    prod.drop(columns=['authors_full_name_normalized','ciber', 'email'], inplace=True)
    xlm = convert_df(prod)

    st.download_button(
        label="Descargar resultados en Excel",
        data=xlm,
        file_name=f'results/registro_publicaciones_{today.day}-{today.month}-{today.year}.xlsx',
        mime='text/xlsx',
    )
    st.write('Resultados guardados correctamente')


def create_dataframe(pmids_file, authors_file, jcr_file):
# READ FILES
    global jcr, names_df
    df = pd.DataFrame(pd.read_excel(pmids_file))
    jcr = pd.DataFrame(pd.read_excel(jcr_file))
    names_df = pd.DataFrame(pd.read_excel(authors_file))


    df = df.dropna(subset=['pmids'])
    df.pmids = df.apply(lambda row: int(row['pmids']), axis=1)
    pmids = df.pmids.values[:10]

    extract_articles_from_pmids(pmids)

    create_dataframe_from_articles(pmids)

    create_authors_columns(pmids) 

    check_ciber()

    st.write('Autores identificados correctamente')

if __name__ == "__main__":
    uploaded_file_pmids = st.file_uploader("Carga el archivo de pmids")
    uploaded_file_authors= st.file_uploader("Carga el archivo de autores")
    uploaded_file_jcr= st.file_uploader("Carga el archivo de revistas")

    if 'clicked' not in st.session_state:
        st.session_state.clicked = False

    st.button('Cargar Archivos', on_click=upload_clicked)

    if st.session_state.clicked and uploaded_file_pmids is not None and uploaded_file_authors is not None and uploaded_file_jcr is not None:
        st.write('Archivos cargados correctamente')
        create_dataframe(uploaded_file_pmids,uploaded_file_authors,uploaded_file_jcr) 
        save_results()