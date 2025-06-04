import unicodedata
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException, TimeoutException
import pandas as pd
from webdriver_manager.chrome import ChromeDriverManager


def standardize_string(s):
    s = unicodedata.normalize('NFKC', s).strip().lower()
    s = s.replace('–', '-').replace('—', '-')
    return s


df = pd.read_excel('Substances.xlsx')

raw_properties = [
    'adrenergic', 'Adrenergic', 'beta-2 agonist', 'beta-2 adrenergic agonist', 'amphetamine',
    'anabolic', 'Anabolic', 'anaesthetic', 'analgesic', 'Analgesic', 'androgen', 'androgenic',
    'Anesthetic', 'anesthetic', 'antianxiety', 'Antianxiety', 'anticonvulsant', 'antidepressant',
    'Anti-inflammatory', 'anti-inflammatory', 'antihistamine', 'antioxidant', 'antiphlogistic',
    'antipsychotic', 'antitussive', 'Anxiolytics', 'barbiturate', 'Barbiturates', 'benzodiazepine',
    'Benzodiazepine', 'bronchodilator', 'calming effect', 'cannabinoid', 'cathinone', 'cathine',
    'cell growth', 'controlled substance', 'corticosteroid', 'Depressant', 'dilation', 'diuretic',
    'Diuretic', 'erythropoietin', 'estrogen', 'fatty acid oxidation', 'gamma-aminobutyric acid',
    'GHRH', 'glucocorticoid', 'growth hormone', 'growth hormone-releasing factor', 'growth hormone',
    'releasing peptide', 'hallucinogenic', 'hemorrhage', 'Hypnotic', 'hypoxia-inducible factor',
    'metabolic modulator', 'mucolytic', 'Muscarinic', 'Muscle relaxants', 'narcotic', 'NSAID',
    'Opiate', 'opioid', 'progestin', 'Psycholeptics', 'quieting effect', 'radical scavenger',
    'SARM', 'Schedule I', 'Schedule II', 'Schedule III', 'sedation', 'sedative', 'Sedative',
    'selective androgen receptor modulator', 'steroid', 'Steroids', 'stimulant', 'stimulants',
    'Stimulants', 'stimulating', 'stimulation', 'tranquilizer', 'vasodilator', 'non-steroidal', 'non-narcotic'
]
PROPERTIES = sorted(set([p.lower() for p in raw_properties]))

for prop in PROPERTIES:
    df[prop] = 'FALSE'

df.insert(df.columns.get_loc('adrenergic'), 'Summary', 'Non-doping agent')


def init_driver():
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument('--disable-gpu')
    chrome_options.add_argument('--headless')
    service = Service(ChromeDriverManager().install())
    # service = Service("C:/Users/ruixi/Downloads/chromedriver-win64/chromedriver.exe")
    driver = webdriver.Chrome(service=service, options=chrome_options)
    driver.implicitly_wait(10)
    return driver


def standardize_compound_name(name):
    return name.replace('[', '(').replace(']', ')')


def click_and_verify(driver, compound_name, max_attempts=3):
    special_compounds = {
        "N-Benzylthieno[2,3-d]pyrimidin-4-amine": True,
        "[3-Chloro-4-(isopropylcarbamoyl)phenyl]boronic acid": True
    }

    is_special = compound_name in special_compounds
    standardized_name = standardize_compound_name(compound_name)

    print(f"processing compound: {compound_name}")

    for attempt in range(max_attempts):
        try:
            WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.ID, "results-container"))
            )

            if is_special:
                try:
                    result_links = driver.find_elements(By.CSS_SELECTOR, "ul.unstyled-list > li:first-child a")
                    if result_links:
                        print(f"special compound, click on the first result: {result_links[0].text}")
                        result_links[0].click()
                        WebDriverWait(driver, 20).until(
                            EC.presence_of_element_located((By.XPATH, "//*[contains(text(), 'Structure')]"))
                        )
                        return True
                except Exception as e:
                    print(f"error dealing with special case: {e}")

            result_links = WebDriverWait(driver, 20).until(
                EC.presence_of_all_elements_located((By.CSS_SELECTOR, "a[href*='/compound/']"))
            )

            print(f"find {len(result_links)} results")
            for i, link in enumerate(result_links[:5]):
                print(f"link {i + 1}: {link.text}")

            for link in result_links:
                link_text = standardize_string(link.text)
                std_name = standardize_string(standardized_name)

                if std_name in link_text or link_text in std_name:
                    print(f"match: {link.text}")
                    link.click()
                    WebDriverWait(driver, 20).until(
                        EC.presence_of_element_located((By.XPATH, "//*[contains(text(), 'Structure')]"))
                    )
                    return True

            if result_links:
                print(f"did not match, click on the first result: {result_links[0].text}")
                result_links[0].click()
                WebDriverWait(driver, 20).until(
                    EC.presence_of_element_located((By.XPATH, "//*[contains(text(), 'Structure')]"))
                )
                return True

        except (TimeoutException, NoSuchElementException) as e:
            print(f"try {attempt + 1} error: {str(e)}")
            if attempt < max_attempts - 1:
                if is_special and attempt == 0:
                    try:
                        driver.get("https://pubchem.ncbi.nlm.nih.gov/")
                        search_box = WebDriverWait(driver, 10).until(
                            EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='text']"))
                        )

                        if "Benzylthieno" in compound_name:
                            search_term = "Benzylthieno pyrimidin"
                        elif "boronic acid" in compound_name.lower():
                            search_term = "Chloro phenyl boronic acid"
                        else:
                            search_term = compound_name

                        print(f"use abv: {search_term}")
                        search_box.send_keys(search_term + Keys.RETURN)
                    except Exception as se:
                        print(f"abv search error: {se}")
                continue

    print(f"tried all attempts but did not match: {compound_name}")
    return False


def scrape_pubchem(compound_name):
    driver = init_driver()
    try:
        compound_search = standardize_compound_name(compound_name)
        driver.get("https://pubchem.ncbi.nlm.nih.gov/")
        search_box = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='text']"))
        )
        search_box.send_keys(compound_search + Keys.RETURN)
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.ID, "results-container"))
        )
        if not click_and_verify(driver, compound_name):
            print(f"Failed to load the correct page for {compound_name}")
            return pd.Series({prop: 'ERROR' for prop in PROPERTIES} | {"Pharmacology": "Absent"})

        page_content = driver.page_source
        text_to_search = standardize_string(page_content)
        print(f"Debug: Text to search for {compound_name}: {text_to_search[:300]}")

        new_df = pd.Series()
        for prop in PROPERTIES:
            new_df[prop] = "TRUE" if standardize_string(prop) in text_to_search else "FALSE"
        new_df["Pharmacology"] = "Present" if "pharmacology" in text_to_search else "Absent"

        matched_props = [prop for prop, value in new_df.items() if value == "TRUE"]
        print(f"Matched properties for {compound_name}: {matched_props}")
        return new_df

    except Exception as e:
        print(f"Error encountered for {compound_name}: {e}")
        return pd.Series({prop: 'ERROR' for prop in PROPERTIES} | {"Pharmacology": "Absent"})
    finally:
        driver.quit()


classification_df = pd.read_excel('Keywords for classification.xlsx')
classification_dict = {}
steroid_keywords = set()
narcotic_keywords = set()
for col in classification_df.columns:
    for keyword in classification_df[col].dropna():
        keyword_lower = keyword.lower()
        classification_dict[keyword_lower] = col
        if col.lower() == 'steroid':
            steroid_keywords.add(keyword_lower)
        elif col.lower() == 'narcotics':
            narcotic_keywords.add(keyword_lower)


def classify_row(row):
    classifications = []
    steroid_keywords_found = [kw for kw in steroid_keywords if row.get(kw, 'FALSE') == 'TRUE']
    narcotic_keywords_found = [kw for kw in narcotic_keywords if row.get(kw, 'FALSE') == 'TRUE']
    if steroid_keywords_found and row.get('non-steroidal', 'FALSE') != 'TRUE':
        classifications.append('Steroid')
    if narcotic_keywords_found and row.get('non-narcotic', 'FALSE') != 'TRUE':
        classifications.append('Narcotics')
    for prop in PROPERTIES:
        if row[prop] == 'TRUE' and (prop not in steroid_keywords_found and prop not in narcotic_keywords_found):
            classification = classification_dict.get(prop, None)
            if classification and classification not in classifications:
                classifications.append(classification)
            if len(classifications) >= 8:
                break
    return classifications + [None] * (8 - len(classifications))


for index, row in df.iterrows():
    compound_name = row['Name']
    properties = scrape_pubchem(compound_name)
    if properties is None:
        print(f"Failed to scrape data for {compound_name}")
        continue
    for prop, value in properties.items():
        df.loc[index, prop] = value
    classifications = classify_row(df.loc[index])
    for i, classification in enumerate(classifications, start=1):
        df.loc[index, f'Classification {i}'] = classification
    doping_agent = any(
        df.loc[index, prop] == 'TRUE' for prop in PROPERTIES if prop not in ['non-steroidal', 'non-narcotic'])
    df.loc[index, 'Summary'] = 'Doping agent' if doping_agent else 'Non-doping agent'
    print(f"Updated DataFrame row for {compound_name}:")
    print(df.loc[index])

doping_count = (df['Summary'] == 'Doping agent').sum()
non_doping_count = (df['Summary'] == 'Non-doping agent').sum()
print(f"Total Doping agents: {doping_count}, Non-doping agents: {non_doping_count}")

classification_columns = [f'Classification {i}' for i in range(1, 9)]
all_columns = df.columns.tolist()
for col in classification_columns:
    all_columns.remove(col)
summary_index = all_columns.index('Summary')
for i, col in enumerate(classification_columns):
    all_columns.insert(summary_index + 1 + i, col)
df = df[all_columns]
print("Columns reordered:")
print(df.head())

print("Before sorting and saving, DataFrame looks like:")
print(df[['Name', 'Pharmacology']].head())
df['Error_Count'] = df[PROPERTIES].apply(lambda x: sum(x == 'ERROR'), axis=1)
df['False_Count'] = df[PROPERTIES].apply(lambda x: sum(x == 'FALSE'), axis=1)
df.sort_values(by=['Summary', 'Error_Count', 'False_Count'], ascending=[True, True, False], inplace=True)
print("After sorting, DataFrame looks like:")
print(df[['Name', 'Pharmacology']].head())
df.drop(columns=['Error_Count', 'False_Count'], inplace=True)
print("Final DataFrame before saving to Excel:")
print(df[['Name', 'Pharmacology']].head())
df.to_excel('Outputted_Pharmacological_Effects_with_Classifications_Sorted.xlsx', index=False)
print("file successfully saved to Outputted_Pharmacological_Effects_with_Classifications_Sorted.xlsx")