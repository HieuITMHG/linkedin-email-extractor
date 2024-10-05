from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException
from oauth2client.service_account import ServiceAccountCredentials
import gspread
import difflib
import time
import sys
import requests
import bs4


# ============== Driver config =================== #
options = webdriver.ChromeOptions()
options.add_argument('--no-sandbox')
options.add_argument("--disable-dev-shm-usage")
options.add_argument('--disable-gpu')
options.add_argument("--window-size=1920, 1200")
options.add_argument('--disable-dev-shm-usage')
driver = webdriver.Chrome(options=options)
# ================================================ #


# =============Your linkedin account============== # 
username = "Henry.Universes@TAHKfoundation.org"
password = "2024@ThanhddxHenry"
# ================================================ #                    

# =====================XPATH====================== #
            
NAME_HEADER = '/html/body/div[6]/div[3]/div/div/div[2]/div/div/main/section[1]/div[2]/div[2]/div[1]/div[1]/span[1]/a/h1'

# ================================================ #


# =================Login Functions================ #
def handle_cookie_acceptance(driver: webdriver.Chrome):
    try:
        driver.find_element(By.XPATH, "//button[span[text()='Accept']]").click()
        print("INFO: COOKIES IS ACCEPTED!")
    except:
        print("INFO: COOKIES IS NOT REQUIRED!")

def handle_code_verification(driver: webdriver.Chrome):
    try:
        # FIND VERIFICATION FIELD.
        ID_FIELD = "input__email_verification_pin"
        CONDITION = EC.presence_of_element_located((By.ID, ID_FIELD))
        verification_field = WebDriverWait(driver, 20).until(CONDITION)
        # FIND SUBMIT BUTTON.
        ID_FIELD = "email-pin-submit-button"
        CONDITION = EC.presence_of_element_located((By.ID, ID_FIELD))
        submit_button = WebDriverWait(driver, 20).until(CONDITION)
        # ENTER VERIFICATION CODE.
        code = input("Verification code required! Check your email and enter the code: ")
        verification_field.send_keys(code)
        time.sleep(1)
        submit_button.click()
        time.sleep(2)
    except:
        print("INFO: NO VERIFICATION DETECTED!")

def login(driver: webdriver.Chrome, username, password):
    try:
        driver.get("https://www.linkedin.com/login")
        # WAIT FOR LOADING PAGE.
        XPATH_USERNAME, XPATH_PASSWORD = '//*[@id="username"]', '//*[@id="password"]'
        XPATH_LOGIN_BUTTON = '//*[@id="organic-div"]/form/div[3]/button'
        username_field = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, XPATH_USERNAME)))
        password_field = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, XPATH_PASSWORD)))
        login_button = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, XPATH_LOGIN_BUTTON)))
        # ENTER USERNAME.
        username_field.send_keys(username)
        time.sleep(2)
        # ENTER PASSWORD.
        password_field.send_keys(password)
        time.sleep(2)
        # CLICK LOGIN BUTTON.
        login_button.click()
    except TimeoutException:
        raise Exception("ERROR: ELEMENT NOT FOUND!")
    except:
        raise Exception("ERROR: LOGIN FAILED!")
    # CHECK VERIFY.
    handle_code_verification(driver)
    handle_cookie_acceptance(driver)
    time.sleep(5)
# ================================================= #

# ==============Google sheet config================ #
SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"]

url_sheet ='https://docs.google.com/spreadsheets/d/1AwJWQ6BBCyOfD2m180jMw6JlhWQp2RR2jsu0Y9kbd8Y/edit?gid=0#gid=0'

worksheet = None

try :
  json_keyfile = 'credentials.json'
  scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
  credentials = ServiceAccountCredentials.from_json_keyfile_name(json_keyfile, scope)
except Exception as e:
  print('miss authentication file, please add file')

try:
    gc = gspread.authorize(credentials)
    sh = gc.open_by_url(url_sheet)
    worksheet = sh.get_worksheet(0)
except Exception as e:
    print(str(e))
# =================================================== #

# =================Get Company Domain================ #
def get_url_company(driver,key_word):
  driver.get('https://www.google.com/search?&hl=en')
  driver.find_element(By.NAME, "q").send_keys(key_word + " " + "website" + Keys.ENTER)
  time.sleep(3)
  first_div_element = driver.find_elements(By.XPATH,'//div[@class="yuRUbf"]')
  if first_div_element:
    best_match_link = None
    best_match_ratio = 0
    # Lặp qua từng phần tử <div>
    for div_element in first_div_element:
      span_elements = div_element.find_elements(By.TAG_NAME, 'a')
      if span_elements:
        for span in span_elements:
          link = span.get_attribute('href')
          if 'linked' not in link and 'facebook'not in link and 'wiki'not in link and 'wmb'not in link :
            # Tính toán tỷ lệ tương tự giữa từ trong tên công ty và từ trong liên kết
            similarity_ratio = difflib.SequenceMatcher(None, key_word.lower(), link.lower()).ratio()
            if len(link) <= 40 and similarity_ratio > best_match_ratio:
                best_match_link = link
                best_match_ratio = similarity_ratio
  return best_match_link
# ==================================================== #

def find_valid_sublist(lst):
    for sublist in lst:
        # Kiểm tra nếu sublist có ít nhất 19 phần tử
        if len(sublist) >= 19:
            # Lấy 19 phần tử đầu tiên của sublist
            first_19_elements = sublist[:19]
            # Kiểm tra nếu tất cả 19 phần tử đầu tiên là "" hoặc " "
            if all(elem == "" or elem == " " for elem in first_19_elements):
                # Nếu có phần tử thứ 20, nó không được là "" hoặc " "
                if len(sublist) > 19 and (sublist[19] == "" or sublist[19] == " "):
                    continue
                return sublist
    # Nếu không tìm thấy sublist phù hợp, trả về None
    return None

# ======================= Main ====================== #

column_names = worksheet.row_values(1)
Linkedin_column_index = None
Email_column_index = None
linkedin_links = None

for i, column_name in enumerate(column_names):
    if column_name == "LinkedIn":
        Linkedin_column_index = i + 1
    if column_name == "Email":
        Email_column_index = i + 1  

if Linkedin_column_index is not None:
    linkedin_links = worksheet.col_values(Linkedin_column_index)[1:]
else:
    print("Linkedin Column not found")
    sys.exit()

login(driver, username, password)

temp = 1
for link in linkedin_links:
    temp +=1
    print(link)
    driver.get(link)
    try:
        username = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, NAME_HEADER))).get_attribute('innerHTML')
        company_name = None
        try:
            WebDriverWait(driver, 10).until(
                EC.presence_of_all_elements_located((By.CLASS_NAME, "artdeco-list__item"))
            )
        except:
            print("Không tìm thấy thẻ artdeco-list__item")

        profile_html=driver.page_source
        soup=bs4.BeautifulSoup(profile_html,'html.parser')
        lis = soup.find_all('li', class_='artdeco-list__item')
        ls = []
        for x in lis:
            ls.append(x.getText().split('\n'))

        vls = find_valid_sublist(ls)
            
        try:
            company_name = vls[25]
            if company_name == "" or company_name == " " or ("yrs" in company_name):
                company_name = vls[19]
            if company_name == "" or company_name == " " or ("yrs" in company_name):
                company_name = vls[26]
        except:
            pass

        print(company_name)
        company_url = get_url_company(driver,company_name)
        print(company_url)
        r = requests.get(url=f"https://api.hunter.io/v2/email-finder?domain={company_url}///&full_name={username}&api_key=cf4ce56d1b356b60d43db4465a4db1c56e182ad0")
        email = r.json()['data']['email']
        
        if email:
            worksheet.update_cell(temp, Email_column_index, email)
        else:
            worksheet.update_cell(temp, Email_column_index, '')
    except:
        continue







