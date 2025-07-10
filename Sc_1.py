from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd
import time
from datetime import datetime

# WebDriver का पथ सेट करें (Selenium 4.x के लिए Service का उपयोग करें)
#service = Service(ChromeDriverManager().install())
#driver = webdriver.Chrome(service=service)
driver = webdriver.Chrome()
# वेबसाइट को ओपन करें
driver.get("https://pacsup.nectarinfotel.in/scp/login.php")

# लॉगिन क्रेडेंशियल्स दर्ज करें
driver.find_element(By.ID, "name").send_keys("NLPSV.up")
driver.find_element(By.ID, "pass").send_keys("abcd@1234")
driver.find_element(By.CLASS_NAME, "submit").click()

# पोर्टल पर रीडायरेक्ट होने तक प्रतीक्षा करें
WebDriverWait(driver, 20).until(EC.url_to_be("https://pacsup.nectarinfotel.in/scp/index.php"))

# एक्सेल शीट से डेटा लोड करें
input_file = "ticket.xlsx"  # अपनी एक्सेल फाइल का नाम यहाँ डालें
data = pd.read_excel(input_file)

# सुनिश्चित करें कि "Ticket Number" कॉलम मौजूद है
if "Ticket Number" not in data.columns:
    raise Exception("Excel file must contain a column named 'Ticket Number'.")

# परिणामों को स्टोर करने के लिए नई कॉलम बनाएं
data["NLPSV Ticket ID"] = ""  # Renamed from "Result" to "NLPSV Ticket ID"
data["District"] = ""
data["PACS Name"] = ""  # New column for field_42
data["ERP ID"] = ""  # New column for field_40
data["Help Topic/Module"] = ""  # New column for field_76
data["Subject"] = ""  # New column for clear tixTitle has_bottom_border
data["Phase 1 or 2"] = ""  # New column for field_55
data["Issue Details"] = ""  # New column for field_53

# हर टिकट नंबर के लिए डेटा प्रोसेस करें
for index, row in data.iterrows():
    ticket_number = row["Ticket Number"]
    
    # *स्किप करें यदि टिकट नंबर पूर्णांक नहीं है*
    if not str(ticket_number).isdigit():
        print(f"Skipping non-numeric ticket number: {ticket_number}")
        data.at[index, "NLPSV Ticket ID"] = "Invalid Ticket Number"  # Updated column name
        continue  # अगले रिकॉर्ड पर जाएं
    
    try:
        # सर्च बार में टिकट नंबर दर्ज करें
        search_field = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.CLASS_NAME, "basic-search")))
        search_field.clear()
        search_field.send_keys(str(ticket_number))
        search_field.send_keys(Keys.RETURN)
        
        # सर्च परिणाम लोड होने तक प्रतीक्षा करें
        time.sleep(10)  # कुछ समय प्रतीक्षा करें (डायनामिक लोडिंग को संभालने के लिए)
        
        # "preview" बटन का इंतजार करें और क्लिक करें
        preview_button = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.CLASS_NAME, "preview"))
        )
        preview_button.click()
        #time.sleep(1)
        
        # फील्ड्स से डेटा स्क्रैप करें
        issue_details = WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.ID, "field_53"))).text
        help_topic_module = WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.ID, "field_76"))).text
        erp_id = WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.ID, "field_40"))).text
        district = WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.ID, "field_39"))).text
        nlpsv_ticket_id = WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.ID, "field_54"))).text
        
        # नए फील्ड्स को स्क्रैप करें
        pacs_name = WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.ID, "field_42"))).text
        phase = WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.ID, "field_55"))).text
        subject = WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.CLASS_NAME, "clear.tixTitle.has_bottom_border"))).text
        
        # "Back to Home" या पोर्टल के होमपेज पर रीडायरेक्ट करें
        driver.get("https://pacsup.nectarinfotel.in/scp/index.php")
        
    except Exception as e:
        issue_details = "No data found"
        help_topic_module = "No data found"
        erp_id = "No data found"
        district = "No data found"
        nlpsv_ticket_id = "No data found"  # Updated column name
        pacs_name = "No data found"
        phase = "No data found"
        subject = "No data found"
    
    # डेटा को एक्सेल में स्टोर करें
    data.at[index, "NLPSV Ticket ID"] = nlpsv_ticket_id  # Updated column name
    data.at[index, "District"] = district
    data.at[index, "PACS Name"] = pacs_name  # Store PACS Name
    data.at[index, "ERP ID"] = erp_id  # Store ERP ID
    data.at[index, "Help Topic/Module"] = help_topic_module  # Store Help Topic/Module
    data.at[index, "Subject"] = subject  # Store Subject
    data.at[index, "Phase 1 or 2"] = phase  # Store Phase
    data.at[index, "Issue Details"] = issue_details  # Store Issue Details

# आउटपुट फाइल का नाम सेट करें (आज की तारीख और समय के साथ)
current_datetime = datetime.now().strftime("%d-%m-%Y_%H-%M-%S")  # Corrected date format
output_file = f"updated_ticket_{current_datetime}.xlsx"

# अपडेटेड डेटा को एक्सेल फाइल में सेव करें
try:
    data.to_excel(output_file, index=False)
    print(f"Scraped data saved to '{output_file}'.")
except Exception as e:
    print(f"Error saving data to Excel file: {e}")

# ब्राउज़र को बंद करें
driver.quit()