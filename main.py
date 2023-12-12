import pandas as pd
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
import tkinter as tk
from tkinter import filedialog

def process_data(file_path):
    # Read data from Excel file
    df = pd.read_excel(file_path)
    # Assuming your Excel file has columns like 'upid' and 'id'
    for index, row in df.iterrows():
        if str(row['คำนำหน้า']) == 'จ.ส.ต.':
            choice_1 = '4' 
        if str(row['คำนำหน้า']) == 'นาง':
            choice_1 = '2' 
            
        if str(row['จังหวัด']) == 'สุรินทร์':
            choice_2 = '20' 
        if str(row['จังหวัด']) == 'นนทบุรี':
            choice_2 = '2' 
        
        if str(row['อำเภอ']) == 'เมืองสุรินทร์':
            choice_3 = '1' 
        if str(row['อำเภอ']) == 'ไทรน้อย':
            choice_3 = '1' 
            
        if str(row['แขวง/ตำบล']) == 'ในเมือง':
            choice_4 = '6' 
        if str(row['แขวง/ตำบล']) == 'ไทรน้อย':
            choice_4 = '1' 
            
        if str(row['หมู่ที่/หมู่บ้าน']) == 'ไม่ระบุ':
            choice_5 = '0'
        if str(row['หมู่ที่/หมู่บ้าน']) == '02 - บ้านเจ้าเฟื่อง':
            choice_5 = '2'
                    
        if str(row['อาชีพหลัก']) == 'รับเงินประเดือนประจำ':
            choice_6 = '1' 
        if str(row['อาชีพหลัก']) == 'รับจ้างทางการเกษตร':
            choice_6 = '2' 
            
        if str(row['อาชีพรอง']) == 'ประกอบอาชีพทางการเกษตร':
            choice_7 = '0'     
        if str(row['อาชีพรอง']) == 'ไม่มีอาชีพรอง':
            choice_7 = '6'     
            
        if str(row['โครงการ']) == 'แปลงใหญ่':
            choice_8 = '0'  
            
        if str(row['การถือครอง']) == 'พื่นที่ของตน':
            choice_9 = '0'  
        if str(row['การถือครอง']) == 'เช่า':
            choice_9 = '1'  
            
        if str(row['ปัญหา']) == 'ด้านทุน':
            choice_10 = '0'     
        if str(row['ปัญหา']) == 'ด้านแรงงาน':
            choice_10 = '1'     
        
            
        # URL of the form
        form_url = 'https://eregist.dld.go.th/register/addfarmer'
        driver = webdriver.Chrome()
        driver.maximize_window()
        driver.get(form_url)
        
        # รหัสประจำตัวประชาชน
        name_input = driver.find_element(By.ID, "upid")
        name_input.clear()
        name_input.send_keys(str(row['รหัสประจำตัวประชน']))

        # คำนำหน้า
        dropdown_element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "prefixName"))
        )

        dropdown_element.click()
        option_css_selector = f"#react-select-2-option-{choice_1}"
        
        dropdown_options = WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located((By.CSS_SELECTOR, option_css_selector))
        )
        if dropdown_options:
            dropdown_options[0].click()
            
        # ชื่อ
        input = driver.find_element(By.ID, "uname")
        input.clear()
        input.send_keys(str(row['ชื่อ'])) 
        
        # นามสกุล
        input = driver.find_element(By.ID, "ulname")
        input.clear()
        input.send_keys(str(row['สกุล']))       
        
        # โทรศัพท์ 
        input = driver.find_element(By.ID, "umobile")
        input.clear()
        input.send_keys(str(row['โทรศัพท์มือถือ']))      
        
        # อีเมล
        input = driver.find_element(By.ID, "uemail")
        input.clear()
        input.send_keys(str(row['อีเมล']))      
        
        
        # จังหวัด
        dropdown_element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "provinceName"))
        )

        dropdown_element.click()
        option_css_selector = f"#react-select-3-option-{choice_2}"
        
        dropdown_options = WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located((By.CSS_SELECTOR, option_css_selector))
        )
        if dropdown_options:
            dropdown_options[0].click()
            
        # อำเภอ
        dropdown_element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "amphurName"))
        )

        dropdown_element.click()
        option_css_selector = f"#react-select-4-option-{choice_3}"
        
        dropdown_options = WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located((By.CSS_SELECTOR, option_css_selector))
        )
        if dropdown_options:
            dropdown_options[0].click()
            
        # แขวง/ตำบล
        dropdown_element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "tambolName"))
        )

        dropdown_element.click()
        option_css_selector = f"#react-select-5-option-{choice_4}"
        
        dropdown_options = WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located((By.CSS_SELECTOR, option_css_selector))
        )
        if dropdown_options:
            dropdown_options[0].click()
            
        # หมู่ที่/หมู่บ้าน
        dropdown_element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "umoo"))
        )

        dropdown_element.click()
        option_css_selector = f"#react-select-6-option-{choice_5}"
        
        dropdown_options = WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located((By.CSS_SELECTOR, option_css_selector))
        )
        if dropdown_options:
            dropdown_options[0].click()
            
        # บ้านเลขที่
        input = driver.find_element(By.ID, "uhomeNo")
        input.clear()
        input.send_keys(str(row['บ้านเลขที่']))  
        
        # ตรอก/ซอย
        input = driver.find_element(By.ID, "usoi")
        input.clear()
        input.send_keys(str(row['ตรอก/ซอย']))  
        
        # ถนน
        input = driver.find_element(By.ID, "uroad")
        input.clear()
        input.send_keys(str(row['ถนน']))  
        
        # รหัสไปรษณีย์
        input = driver.find_element(By.ID, "postCode")
        input.clear()
        input.send_keys(str(row['รหัสไปรษณีย์']))  
        
        # เลขประจำตัวประชาชน
        input = driver.find_element(By.ID, "pid")
        input.clear()
        input.send_keys(str(row['รหัสประจำตัวประชน']))  
        
        # เลขประจำตัวประชาชน สมาชิกในครัวเรือน
        dropdown_element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "memPrefix"))
        )

        dropdown_element.click()
        option_css_selector = f"#react-select-7-option-{choice_1}"
        
        dropdown_options = WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located((By.CSS_SELECTOR, option_css_selector))
        )
        if dropdown_options:
            dropdown_options[0].click()
            
        # ชื่อ
        input = driver.find_element(By.ID, "latitude")
        input.clear()
        input.send_keys(str(row['ชื่อ'])) 
        
        # นามสกุล
        input = driver.find_element(By.ID, "memLastName")
        input.clear()
        input.send_keys(str(row['สกุล']))    
        
        
        # อาชีพหลัก
        driver.execute_script("window.scrollTo(0, 0);")
        dropdown_element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "mainJobsName"))
        )

        dropdown_element.click()
        option_css_selector = f"#react-select-8-option-{choice_6}"
        
        dropdown_options = WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located((By.CSS_SELECTOR, option_css_selector))
        )
        if dropdown_options:
            dropdown_options[0].click()
                
        # อาชีพรอง
        dropdown_element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "secondJobsName"))
        )

        dropdown_element.click()
        option_css_selector = f"#react-select-9-option-{choice_7}"
        
        dropdown_options = WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located((By.CSS_SELECTOR, option_css_selector))
        )
        if dropdown_options:
            dropdown_options[0].click()
            
        # เลือกจังหวัด
        dropdown_element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "selectProvince"))
        )

        dropdown_element.click()
        option_css_selector = f"#react-select-10-option-{choice_2}"
        
        dropdown_options = WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located((By.CSS_SELECTOR, option_css_selector))
        )
        if dropdown_options:
            dropdown_options[0].click()
            
        # โครงการ
        if str(row['โครงการ']) != 'ไม่มี':
            dropdown_element = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "selectProject"))
            )

            dropdown_element.click()
            option_css_selector = f"#react-select-11-option-{choice_8}"
            
            dropdown_options = WebDriverWait(driver, 10).until(
                EC.presence_of_all_elements_located((By.CSS_SELECTOR, option_css_selector))
            )
            if dropdown_options:
                dropdown_options[0].click() 
            
        # การถือครอง
        dropdown_element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "holdingLandsName"))
        )

        dropdown_element.click()
        option_css_selector = f"#react-select-13-option-{choice_9}"
        
        dropdown_options = WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located((By.CSS_SELECTOR, option_css_selector))
        )
        if dropdown_options:
            dropdown_options[0].click()
            
        # รายได้รวมภาตเกษตร
        input = driver.find_element(By.ID, "debtAmount")
        input.clear()
        input.send_keys(str(row['รายได้รวมภาตเกษตร']))
        
        # ปัญหา
        dropdown_element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "problemName"))
        )

        dropdown_element.click()
        option_css_selector = f"#react-select-14-option-{choice_10}"
        
        dropdown_options = WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located((By.CSS_SELECTOR, option_css_selector))
        )
        if dropdown_options:
            dropdown_options[0].click()
            
        # Accept terms and conditions
        policy_checkbox = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "policy-agree"))
        )
        driver.execute_script("arguments[0].scrollIntoView();", policy_checkbox)
        driver.execute_script("arguments[0].click();", policy_checkbox)

                
        time.sleep(15)
        driver.quit()

def browse_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    file_path_entry.delete(0, tk.END)
    file_path_entry.insert(0, file_path)
    
    
def start_processing():
    file_path = file_path_entry.get()
    if file_path:
        process_data(file_path)
    else:
        tk.messagebox.showerror("Error", "Please select an Excel file.")


root = tk.Tk()
root.title("โปรแกรมกรอกฟอร์ม")

# build the GUI 
file_path_label = tk.Label(root, text="Excel File:")
file_path_label.grid(row=0, column=0, padx=10, pady=10)

file_path_entry = tk.Entry(root, width=50)
file_path_entry.grid(row=0, column=1, padx=10, pady=10)

browse_button = tk.Button(root, text="Browse", command=browse_file)
browse_button.grid(row=0, column=2, padx=10, pady=10)

start_button = tk.Button(root, text="Start Processing", command=start_processing)
start_button.grid(row=1, column=1, pady=10)

# Run the GUI
root.mainloop()

