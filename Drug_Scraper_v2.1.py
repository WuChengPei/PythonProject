import subprocess
import sys
import tkinter as tk
from tkinter import messagebox, ttk, Canvas, Entry, Button, PhotoImage
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
import time
from pathlib import Path
from tkinter import filedialog
import os 

try:
    from selenium import webdriver
    from selenium.webdriver.chrome.service import Service
    from selenium.webdriver.common.by import By
    from selenium.webdriver.common.keys import Keys
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.common.exceptions import TimeoutException, NoSuchElementException
    from datetime import datetime   
    import random
    import pandas as pd
except ModuleNotFoundError:
    subprocess.check_call([
        sys.executable,
        "-m", "pip", "install", "selenium",
        "--trusted-host", "pypi.org",
        "--trusted-host", "pypi.python.org",
        "--trusted-host", "files.pythonhosted.org"
    ])
    from selenium import webdriver
    from selenium.webdriver.chrome.service import Service
    from selenium.webdriver.common.by import By
    from selenium.webdriver.common.keys import Keys
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.chrome.options import Options
    from selenium.webdriver.support import expected_conditions as EC
    from webdriver_manager.chrome import ChromeDriverManager
    from selenium.common.exceptions import TimeoutException, NoSuchElementException, WebDriverException
    from datetime import datetime
    import random
    import re
    import pandas as pd
    from pandas import ExcelWriter

def scrape_page(driver, date_elements, product_links, start_date, end_date, data):
    """處理單個頁面的藥品資料"""
    stop_scraping = False
    
    if len(date_elements) != len(product_links):
        print(f"警告：日期元素數量({len(date_elements)})與產品連結數量({len(product_links)})不匹配")
        return True

    for i in range(len(date_elements)):
        try:
            # 重新獲取元素
            date_elements = driver.find_elements(By.XPATH, "//span[contains(@id, 'lblPublishUpdateTime')]")
            product_links = driver.find_elements(By.XPATH, "//a[contains(@id, 'lbtnProductNameC')]")
            
            if i >= len(date_elements) or i >= len(product_links):
                print("警告：索引超出範圍，跳過當前項目")
                continue

            date_text = date_elements[i].text.strip()
            date_obj = datetime.strptime(date_text, "%Y/%m/%d")
            
            if not (start_date <= date_obj <= end_date):
                print(f"日期不符合條件: {date_text}")
                return True  # 停止爬取
                
            product_link = product_links[i]
            product_name = product_link.text.strip()
            print(f"蒐集資料中: {product_name} (日期: {date_text})")
            
            try:
                # 點擊並等待加載
                driver.execute_script("arguments[0].click();", product_link)
                WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.ID, "ContentPlaceHolder1_lblProductNameC"))
                )
            except Exception as e:
                print(f"點擊產品連結時發生錯誤: {e}")
                continue
            
            try:
                # 獲取詳細資料
                product_name_full = driver.find_element(By.ID, "ContentPlaceHolder1_lblProductNameC").text.strip()
                license_number = driver.find_element(By.ID, "ContentPlaceHolder1_lblLicense").text.strip()
                shortage_element = driver.find_element(By.XPATH, "//span[contains(@id, 'lblPublishContent_0')]")
                
                # 根據不同頁面類型處理資料
                if "恢復供應期間" in data[0].keys() if data else False:
                    data.append({
                        "項次": len(data) + 1,
                        "藥品名稱(許可證字號)": f"{product_name_full} ({license_number})",
                        "恢復供應期間": shortage_element.text.strip()
                    })
                else:
                    shortage_text = shortage_element.text.split("\n")[0] if shortage_element.text else ""
                    replacement_text = "\n".join(shortage_element.text.split("\n")[1:]) if shortage_element.text else ""
                    data.append({
                        "項次": len(data) + 1,
                        "藥品名稱(許可證字號)": f"{product_name_full} ({license_number})",
                        "短缺期間": shortage_text,
                        "替代藥品": replacement_text
                    })
                
                # 返回主頁
                back_button = driver.find_element(By.ID, "ContentPlaceHolder1_btnBack")
                driver.execute_script("arguments[0].click();", back_button)
                WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, "//span[contains(@id, 'lblPublishUpdateTime')]"))
                )
                
            except Exception as e:
                print(f"處理產品詳細資料時發生錯誤: {e}")
                try:
                    driver.navigate().back()
                except:
                    pass
                continue
            
        except Exception as e:
            print(f"處理藥品資料時發生錯誤: {e}")
            return True  # 停止爬取
            
    return False

def navigate_next_page(driver, page_num):
    """導航到下一頁"""
    try:
        next_button = driver.find_element(By.XPATH, f"//a[contains(@href, 'Page${page_num}')]")
        if not next_button:
            return False
            
        print("嘗試點擊下一頁按鈕...")
        next_button.click()
        time.sleep(2)
        
        current_first_date = driver.find_element(
            By.XPATH, "//span[contains(@id, 'lblPublishUpdateTime')]"
        ).text
        
        WebDriverWait(driver, 20).until(
            EC.text_to_be_present_in_element(
                (By.XPATH, "//span[contains(@id, 'lblPublishUpdateTime')]"),
                current_first_date
            )
        )
        print("成功進入下一頁。")
        return True
        
    except (NoSuchElementException, TimeoutException) as e:
        print(f"導航到下一頁時發生錯誤: {e}")
        return False

def scrape_data(driver, start_date, end_date, data):
    """通用的爬取函數"""
    start_date = datetime.strptime(start_date, "%Y/%m/%d")
    end_date = datetime.strptime(end_date, "%Y/%m/%d")
    page_num = 1
    
    try:
        WebDriverWait(driver, 60).until(
            EC.presence_of_element_located((By.XPATH, "//span[contains(@id, 'lblPublishUpdateTime')]"))
        )
        
        while True:
            date_elements = driver.find_elements(By.XPATH, "//span[contains(@id, 'lblPublishUpdateTime')]")
            product_links = driver.find_elements(By.XPATH, "//a[contains(@id, 'lbtnProductNameC')]")
            
            if not date_elements or not product_links:
                print("指定日期範圍內無符合的資料。")
                break
                
            if scrape_page(driver, date_elements, product_links, start_date, end_date, data):
                break
                
            page_num += 1
            if not navigate_next_page(driver, page_num):
                break
                
            time.sleep(random.uniform(2, 4))
            
    except Exception as e:
        print(f"執行爬取時發生錯誤: {e}")

def scrape_drug_data(driver, start_date, end_date, data):
    scrape_data(driver, start_date, end_date, data)

def scrape_restore_drug_data(driver, start_date, end_date, data):
    start_date = datetime.strptime(start_date, "%Y/%m/%d")
    end_date = datetime.strptime(end_date, "%Y/%m/%d")

    try:
        WebDriverWait(driver, 60).until(
            EC.presence_of_element_located((By.XPATH, "//span[contains(@id, 'lblPublishUpdateTime')]"))
        )
        while True:
            date_elements = driver.find_elements(By.XPATH, "//span[contains(@id, 'lblPublishUpdateTime')]")
            product_links = driver.find_elements(By.XPATH, "//a[contains(@id, 'lbtnProductNameC')]")

            if not date_elements or not product_links:
                print("指定日期範圍內無符合的資料。")
                break

            stop_scraping = False

            for i in range(len(date_elements)):
                try:
                    date_elements = driver.find_elements(By.XPATH, "//span[contains(@id, 'lblPublishUpdateTime')]")
                    product_links = driver.find_elements(By.XPATH, "//a[contains(@id, 'lbtnProductNameC')]")

                    date_text = date_elements[i].text.strip()
                    product_link = product_links[i]

                    date_obj = datetime.strptime(date_text, "%Y/%m/%d")

                    if start_date <= date_obj <= end_date:
                        product_name = product_link.text.strip()
                        print(f"蒐集資料中: {product_name} (日期: {date_text})")
                        product_link.click()
                        WebDriverWait(driver, 10).until(
                            EC.presence_of_element_located((By.ID, "ContentPlaceHolder1_lblProductNameC"))
                            )
                        product_name = driver.find_element(By.ID, "ContentPlaceHolder1_lblProductNameC").text.strip()
                        license_number = driver.find_element(By.ID, "ContentPlaceHolder1_lblLicense").text.strip()
                        restore_date_element = driver.find_element(By.XPATH,
                                                                   "//span[contains(@id, 'lblPublishContent_0')]")
                        restore_date = restore_date_element.text.strip()  # 提取文本內容

                        data.append({
                            "項次": len(data) + 1,
                            "藥品名稱(許可證字號)": f"{product_name} ({license_number})",
                            "恢復供應期間": restore_date
                        })
                        # 使用「返回上頁」按鈕返回主頁
                        back_button = driver.find_element(By.ID, "ContentPlaceHolder1_btnBack")
                        back_button.click()
                        WebDriverWait(driver, 10).until(
                            EC.presence_of_element_located((By.XPATH, "//span[contains(@id, 'lblPublishUpdateTime')]"))
                        )
                    else:
                        stop_scraping = True
                        print(f"日期不符合條件: {date_text}")
                        break
                except Exception as e:
                    print(f"處理藥品資料時發生錯誤: {e}")
                    stop_scraping = True
                    break
            if stop_scraping:
                break

            try:
                next_button2 = driver.find_element(By.XPATH, "//a[contains(@href, 'Page$2')]")
                if next_button2:
                    print("嘗試點擊下一頁按鈕...")
                    next_button2.click()

                    # 獲取當前頁面的第一個日期
                    current_first_date2 = driver.find_element(By.XPATH,
                                                              "//span[contains(@id, 'lblPublishUpdateTime')]").text

                    # 點擊按鈕並等待內容更新

                    WebDriverWait(driver, 20).until(
                        EC.text_to_be_present_in_element(
                            (By.XPATH, "//span[contains(@id, 'lblPublishUpdateTime')]"),
                            current_first_date2
                        )
                    )
                    print("成功進入下一頁。")
                else:
                    print("未找到下一頁按鈕，結束爬取。")
                    break
            except NoSuchElementException:
                print("未找到下一頁按鈕，結束爬取。")
                break
            except TimeoutException:
                print("下一頁內容加載超時，結束爬取。")
                break

            time.sleep(random.uniform(2, 4))

            date_elements2 = driver.find_elements(By.XPATH, "//span[contains(@id, 'lblPublishUpdateTime')]")
            product_links2 = driver.find_elements(By.XPATH, "//a[contains(@id, 'lbtnProductNameC')]")

            if not date_elements2 or not product_links2:
                print("指定日期範圍內無符合的資料，請確認日期設定是否正確。")
                break
            stop_scraping = False

            for i in range(len(date_elements2)):
                try:
                    date_elements2 = driver.find_elements(By.XPATH, "//span[contains(@id, 'lblPublishUpdateTime')]")
                    product_links2 = driver.find_elements(By.XPATH, "//a[contains(@id, 'lbtnProductNameC')]")

                    date_text2 = date_elements2[i].text.strip()
                    product_link2 = product_links2[i]

                    date_obj2 = datetime.strptime(date_text2, "%Y/%m/%d")

                    if start_date <= date_obj2 <= end_date:
                        product_name2 = product_link2.text.strip()
                        print(f"蒐集資料中: {product_name2} (日期: {date_text2})")
                        product_link2.click()
                        WebDriverWait(driver, 10).until(
                            EC.presence_of_element_located((By.ID, "ContentPlaceHolder1_lblProductNameC"))
                        )
                        product_name2 = driver.find_element(By.ID, "ContentPlaceHolder1_lblProductNameC").text.strip()
                        license_number2 = driver.find_element(By.ID, "ContentPlaceHolder1_lblLicense").text.strip()
                        restore_date_element = driver.find_element(By.XPATH,
                                                                   "//span[contains(@id, 'lblPublishContent_0')]")
                        restore_date2 = restore_date_element.text.strip()

                        data.append({
                            "項次": len(data) + 1,
                            "藥品名稱(許可證字號)": f"{product_name2} ({license_number2})",
                            "恢復供應期間": restore_date2
                        })
                        # 使用「返回上頁」按鈕返回主頁
                        back_button2 = driver.find_element(By.ID, "ContentPlaceHolder1_btnBack")
                        back_button2.click()
                        WebDriverWait(driver, 10).until(
                            EC.presence_of_element_located((By.XPATH, "//span[contains(@id, 'lblPublishUpdateTime')]"))
                        )
                    else:
                        stop_scraping = True
                        print(f"日期不符合條件: {date_text2}")
                        break
                except Exception as e:
                    print(f"處理藥品資料時發生錯誤: {e}")
                    stop_scraping = True
                    break
                if stop_scraping:
                    break

            try:
                next_button3 = driver.find_element(By.XPATH, "//a[contains(@href, 'Page$3')]")
                if next_button3:
                    print("嘗試點擊下一頁按鈕...")
                    next_button3.click()

                    # 獲取當前頁面的第一個日期
                    current_first_date3 = driver.find_element(By.XPATH,
                                                                  "//span[contains(@id, 'lblPublishUpdateTime')]").text

                        # 點擊按鈕並等待內容更新

                    WebDriverWait(driver, 60).until(
                        EC.text_to_be_present_in_element(
                                (By.XPATH, "//span[contains(@id, 'lblPublishUpdateTime')]"),
                                current_first_date3
                            )
                        )
                    print("成功進入下一頁。")
                else:
                    print("未找到下一頁按鈕，結束爬取。")
                    break
            except NoSuchElementException:
                print("未找到下一頁按鈕，結束爬取。")
                break
            except TimeoutException:
                print("下一頁內容加載超時，結束爬取。")
                break

            time.sleep(random.uniform(2, 4))

            date_elements3 = driver.find_elements(By.XPATH, "//span[contains(@id, 'lblPublishUpdateTime')]")
            product_links3 = driver.find_elements(By.XPATH, "//a[contains(@id, 'lbtnProductNameC')]")

            if not date_elements3 or not product_links3:
                print("指定日期範圍內無符合的資料，請確認日期設定是否正確。")
                break
            stop_scraping = False

            for i in range(len(date_elements3)):
                try:
                    date_elements3 = driver.find_elements(By.XPATH, "//span[contains(@id, 'lblPublishUpdateTime')]")
                    product_links3 = driver.find_elements(By.XPATH, "//a[contains(@id, 'lbtnProductNameC')]")

                    date_text3 = date_elements3[i].text.strip()
                    product_link3 = product_links3[i]

                    date_obj3 = datetime.strptime(date_text3, "%Y/%m/%d")

                    if start_date <= date_obj3 <= end_date:
                        product_name3 = product_link3.text.strip()
                        print(f"蒐集資料中: {product_name3} (日期: {date_text3})")
                        product_link3.click()
                        WebDriverWait(driver, 10).until(
                            EC.presence_of_element_located((By.ID, "ContentPlaceHolder1_lblProductNameC"))
                        )
                        product_name3 = driver.find_element(By.ID, "ContentPlaceHolder1_lblProductNameC").text.strip()
                        license_number3 = driver.find_element(By.ID, "ContentPlaceHolder1_lblLicense").text.strip()
                        restore_date_element = driver.find_element(By.XPATH,
                                                                       "//span[contains(@id, 'lblPublishContent_0')]")
                        restore_date3 = restore_date_element.text.strip()

                        data.append({
                                "項次": len(data) + 1,
                                "藥品名稱(許可證字號)": f"{product_name3} ({license_number3})",
                                "恢復供應期間": restore_date3
                        })
                            # 使用「返回上頁」按鈕返回主頁
                        back_button3 = driver.find_element(By.ID, "ContentPlaceHolder1_btnBack")
                        back_button3.click()
                        WebDriverWait(driver, 10).until(
                                EC.presence_of_element_located((By.XPATH, "//span[contains(@id, 'lblPublishUpdateTime')]"))
                        )
                    else:
                        stop_scraping = True
                        print(f"日期不符合條件: {date_text3}")
                        break
                except Exception as e:
                    print(f"處理藥品資料時發生錯誤: {e}")
                    stop_scraping = True
                    break
                if stop_scraping:
                    break

            try:
                next_button4 = driver.find_element(By.XPATH, "//a[contains(@href, 'Page$4')]")
                if next_button4:
                    print("嘗試點擊下一頁按鈕...")
                    next_button4.click()

                        # 獲取當前頁面的第一個日期
                    current_first_date4 = driver.find_element(By.XPATH,
                                                                  "//span[contains(@id, 'lblPublishUpdateTime')]").text

                        # 點擊按鈕並等待內容更新

                    WebDriverWait(driver, 60).until(
                        EC.text_to_be_present_in_element(
                (By.XPATH, "//span[contains(@id, 'lblPublishUpdateTime')]"),
                                current_first_date4
                        )
                    )
                    print("成功進入下一頁。")
                else:
                    print("未找到下一頁按鈕，結束爬取。")
                    break
            except NoSuchElementException:
                print("未找到下一頁按鈕，結束爬取。")
                break
            except TimeoutException:
                print("下一頁內容加載超時，結束爬取。")
                break

            time.sleep(random.uniform(2, 4))

            date_elements4 = driver.find_elements(By.XPATH, "//span[contains(@id, 'lblPublishUpdateTime')]")
            product_links4 = driver.find_elements(By.XPATH, "//a[contains(@id, 'lbtnProductNameC')]")

            if not date_elements4 or not product_links4:
                print("指定日期範圍內無符合的資料，請確認日期設定是否正確。")
                break
            stop_scraping = False

            for i in range(len(date_elements4)):
                try:
                    date_elements4 = driver.find_elements(By.XPATH, "//span[contains(@id, 'lblPublishUpdateTime')]")
                    product_links4 = driver.find_elements(By.XPATH, "//a[contains(@id, 'lbtnProductNameC')]")

                    date_text4 = date_elements4[i].text.strip()
                    product_link4 = product_links4[i]

                    date_obj4 = datetime.strptime(date_text4, "%Y/%m/%d")

                    if start_date <= date_obj4 <= end_date:
                        product_name4 = product_link4.text.strip()
                        print(f"蒐集資料中: {product_name4} (日期: {date_text4})")
                        product_link4.click()
                        WebDriverWait(driver, 10).until(
                                EC.presence_of_element_located((By.ID, "ContentPlaceHolder1_lblProductNameC"))
                        )
                        product_name4 = driver.find_element(By.ID, "ContentPlaceHolder1_lblProductNameC").text.strip()
                        license_number4 = driver.find_element(By.ID, "ContentPlaceHolder1_lblLicense").text.strip()
                        restore_date_element = driver.find_element(By.XPATH,
                                                                       "//span[contains(@id, 'lblPublishContent_0')]")
                        restore_date4 = restore_date_element.text.strip()

                        data.append({
                                "項次": len(data) + 1,
                                "藥品名稱(許可證字號)": f"{product_name4} ({license_number4})",
                                "恢復供應期間": restore_date4
                        })
                            # 使用「返回上頁」按鈕返回主頁
                        back_button4 = driver.find_element(By.ID, "ContentPlaceHolder1_btnBack")
                        back_button4.click()
                        WebDriverWait(driver, 10).until(
                            EC.presence_of_element_located((By.XPATH, "//span[contains(@id, 'lblPublishUpdateTime')]"))
                        )
                    else:
                        stop_scraping = True
                        print(f"日期不符合條件: {date_text4}")
                        break
                except Exception as e:
                    print(f"處理藥品資料時發生錯誤: {e}")
                    stop_scraping = True
                    break
                if stop_scraping:
                    break
    except Exception as e:
        print(f"執行爬取時發生錯誤: {e}")

def relative_to_assets(path: str) -> Path:
    if getattr(sys, 'frozen', False):  # PyInstaller 打包後
        base_path = Path(sys.executable).parent  # `.exe` 的所在目錄
    else:  # 開發環境
        base_path = Path(__file__).parent
    assets_path = base_path / "assets"  # `assets` 應該在 `.exe` 同個資料夾
    full_path = assets_path / path

    if not full_path.exists():
        print(f"⚠️ 錯誤: 找不到 {full_path}，請確認 `assets` 資料夾是否正確！")
    return full_path

def get_chromedriver_path() -> str:
    if getattr(sys, 'frozen', False):  # PyInstaller 打包後
        base_path = Path(sys.executable).parent  # `.exe` 的所在目錄
    else:  # 開發環境
        base_path = Path(__file__).parent

    chromedriver_path = base_path / "chromedriver.exe"

    if not chromedriver_path.exists():
        print(f"⚠️ 錯誤: 找不到 {chromedriver_path}，請確認 `chromedriver.exe` 是否存在！")

    return str(chromedriver_path)

class CrawlerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Supply Scraper v2.1")
        self.root.geometry("860x540")
        self.root.configure(bg="#FFFFFF")

        icon_path = relative_to_assets("P.ico")  # 使用相對路徑函式
        if icon_path.exists():
            self.root.iconbitmap(str(icon_path))  # 設定 icon
        else:
            print("⚠️ 錯誤: 找不到 P.ico，請確認 assets 目錄是否正確！")

        canvas = Canvas(
            root,
            bg="#FFFFFF",
            height=540,
            width=860,
            bd=0,
            highlightthickness=0,
            relief="ridge"
        )
        canvas.place(x=0, y=0)
        canvas.create_rectangle(2.0, 2.0, 858.0, 538.0, fill="#FAE2DC", outline="")
        canvas.create_rectangle(2.0, 3.0, 428.0, 538.0, fill="#D7927C", outline="")

        self.image_image_1 = PhotoImage(file=relative_to_assets("image_1.png"))
        canvas.create_image(226.0, 143.0, image=self.image_image_1)

        self.image_image_2 = PhotoImage(file=relative_to_assets("image_2.png"))
        canvas.create_image(806.0, 512.0, image=self.image_image_2)

        canvas.create_text(
            473.0, 64.0, anchor="nw", text="Supply \nScraper v2.1 ",
            fill="#220000", font=("Arial", 36 * -1)
        )

        self.image_image_3 = PhotoImage(file=relative_to_assets("image_3.png"))
        canvas.create_image(191.0, 391.0, image=self.image_image_3)

        canvas.create_text(
            473.0, 193.0, anchor="nw", text="Start Date (format: YYYY/MM/DD)",
            fill="#000000", font=("Arial", 20 * -1)
        )

        canvas.create_text(
            473.0, 279.0, anchor="nw", text="End Date  (format: YYYY/MM/DD)",
            fill="#000000", font=("Arial", 20 * -1)
        )

        self.entry_1 = Entry(root, bd=0, bg="#FFFFFF", fg="#000716", highlightthickness=0)
        self.entry_1.place(x=473.0, y=226.0, width=317.0, height=29.0)

        self.entry_2 = Entry(root, bd=0, bg="#FFFFFF", fg="#000716", highlightthickness=0)
        self.entry_2.place(x=473.0, y=312.0, width=317.0, height=29.0)

        self.button_image_1 = PhotoImage(file=relative_to_assets("button_1.png"))
        self.start_button = Button(
            root, image=self.button_image_1, borderwidth=0, highlightthickness=0,
            background="#FAE2DC", activebackground="#FAE2DC", command=self.start_crawling,
            relief="flat"
        )
        self.start_button.place(x=588.0, y=387.0, width=86.0, height=45.0)
        
        # 添加儲存位置選擇按鈕
        self.button_image_2 = PhotoImage(file=relative_to_assets("button_2.png"))  # 需要添加新的按鈕圖片
        self.select_path_button = Button(
            root, image=self.button_image_2, borderwidth=0, highlightthickness=0,
            background="#FAE2DC", activebackground="#FAE2DC", command=self.select_save_path,
            relief="flat"
        )
        self.select_path_button.place(x=473.0, y=350.0, width=317.0, height=29.0)
        
        # 儲存路徑的變量
        self.save_path = os.path.expanduser("~/Documents")  # 預設為文件夾

        # 調整按鈕位置
        button_y = 387.0  # 兩個按鈕的垂直位置相同
        button_spacing = 30  # 按鈕之間的間距
        
        # 計算兩個按鈕的水平位置，使其置中對齊
        total_width = 86.0 + button_spacing + 86.0  # 兩個按鈕加間距的總寬度
        start_x = 473.0 + (317.0 - total_width) / 2  # 從輸入框的左邊開始計算置中位置
        
        # 選擇儲存位置按鈕
        self.button_image_2 = PhotoImage(file=relative_to_assets("button_2.png"))
        self.select_path_button = Button(
            root, image=self.button_image_2, borderwidth=0, highlightthickness=0,
            background="#FAE2DC", activebackground="#FAE2DC", command=self.select_save_path,
            relief="flat"
        )
        self.select_path_button.place(x=start_x, y=button_y, width=86.0, height=45.0)
        
        # 開始爬取按鈕
        self.button_image_1 = PhotoImage(file=relative_to_assets("button_1.png"))
        self.start_button = Button(
            root, image=self.button_image_1, borderwidth=0, highlightthickness=0,
            background="#FAE2DC", activebackground="#FAE2DC", command=self.start_crawling,
            relief="flat"
        )
        self.start_button.place(x=start_x + 86.0 + button_spacing, y=button_y, width=86.0, height=45.0)

        # 確保與上方輸入框保持一致的間距
        input_bottom = 312.0 + 29.0  # 最後一個輸入框的底部位置
        button_spacing_top = button_y - input_bottom  # 按鈕與輸入框的間距
        
    def select_save_path(self):
        """選擇檔案儲存位置"""
        path = filedialog.askdirectory(initialdir=self.save_path)
        if path:
            self.save_path = path
            messagebox.showinfo("成功", f"已選擇儲存位置：{path}")

    def start_crawling(self):
        start_date = self.entry_1.get().strip()
        end_date = self.entry_2.get().strip()

        if not start_date or not end_date:
            messagebox.showwarning("錯誤", "請輸入完整的開始日期和結束日期！")

        try:
            start_date_t = datetime.strptime(start_date, "%Y%m%d").strftime("%Y/%m/%d")
            end_date_t = datetime.strptime(end_date, "%Y%m%d").strftime("%Y/%m/%d")
        except ValueError:
            messagebox.showwarning("錯誤", "日期格式錯誤，請輸入 YYYYMMDD 或 YYYY/MM/DD")
            return

        chrome_options = Options()
        #chrome_options.add_argument("--headless")
        chrome_options.add_argument("--disable-gpu")
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")

        try:
            driver = webdriver.Chrome(service=Service(), options=chrome_options)
            driver.maximize_window()
            url = "http://dsms.fda.gov.tw/DrugList.aspx?s=3"

            driver.get(url)
            time.sleep(5)
            data = []
            scrape_drug_data(driver, start_date_t, end_date_t, data)
            driver.quit()

            if data:
                df = pd.DataFrame(data)
                output_file = os.path.join(self.save_path, "建議使用替代品項.xlsx")  # 將路徑與檔名結合
                df.to_excel(output_file, index=False)
                df["短缺期間"] = df["短缺期間"].str.replace(r'^1\.?', '', regex=True)
                df["替代藥品"] = df["替代藥品"].str.replace(r'^2\.?', '', regex=True)
                df.to_excel(output_file, index=False)
                messagebox.showinfo("通知", "跑完一半，請按確定，再加油^^")
            else:
                messagebox.showinfo("完成", "未蒐集到任何資料。")

        except Exception as e:
            messagebox.showerror("錯誤", f"發生錯誤: {str(e)}")

        try:
            driver = webdriver.Chrome(service=Service(), options=chrome_options)
            driver.maximize_window()
            url2 = "http://dsms.fda.gov.tw/DrugList.aspx?s=2"

            driver.get(url2)
            time.sleep(5)
            data2 = []
            scrape_restore_drug_data(driver, start_date_t, end_date_t, data2)
            driver.quit()

            if data2:
                df2 = pd.DataFrame(data2)
                output_file2 = os.path.join(self.save_path, "已恢復供應品項.xlsx")
                df2.to_excel(output_file2, index=False)
                df2.to_excel(output_file2, index=False)
                messagebox.showinfo("完成", f"爬取完成，資料已儲存至 {output_file2}")
            else:
                messagebox.showinfo( "完成", "未蒐集到任何資料。")
        except Exception as e:
            messagebox.showerror( "錯誤", f"發生錯誤: {str(e)}")


if __name__ == "__main__":
    root = tk.Tk()
    app = CrawlerApp(root)
    root.resizable(False, False)
    root.mainloop()
