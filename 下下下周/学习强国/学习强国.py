import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

def extract_links_from_page(driver, link_xpath):
    elements = driver.find_elements(By.XPATH, link_xpath)
    links = [element.get_attribute('href') for element in elements]
    return links

def scrape_links(url, link_xpath, next_button_xpath, click_times):
    driver = webdriver.Chrome()
    all_links = []

    try:
        driver.get(url)

        for _ in range(click_times):
            time.sleep(1)  # 等待页面加载

            # 提取当前页面的链接
            links = extract_links_from_page(driver, link_xpath)
            all_links.extend(links)

            # 点击下一页按钮
            next_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, next_button_xpath))
            )
            next_button.click()

    finally:
        driver.quit()

    return all_links

# 示例用法
url = 'https://www.xuexi.cn/c06bf4acc7eef6ef0a560328938b5771/9a3668c13f6e303932b5e0e100fc248b.html'  # 替换为实际网站
link_xpath = '//*[@id="app"]/div/div[2]/div[3]/div[2]/ul/li[1]/a'  # 替换为实际包含链接的div的XPath
next_button_xpath = '//*[@id="app"]/div/div[2]/div[3]/div[2]/div/div/button[2]/i'  # 替换为实际的“下一页”按钮的XPath
click_times = 5  # 替换为需要点击“下一页”的次数

all_links = scrape_links(url, link_xpath, next_button_xpath, click_times)

# 输出所有链接
for link in all_links:
    print(link)
