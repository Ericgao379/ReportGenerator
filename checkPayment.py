from selenium.webdriver.common.by import By

def detect_payment_logos_from_footer_selenium(driver):
    try:
        footer = driver.find_element(By.TAG_NAME, 'footer')
        img_elements = footer.find_elements(By.TAG_NAME, 'img')

        found = []
        known_keywords = ['visa', 'mastercard', 'paypal', 'applepay', 'amex', 'klarna', 'maestro']

        for img in img_elements:
            src = img.get_attribute('src') or ""
            alt = img.get_attribute('alt') or ""
            title = img.get_attribute('title') or ""
            combined = (src + alt + title).lower()

            for keyword in known_keywords:
                if keyword in combined:
                    found.append(keyword.capitalize())

        return list(set(found))  # 去重返回
    except Exception as e:
        print(f"⚠️ Failed to detect logos in footer: {e}")
        return []

