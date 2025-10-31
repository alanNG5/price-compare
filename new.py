from playwright.sync_api import sync_playwright
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Side
import time
import random
import readfile

start = time.time()

with sync_playwright() as p:
    browser = p.chromium.launch(headless=True)

    # page = browser.new_page(user_agent="Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/103.0.5060.114 Safari/537.36")
    page = browser.new_page(user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")

    catalogue = []

    page.goto("https://www.hktvmall.com/hktv/zh/login", wait_until="networkidle", timeout=0)

    page.locator("#j_username_active").fill("98079763")
    page.locator("#j_pwd_mobile").fill("shanwei8385")
    # page.locator(".login-mobileLoginForm").locator("button[type='submit']").click()
    page.keyboard.press("Enter")

    page.wait_for_load_state("networkidle", timeout=0)

    for query in readfile.data:

        code = query[0]

        word = query[1]

        init_info = {
        "Code": code,
        "Keyword": word,
        "Product Name": "",
        "Price": 0,
        "Packing": "",
        "Stock": 0,
        "Vendor 1": "",
        "Vendor 1 - Product Name": "NIL",
        }

        page.goto("https://www.hktvmall.com/hktv/zh/main/Nippon-Premium-Foods/s/H8224001/p/H8224001_S_{}".format(code), wait_until="networkidle", timeout=0)

        product_info = page.evaluate("""
            function () {

                let priceTag = productData.priceList.length > 1 ? parseFloat(productData?.priceList[1].value.toFixed(2)) : parseFloat(productData?.priceList[0].value.toFixed(2))

                let bundleOffer = ""

                if (productData.thresholdPromotionList.length > 0) {
                bundleOffer = " ( " + productData.thresholdPromotionList[0].label + ": " + productData.thresholdPromotionList[0].description + " )"
                }

                const productObj = {
                    "Product Name": productData?.brandName + " - " + productData?.name + bundleOffer,
                    "Price": priceTag,
                    "Packing": productData?.packingSpec,
                    "Stock": productData?.stock?.stockLevel,
                }
                return productObj
            }
        """)

        init_info.update(product_info)


        # params of "sort=hktvProductSellingPriceAscV2" triggers the option of select
        page.goto("https://www.hktvmall.com/hktv/zh/search_a?keyword={}&sort=hktvProductSellingPriceAscV2".format(word), wait_until="networkidle", timeout=0)

        search_links = page.evaluate("""
            function (code) {

                const info = document.querySelectorAll(".product-brief")

                let sortedList = []

                for (const item of info) {

                    let checkOutOfStock = item.querySelector(".out-of-stock-label")
                    let checkStoreCode = item.getAttribute("data-id").slice(0,8)
                    let pageLink = item.querySelector(".product-brief .info-wrapper~a").href

                    if ( checkStoreCode !== "H8224001" && checkOutOfStock === null ) {
                    sortedList.push(pageLink)
                    }
                }

                sortedThreeItems = sortedList.slice(0, 3)
                return sortedThreeItems
            }
        """, code)

        looper = 1
        time.sleep(random.randint(10, 20))

        for link in search_links:

            page.goto(link, wait_until="networkidle", timeout=0)

            vendors_info = page.evaluate("""
                function (looper) {

                    let priceTag = productData.priceList.length > 1 ? parseFloat(productData?.priceList[1].value.toFixed(2)) : parseFloat(productData?.priceList[0].value.toFixed(2))

                    let productObj = {
                        [`Vendor ${looper}`]: productData?.store?.name,
                        [`Vendor ${looper} - Product Name`]: productData?.brandName + " - " + productData?.name,
                        [`Vendor ${looper} - Price`]: priceTag,
                        [`Vendor ${looper} - Packing`]: productData?.packingSpec,
                        [`Vendor ${looper} - Stock`]: productData?.stock?.stockLevel,
                    }

                    return productObj
                }
            """, looper)

            looper += 1

            init_info.update(vendors_info)

            time.sleep(int(time.time()%60)%6)

        catalogue.append(init_info)

        print(len(catalogue))
        print(f"Item(s) for keyword of \"{word}\" is/are captured.")

    if len(catalogue) %50 == 0:
        df = pd.DataFrame(catalogue)
        if not df.empty:
            df.to_excel('backup.xlsx', index=False)

    print(f"Total items: {len(catalogue)}")

    df = pd.DataFrame(catalogue)

    if not df.empty:
        df.to_excel('hktvmall.xlsx', index=False)

        workbook = load_workbook('hktvmall.xlsx')
        worksheet = workbook.active

        for column in worksheet.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = (max_length + 2)
            worksheet.column_dimensions[column_letter].width = adjusted_width

        # thick_border = Border(right = Side(style='thick'),)

        for row in worksheet.iter_rows(min_col=7, max_col=7): # col 7 as column "G"
            for cell in row:
                # cell.border = thick_border
                cell.alignment = Alignment(wrap_text=True, vertical="center", wrapText=True)
                # Enable text wrapping for the cell: For Excel to display the \n characters as actual line breaks, the cell's wrapText alignment property must be set to True.

        workbook.save('hktvmall.xlsx')

    else:
        print("DataFrame is empty, nothing to export.")


    browser.close()

print(f"Time taken: {round(time.time() - start, 3)} seconds")