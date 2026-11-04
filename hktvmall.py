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

    page = browser.new_page(user_agent="Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/103.0.5060.114 Safari/537.36")

    catalogue = []

    page.goto("https://www.hktvmall.com/hktv/zh/login", wait_until="networkidle", timeout=0)

    page.locator("#j_username_active").fill("00000000")
    page.locator("#j_pwd_mobile").fill("xxxxxxxx")
    # must substitute "00000000" & "xxxxxxxx" with registered phone number & password respectively
    page.locator(".login-mobileLoginForm").locator("button[type='submit']").click()
    # page.keyboard.press("Enter")

    page.wait_for_load_state("networkidle", timeout=0)

    for query in readfile.data:

        code = query[0]

        word = query[1]

        # temp_var = ""
        # reg = "^.*{}.*$".format(temp_var)

        pattern = f"^(?=.*{query[2]})(?=.*{query[3]})$"

        print(pattern)

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

                if (typeof productData !== "undefined") {

                    if (productData.hasOwnProperty("thresholdPromotionList") && productData.thresholdPromotionList.length > 0) {
                        bundleOffer = " ( 大手折扣：" + productData.thresholdPromotionList[0].label + ": " + productData.thresholdPromotionList[0].description + " )"
                    }
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

        # pdt_info = product_info["Product Name"].replace(" ( 大手折扣", "\n\n ( 大手折扣")

        init_info.update(product_info)

        page.goto("https://www.hktvmall.com/hktv/zh/search_a?keyword={}".format(word), wait_until="networkidle", timeout=0)

        vendors_product_details = page.evaluate("""
            function (pattern) {

                const info = document.querySelectorAll(".product-brief-wrapper .product-brief")

                let matchedList = []

                for (const item of info) {

                    let product = {}

                    let checkOutOfStock = item.querySelector(".out-of-stock-label")
                    let checkStoreCode = item.getAttribute("data-id").slice(0,8)
                    let checkPackingSpec = item.querySelector(".info-wrapper .packing-spec")
                    let packingSpec = checkPackingSpec != null? checkPackingSpec.innerText : ""

                    const regexp_pattern = new RegExp(pattern, "i")
                    let productName = item.querySelector(".info-wrapper .brand-product-name").innerText

                    if ( checkStoreCode !== "H8224001" && checkOutOfStock === null && regexp_pattern.test(productName)) {
                        product["name"] = productName
                        product["url"] = item.querySelector(".info-wrapper~a").href
                        product["price"] = parseFloat(item.querySelector(".info-wrapper .lower-wrapper .price-label .price").innerText.replace(/[$,]/g,""))
                        product["packing"] = packingSpec
                    }

                    if (Object.keys(product).length > 0) {
                        matchedList.push(product)
                    }
                }

                let sortedList = matchedList.sort((a, b) => a.price - b.price)

                let sortedThreeItems = sortedList.slice(0, 3)

                return sortedThreeItems
            }
        """, pattern)

        looper = 1

        time.sleep(random.randint(10, 20))

        print(vendors_product_details)

        for product_details in vendors_product_details:

            init_info[f"Vendor {looper}"] = ""
            init_info[f"Vendor {looper} - Product Name"] = product_details["name"]
            init_info[f"Vendor {looper} - Price"] = product_details["price"]
            init_info[f"Vendor {looper} - Packing"] = product_details["packing"]

            link = product_details["url"]

            page.goto(link, wait_until="networkidle", timeout=0)

            other_vendor_info = page.evaluate("""
                function (looper) {

                    let productObj = {
                        [`Vendor ${looper}`]: productData?.store?.name,
                        [`Vendor ${looper} - Stock`]: productData?.stock?.stockLevel,
                    }

                    return productObj
                }
            """, looper)

            looper += 1

            init_info.update(other_vendor_info)

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

        thick_border = Border(right = Side(style='thick'),)

        for row in worksheet.iter_rows(min_col=7, max_col=7): # col 7 as column "G"
            for cell in row:
                cell.border = thick_border
                cell.alignment = Alignment(wrap_text=True, vertical="center", wrapText=True)
                # Enable text wrapping for the cell: For Excel to display the \n characters as actual line breaks, the cell's wrapText alignment property must be set to True.

        workbook.save('hktvmall.xlsx')

    else:
        print("DataFrame is empty, nothing to export.")

    browser.close()

print(f"Time taken: {round(time.time() - start, 3)} seconds")



