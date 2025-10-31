from playwright.sync_api import sync_playwright
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Side
import time
import readfile

start = time.time()


with sync_playwright() as p:
    browser = p.chromium.launch(headless=True)

    page = browser.new_page(user_agent="Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/103.0.5060.114 Safari/537.36")

    catalogue = []

    for query in readfile.data:

        try:
            code = query[0]

            word = query[1]

            init_info = {
            "Code": code,
            "Keyword": word,
            "Product Name": "",
            "Price": 0,
            "Packing": "",
            "Stock": 0,
            }

            page.goto("https://www.hktvmall.com/hktv/zh/main/Nippon-Premium-Foods/s/H8224001/p/H8224001_S_{}".format(code), wait_until="networkidle", timeout=0)

            print(code, word)

            product_info = page.evaluate("""
                function () {

                    let priceTag = productData.priceList.length > 1 ? parseFloat(productData?.priceList[1].value.toFixed(2)) : parseFloat(productData?.priceList[0].value.toFixed(2))

                    let bundleOffer = ""

                    if (productData !== "undefined") {
                        if (productData.hasOwnProperty("thresholdPromotionList") && productData.thresholdPromotionList.length > 0) {
                            bundleOffer = " ( " + productData.thresholdPromotionList[0].label + ": " + productData.thresholdPromotionList[0].description + " )"
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

            init_info.update(product_info)

            catalogue.append(init_info)

        except Exception as e:
            print(f"Error processing code {code}: {e}")

    print(f"Total items: {len(catalogue)}")

    df = pd.DataFrame(catalogue)

    if not df.empty:
        df.to_excel('npf.xlsx', index=False)

        workbook = load_workbook('npf.xlsx')
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

        workbook.save('npf.xlsx')

    else:
        print("DataFrame is empty, nothing to export.")

    browser.close()

print(f"Time taken: {round(time.time() - start, 3)} seconds")