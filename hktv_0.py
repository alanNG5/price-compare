from playwright.sync_api import sync_playwright
import pandas as pd
from openpyxl import load_workbook
import time
import readfile

queries = ["ORIHIRO - 白桃+巨峰提子味蒟蒻啫喱", "HOTEi 低脂水煮吞拿魚罐頭 3罐裝", "味之素 Knorr湯意粉 蕃茄湯"]

start = time.time()

with sync_playwright() as p:
    browser = p.chromium.launch(headless=True)

    page = browser.new_page(user_agent="Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/103.0.5060.114 Safari/537.36")

    catalogue = []

    for query in readfile.data:

        word = query[1]
        items_same_category = []

        # loop through page > 1 if needed; starting from 0
        # for page_number in range(0, 11):
        #         page.goto("https://www.hktvmall.com/hktv/zh/search_a?keyword={}&page={}".format(word, page_number))

        page.goto("https://www.hktvmall.com/hktv/zh/search_a?keyword={}".format(word), wait_until="networkidle")

        search_links = page.evaluate("""
            (function () {
                const urlList = [];
                const urlSelector = document.querySelectorAll(".product-brief > a");
                urlSelector.forEach((url) => {
                    urlList.push(url.href);
                });
                return urlList;
            })();
        """)

        for link in search_links:
            time.sleep(1)

            page.goto(link, wait_until="networkidle")

            product_info = page.evaluate("""
                (function () {
                    const product = {
                        name: productData?.name,
                        vendor: productData?.storeName,
                        price: parseInt(productData?.priceList[1]?.value?.toFixed(2)),
                        packing: productData?.packingSpec,
                        stock: productData?.stock?.stockLevel,
                    };
                    return product;
                })();
            """)

            product_info["code"] = query[0]
            product_info["keyword"] = query[1]
            items_same_category.append(product_info)

        def sort_by_price(e):
            return e["price"]

        items_same_category.sort(reverse=True, key=sort_by_price)

        catalogue.extend(items_same_category)

    # print(catelogue)
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

        workbook.save('hktvmall.xlsx')

    else:
        print("DataFrame is empty, nothing to export.")


    browser.close()


print(f"Time taken: {round(time.time() - start, 3)} seconds")

