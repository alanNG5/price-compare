import asyncio
from playwright.async_api import async_playwright, Playwright
import pandas as pd
from openpyxl import load_workbook
import time

input = "Nippon Premium Foods"

start = time.time()

async def run(p: Playwright):

    browser = await p.chromium.launch(headless=True)

    page = await browser.new_page(user_agent="Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/103.0.5060.114 Safari/537.36")

    # txt = input.replace(" ", "")

    all_items_list = []

    # loop through page > 1 if needed; starting from 0
    for page_number in range(0, 8):
        await page.goto("https://www.hktvmall.com/hktv/zh/search_a?keyword={}&page={}".format(input, page_number), wait_until="networkidle")

    # await page.goto("https://www.hktvmall.com/hktv/zh/search_a?keyword={}".format(input),wait_until='networkidle')

        items_list = await page.evaluate("""
            (function () {
                const productList = [];
                const products = document.querySelectorAll(".product-brief");
                for(const details of products) {
                    const product = {
                        code: details.getAttribute("data-id").substring(11),
                        product_name: details.querySelector(".info-wrapper .upper-wrapper .brand-product-name")?.innerText,
                        price: parseFloat(details.querySelector(".info-wrapper .lower-wrapper .price-label .price")?.innerText.replace("$ ", "")),
                        packing: details.querySelector(".info-wrapper .upper-wrapper .packing-spec")?.innerText,
                        store: details.querySelector(".info-wrapper .lower-wrapper a.store-name-label div span")?.innerText,
                    };
                    productList.push(product);
                };
                let sortList = productList.sort( function (a, b) { return b.price - a.price});
                return sortList;
            })();
        """)

        all_items_list.extend(items_list)


    # print(items_list)
    print(f"Total items: {len(all_items_list)}")
    df = pd.DataFrame(all_items_list)

    if not df.empty:
        df.to_excel('store.xlsx', index=False)

        workbook = load_workbook('store.xlsx')
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

        workbook.save('store.xlsx')

    else:
        print("DataFrame is empty, nothing to export.")

    await browser.close()

async def main():
    async with async_playwright() as p:
        await run(p)

asyncio.run(main())
print(f"Time taken: {time.time() - start} seconds")