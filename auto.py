import asyncio
import pandas as pd
from pyppeteer import launch


async def process_row(page, row):
    await page.goto('https://rpa-test.pages.dev/')  # 替换为你的Cloudflare Pages URL

    # 等待页面加载完成
    await page.waitForSelector('form#searchForm')

    patient_name = row['Patient Name']

    # 填写表单
    await page.type('input#serviceDate', row['Service Date'].strftime('%m/%d/%Y'))
    await page.type('input#patientName', row['Patient Name'].strip())
    await page.type('input#patientDOB', row['Patient DOB'].strftime('%m/%d/%Y'))
    await page.type('input#payerSubscriberNo', row['Payer Subscriber No'].strip())

    # 点击搜索按钮
    await page.click('button[type="submit"]')

    # 等待结果显示
    await page.waitForSelector('div#result')

    # 获取结果文本
    result_element = await page.querySelector('div#result')
    result_text = await page.evaluate('(element) => element.textContent', result_element)

    print(f"{patient_name}, Search result: {result_text}")


async def main():
    # 读取Excel文件
    df = pd.read_excel('demo.xlsx')

    # 启动浏览器
    browser = await launch(headless=False, args=['--no-sandbox'])
    page = await browser.newPage()

    for _, row in df.iterrows():
        await process_row(page, row)

    await browser.close()


# 运行主函数
asyncio.get_event_loop().run_until_complete(main())