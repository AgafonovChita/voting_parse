import json
from base64 import b64decode
from hashlib import md5
from io import BytesIO
from pathlib import Path
from time import sleep, time
from typing import List, Optional
import pandas
import pytesseract
from PIL import Image
from selenium.webdriver import Chrome, ChromeOptions
import pytesseract

pytesseract.pytesseract.tesseract_cmd = "C:\\Program Files\\Tesseract-OCR\\tesseract.exe"

def init_driver() -> Chrome:
    options = ChromeOptions()
    options.headless = True
    driver = Chrome(options=options)
    driver.set_window_size(60000, 6000)
    return driver


def ocr_cell(cell, save_name: Optional[str] = None) -> List[str]:
    image = Image.open(BytesIO(b64decode(cell.screenshot_as_base64)))
    #image.save(Path('screenshots') / f'{save_name}.png')

    text = pytesseract.image_to_string(image, config='--psm 6 -c "tessedit_char_whitelist=0123456789.%"')
    lines = text.strip().splitlines()
    if len(lines) >= 3 and lines[-1] == '%':
        lines = (*lines[:-2], lines[-2] + lines[-1])
    return lines


def cut_numeration(text: str) -> str:
    if '. ' in text[:10]:
        text = text.split('. ', 1)[1]
    return text


def gen_img_name(url: str, num1: int, num2: Optional[int] = None) -> str:
    url_hash = md5(url.encode()).hexdigest()[:16]
    if num2 is None:
        return f'{url_hash}_{num1:03d}'
    return f'{url_hash}_{num1:03d}_{num2:03d}'



def get_page_data(driver: Chrome, url: str, debug=False) -> List[List[str]]:
    driver.delete_all_cookies()
    t1 = time()
    driver.get(url)
    t2 = time()
    print(f'load: {t2 - t1:.2f} sec')
    sleep(1)

    # increase readability
    script = '''
        document.styleSheets[0].insertRule("td {
            font-size: 2em !important;
            max-width: 100% !important;
            background-color: #ffffff !important;
        }", 0 );
        document.styleSheets[0].insertRule("table {
            height: max-content !important;
            width: max-content !important;
        }", 0 )
    '''
    
    driver.execute_script(script.replace('\n', ''))

    results = []

    # text above table
    head_rows = driver.find_elements_by_css_selector('.table-borderless tr')
    for row in head_rows:
        cols = row.find_elements_by_tag_name('td')
        line = []
        for col in cols:
            text = col.text.strip()
            if text:
                line.append(text)
        if line:
            print(line)
            results.append(line)

    small_tables = driver.find_elements_by_css_selector('.table-responsive')
    big_tables = driver.find_elements_by_css_selector('.table-bordered')

    if small_tables:
        print('МАЛЕНЬКАЯ ТАБЛИЦА')
        table = small_tables[0]
        rows = table.find_elements_by_css_selector('tr')
        for i, row in enumerate(rows):
            cells = row.find_elements_by_tag_name('td')
            if len(cells) < 3:  # skip empty lines
                continue
            name = cut_numeration(cells[1].text.strip())
            print(name)
            results.append(name, *ocr_cell(cells[2], save_name=gen_img_name(url, i)))

    elif big_tables:
        print('Большая таблица')
        table = big_tables[0]
        headers = table.find_elements_by_css_selector('th')
        results.append([h.text for h in headers][1:])
        print(results)
        rows = table.find_elements_by_css_selector('tr')
        for i, row in enumerate(rows):
            print(f'row {i + 1} / {len(rows)}')
            cells = row.find_elements_by_css_selector('td')
            if len(cells) < 3:  # skip empty lines
                continue
            name_row = [cut_numeration(cells[1].text)]
            result = [cut_numeration(cells[1].text)]
            result1 = [cut_numeration(cells[1].text)]
            print(result)
            for j, cell in enumerate(cells[2:]):
                #if i+1<14:
                result.append(int(*ocr_cell(cell, save_name=gen_img_name(url, i, j))))
                   
                # if i+1>13:
                    
                #     num = [*ocr_cell(cell, save_name=gen_img_name(url, i, j))]
                #     resul1=result
                #     #result.append(' / '.join(num))
                #     result.append(num[0])
                #     result1.append(num[1])
                    
            results.append(result) 
            # if i+1<14:
               
            # if i+1>13:
            #     results.append(result)
            #     results.append(result1)


            

    else:
        print('no tables found')

    t3 = time()
    print(f'processed: {t3 - t2:.2f} sec')
    return results


def main():
    driver = init_driver()
    results = {}
    id_list = range(27520001371100, 27520001371139)
    for id_tvd in id_list:
        print('____________________________')
        print(id_tvd)
    #id_tvd = 27520001371105
        url ='http://www.zabkray.vybory.izbirkom.ru/region/izbirkom?action=show&root=1000303&tvd='+str(id_tvd)+'&vrn=100100225883172&prver=0&pronetvd=null&region=92&sub_region=92&type=464&report_mode=null'
        try:
            results=get_page_data(driver, url)
        except Exception as e:
            print(f'ERR: {e!r}')
      
        name_json='C:\\Users\\School №47\\Desktop\\voting_parse\\json\\'+str(id_tvd)+'.json'
        with open(name_json, 'wt') as f:
            json.dump(results, f)
        
        name_data ='C:\\Users\\School №47\\Desktop\\voting_parse\\data\\'+str(id_tvd)+'.xlsx'
        pandas.read_json(name_json).to_excel(name_data)

main()
