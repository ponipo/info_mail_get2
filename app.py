import streamlit as st
import pandas as pd
import threading
import concurrent.futures
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from urllib.parse import urlparse
import re
from datetime import datetime
from io import BytesIO
import base64
import time

# スレッドごとにWebDriverを保持するためのthreading.localオブジェクト
thread_local = threading.local()

def get_driver():
    driver = getattr(thread_local, 'driver', None)
    if driver is None:
        # ヘッドレスChromeのオプションを設定
        options = Options()
        options.add_argument('--headless')
        options.add_argument('--disable-gpu')
        options.add_argument('--no-sandbox')
        driver = webdriver.Chrome(options=options)
        thread_local.driver = driver
    return driver

def close_driver():
    driver = getattr(thread_local, 'driver', None)
    if driver is not None:
        driver.quit()
        del thread_local.driver

# コードAの機能
def main_codeA():
    st.header('URL取得アプリ')
    uploaded_file = st.file_uploader("企業名が含まれるExcelファイルをアップロードしてください", type=['xlsx'], key='uploadA')
    if uploaded_file is not None:
        df = pd.read_excel(uploaded_file)
        total_items = len(df)
        st.write(f"データ数: {total_items}")
        if st.button('開始', key='startA'):
            progress_bar = st.progress(0)
            progress_text = st.empty()
            final_results = []
            num_workers = 5
            current_progress = 0
            with concurrent.futures.ThreadPoolExecutor(max_workers=num_workers) as executor:
                futures = {executor.submit(scrape_single_company_codeA, row): index for index, row in df.iterrows()}
                for future in concurrent.futures.as_completed(futures):
                    result = future.result()
                    final_results.append(result)
                    current_progress += 1
                    progress_percentage = current_progress / total_items
                    progress_bar.progress(progress_percentage)
                    progress_text.text(f'処理中: {current_progress}/{total_items} ({progress_percentage*100:.2f}%)')
            # すべてのドライバーを閉じる
            close_driver()
            # 結果のDataFrameを作成し、並び替え
            results_df = pd.DataFrame(final_results, columns=["No.", "会社名", "HP"])
            results_df.sort_values("No.", inplace=True)
            results_df.reset_index(drop=True, inplace=True)
            # 結果を表示
            st.success("完了")
            st.dataframe(results_df)
            # 結果をExcelに保存し、ダウンロードリンクを作成
            output = BytesIO()
            writer = pd.ExcelWriter(output, engine='xlsxwriter')
            results_df.to_excel(writer, index=False)
            writer.close()
            processed_data = output.getvalue()
            b64 = base64.b64encode(processed_data).decode()
            timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
            href = f'<a href="data:application/octet-stream;base64,{b64}" download="info取得_{timestamp}.xlsx">結果をダウンロード</a>'
            st.markdown(href, unsafe_allow_html=True)

def scrape_single_company_codeA(row):
    try:
        driver = get_driver()
        company_name = row['企業名']
        # Bingで検索
        search_query = f"{company_name}"
        driver.get("https://www.bing.com/search?q=" + search_query)
        # 要素が表示されるまで待機
        try:
            element = WebDriverWait(driver, 1).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, 'li.b_algo h2 a'))
            )
            url = element.get_attribute('href')
        except Exception:
            url = "URLが見つかりませんでした。"
    except Exception as e:
        url = "URLが見つかりませんでした。"
        print(f"{row.name + 1}行目でエラーが発生しました: {e}")
    return [row.name + 1, company_name, url]

# コードBの機能
def main_codeB():
    st.header('Mail取得アプリ')
    uploaded_file = st.file_uploader("会社名とHPが含まれるExcelファイルをアップロードしてください", type=['xlsx'], key='uploadB')
    if uploaded_file is not None:
        df = pd.read_excel(uploaded_file)
        total_items = len(df)
        st.write(f"データ数: {total_items}")
        if st.button('開始', key='startB'):
            progress_bar = st.progress(0)
            progress_text = st.empty()
            final_results = []
            num_workers = 5
            current_progress = 0
            with concurrent.futures.ThreadPoolExecutor(max_workers=num_workers) as executor:
                futures = {executor.submit(scrape_single_row_codeB, row): index for index, row in df.iterrows()}
                for future in concurrent.futures.as_completed(futures):
                    result = future.result()
                    final_results.append(result)
                    current_progress += 1
                    progress_percentage = current_progress / total_items
                    progress_bar.progress(progress_percentage)
                    progress_text.text(f'処理中: {current_progress}/{total_items} ({progress_percentage*100:.2f}%)')
            # すべてのドライバーを閉じる
            close_driver()
            # 結果のDataFrameを作成し、並び替え
            results_df = pd.DataFrame(final_results, columns=["No.", "会社名", "HP", "@メール", "取得メールアドレス", "取得結果"])
            results_df.sort_values("No.", inplace=True)
            results_df.reset_index(drop=True, inplace=True)
            # 結果を表示
            st.success("完了")
            st.dataframe(results_df)
            # 結果をExcelに保存し、ダウンロードリンクを作成
            output = BytesIO()
            writer = pd.ExcelWriter(output, engine='xlsxwriter')
            results_df.to_excel(writer, index=False)
            writer.close()
            processed_data = output.getvalue()
            b64 = base64.b64encode(processed_data).decode()
            timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
            href = f'<a href="data:application/octet-stream;base64,{b64}" download="実行結果_{timestamp}.xlsx">結果をダウンロード</a>'
            st.markdown(href, unsafe_allow_html=True)

def scrape_single_row_codeB(row):
    driver = get_driver()
    try:
        company_name = row['会社名']
        url = row['HP']
        domain = urlparse(url).netloc
        domain = re.sub(r'^(www[0-9]?.)?', '', domain)
        at_domain = "@" + domain
        # Bingで検索
        search_query = f"{company_name} {at_domain}"
        driver.get("https://www.bing.com/search?q=" + search_query)
        # ページがロードされるのを1秒待つ
        time.sleep(1)
        # ページ内のメールアドレスを抽出
        email_text = driver.find_element(By.TAG_NAME, 'body').text
        email_pattern = r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}'
        email_matches = re.findall(email_pattern, email_text)
        email = email_matches[0] if email_matches else "メールアドレスが見つかりませんでした。"
        result = "成功" if email_matches else "失敗"
    except Exception as e:
        email = "メールアドレスが見つかりませんでした。"
        result = "失敗"
        print(f"{row.name + 1}行目でエラーが発生しました: {e}")
    return [row.name + 1, company_name, url, at_domain, email, result]

def main():
    main_codeA()
    main_codeB()

if __name__ == '__main__':
    main()
