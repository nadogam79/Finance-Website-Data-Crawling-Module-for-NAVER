import pandas as pd
import requests
import time
import sys
import re
from bs4 import BeautifulSoup

# --- [1. 설정 및 경로] ---
file_path = r'C:\Users\chick\source\repos\chartmaker\company_list.xlsx'
esg_file_path = r'C:\Users\chick\source\repos\chartmaker\(2023~2025)_esg_rate.xlsx'
save_path = r'C:\Users\chick\source\repos\chartmaker\수집결과_최종_통합.xlsx'

# --- [2. 데이터 로드] ---
try:
    df_input = pd.read_excel(file_path)
    company_list = []
    for _, row in df_input.iterrows():
        raw_code = str(row['c_id']).split('.')[0].strip().zfill(6)
        company_list.append({'code': raw_code, 'name': str(row['c_name']).strip()})
    
    esg_master_df = pd.read_excel(esg_file_path)
    esg_master_df['m_code'] = esg_master_df.iloc[:, 2].apply(lambda x: str(x).split('.')[0].strip().zfill(6))
    esg_master_df['m_year'] = esg_master_df.iloc[:, 7].apply(lambda x: str(int(float(x))) if pd.notna(x) else "")
    esg_master_df['m_grade'] = esg_master_df.iloc[:, 3].astype(str).str.strip()
    print("데이터 로드 완료")
except Exception as e:
    print(f"❌ 로드 실패: {e}")
    sys.exit()

# --- [3. 재무 함수: 연간/분기는 제외] ---
def get_finance_data(mCode):
    url = f"https://finance.naver.com/item/main.naver?code={mCode}"
    headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36'}
    
    try:
        response = requests.get(url, headers=headers, timeout=10)
        soup = BeautifulSoup(response.text, 'html.parser')
        finance_area = soup.select_one('.section.cop_analysis')
        if not finance_area: return None

        # [1] 헤더 분석: 날짜가 써진 두 번째 tr(row)의 th(헤드, 항목명)들만 추출  
        header_row = finance_area.select('thead tr')[1]
        ths = header_row.select('th')
        headers_text = [th.get_text(strip=True) for th in ths]
        
        # [2] 인덱스 찾기: 정확히 '연간' 영역(보통 앞 4~5개) 내에서만 검색
        # 정규식을 써서 '2023.12', '2024.12', '2025.12' 형태만 찾음
        idx_23, idx_24, idx_25 = -1, -1, -1
        for i, txt in enumerate(headers_text):
            if i > 4: break # 5번째 칸부터는 분기 영역이므로 차단
            if '2023.12' in txt: idx_23 = i
            elif '2024.12' in txt: idx_24 = i
            elif '2025.12' in txt: idx_25 = i

        # [3] 데이터 추출
        rows = finance_area.select('tbody tr')
        extracted_data = []

        for row in rows:
            # td는 수치
            title = row.select_one('th').get_text(strip=True)
            tds = [td.get_text(strip=True) for td in row.select('td')]
            
            if any(target in title for target in ['ROE', '부채비율']):
                # td만 모았으므로 인덱스는 (찾은 인덱스 - 1)이 됨 (th(헤드)가 빠졌으므로)
                res = {
                    '항목': title,
                    '2023': tds[idx_23] if idx_23 != -1 and len(tds) > idx_23 else '-',
                    '2024': tds[idx_24] if idx_24 != -1 and len(tds) > idx_24 else '-',
                    '2025': tds[idx_25] if idx_25 != -1 and len(tds) > idx_25 else '-'
                }
                extracted_data.append(res)
        return extracted_data
    except:
        return None

# --- [4. 메인 루프] ---
final_rows = []
print("[수집 시작]")

for i, company in enumerate(company_list):
    code, name = company['code'], company['name']
    finance_results = get_finance_data(code)
    
    esg_dict = {'2023': '-', '2024': '-', '2025': '-'}
    matches = esg_master_df[esg_master_df['m_code'] == code]
    for _, row in matches.iterrows():
        y, g = row['m_year'], row['m_grade']
        if y in esg_dict: esg_dict[y] = g

    if finance_results:
        for f in finance_results:
            final_rows.append({
                '종목코드': code, '종목명': name, '항목': f['항목'],
                '재무_2023': f['2023'], '재무_2024': f['2024'], '재무_2025(E)': f['2025'],
                'ESG_2023': esg_dict['2023'], 'ESG_2024': esg_dict['2024'], 'ESG_2025': esg_dict['2025']
            })
        print(f"✅ [{i+1}] {name}({code}) 완료")
    time.sleep(0.1)

# --- [5. 저장] ---
if final_rows:
    pd.DataFrame(final_rows).to_excel(save_path, index=False)
    print(f"\n🎉 추출이 성공적으로 완료되었습니다! 파일: {save_path}")