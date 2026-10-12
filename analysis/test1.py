import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
from collections import Counter
import re
import os
import urllib.request
import ssl

# ==============================================================================
# 1. 파일 경로 설정
# ==============================================================================
FILE_PATH = r'data/[수정본]교양교과목록_최종.xlsx'

# ==============================================================================
# 2. 폰트 자동 다운로드 (깨짐 방지 100% 해결책)
# ==============================================================================
def download_font():
    font_filename = "NanumGothic.ttf"
    font_url = "https://github.com/google/fonts/raw/main/ofl/nanumgothic/NanumGothic-Regular.ttf"
    
    if os.path.exists(font_filename):
        return font_filename
    
    print("⬇️ 폰트 다운로드 중... (NanumGothic)")
    try:
        context = ssl._create_unverified_context()
        with urllib.request.urlopen(font_url, context=context) as response, open(font_filename, 'wb') as out_file:
            out_file.write(response.read())
        return font_filename
    except Exception as e:
        print(f"❌ 폰트 다운로드 실패: {e}")
        return None

# 폰트 준비
FONT_PATH = download_font()
if FONT_PATH:
    FONT_PROP = fm.FontProperties(fname=FONT_PATH)
else:
    FONT_PROP = None # 실패 시 시스템 기본 폰트 사용

# ==============================================================================
# 3. 강력한 전처리 (의미 없는 단어 싹둑)
# ==============================================================================
STOP_WORDS = set([
    '교과목', '수업', '학습', '이해', '학생', '과목', '능력', '목표', '목적', '과정', '내용',
    '강의', '습득', '배양', '함양', '중심', '통한', '구성', '활용', '적용', '기반', '기초',
    '실제', '다양한', '지식', '대해', '관련', '제공', '향상', '강화', '소개', '개요',
    '및', '통해', '위한', '하는', '있는', '대한', '관한', '하여', '따라', '의해', '등',
    '수', '것', '것을', '등을', '경우', '함', '임', '됨', '그', '이', '저', '따라서',
    '한다', '합니다', '된다', '됩니다', '할', '될', '해', '된', '나', '와', '과',
    '연구', '개발', '문제', '해결', '분석', '방법', '이론', '실습', '살펴본다', '알아본다',
    '사고', '창의', '통찰', '역량', '대비'
])

def preprocess_text(text):
    if pd.isna(text): return ""
    text = re.sub(r'[^가-힣a-zA-Z0-9\s]', ' ', str(text))
    words = text.split()
    meaningful_words = []
    
    suffixes = ['합니다', '하는', '한다', '했다', '할', '해', '됩니다', '되는', '된다', '될', '돼',
                '입니다', '이다', '인', '시키다', '시킨다', '시켜', '가지다', '가진다', '있는', '있다', '없다',
                '으로', '에서', '보다', '부터', '까지', '이나', '에는', '과는', '은', '는', '이', '가', '을', '를', '의', '에', '로', '와', '과', '도']

    for word in words:
        if len(word) < 2: continue
        cleaned = word
        
        for _ in range(2): # 꼬리 자르기 2회 반복
            for suffix in suffixes:
                if cleaned.endswith(suffix):
                    if len(cleaned) - len(suffix) >= 2:
                        cleaned = cleaned[:-len(suffix)]
                        break
        
        if len(cleaned) < 2: continue
        if cleaned in STOP_WORDS or word in STOP_WORDS: continue
        meaningful_words.append(cleaned)
            
    return " ".join(meaningful_words)

# ==============================================================================
# 4. 시각화 함수들 (막대 + 꺾은선)
# ==============================================================================
def draw_bar_chart(year, top_keywords):
    """(기존 요청) 가로 막대 그래프"""
    if not top_keywords: return
    words, counts = zip(*top_keywords)
    
    plt.figure(figsize=(10, 6))
    plt.barh(list(reversed(words)), list(reversed(counts)), color='skyblue')
    plt.title(f"{year}년 핵심 키워드 빈도 (TOP 20)", fontsize=15, fontproperties=FONT_PROP)
    plt.xlabel("빈도수", fontproperties=FONT_PROP)
    plt.yticks(fontproperties=FONT_PROP) # Y축 글자 폰트 적용
    plt.grid(axis='x', linestyle='--', alpha=0.5)
    plt.tight_layout()
    plt.show()

def draw_line_chart(trend_data):
    """(추가 기능) 연도별 키워드 트렌드 꺾은선 그래프"""
    # 전체 연도에서 가장 많이 나온 상위 5개 단어만 추적
    total_counter = Counter()
    for year in trend_data:
        total_counter.update(trend_data[year])
    
    top_5_keywords = [word for word, count in total_counter.most_common(5)]
    
    plt.figure(figsize=(12, 6))
    years = sorted(trend_data.keys())
    
    for word in top_5_keywords:
        counts = [trend_data[y].get(word, 0) for y in years]
        plt.plot(years, counts, marker='o', label=word)
    
    plt.title("연도별 핵심 키워드(TOP 5) 등장 빈도 변화", fontsize=15, fontproperties=FONT_PROP)
    plt.xlabel("학년도", fontproperties=FONT_PROP)
    plt.ylabel("빈도수", fontproperties=FONT_PROP)
    plt.legend(prop=FONT_PROP) # 범례 폰트 적용
    plt.xticks(years)
    plt.grid(True, linestyle='--', alpha=0.5)
    plt.tight_layout()
    plt.show()

# ==============================================================================
# 5. 메인 실행
# ==============================================================================
def main():
    if not os.path.exists(FILE_PATH):
        print(f"❌ 파일을 찾을 수 없습니다: {FILE_PATH}")
        return

    try:
        if FILE_PATH.endswith(('.xlsx', '.xls')):
            df = pd.read_excel(FILE_PATH)
        else:
            df = pd.read_csv(FILE_PATH, encoding='utf-8')
    except Exception as e:
        print(f"❌ 파일 열기 실패: {e}")
        return

    target_col = '교과설명' if '교과설명' in df.columns else '교과목명'
    years = sorted(df['학년도'].dropna().unique())
    
    trend_data = {} # 꺾은선 그래프용 데이터 저장소

    print("📊 시각화 생성 시작 (창을 닫으면 다음 그래프가 나옵니다)")

    for year in years:
        year_texts = df[df['학년도'] == year][target_col].tolist()
        
        all_words = []
        for t in year_texts:
            all_words.extend(preprocess_text(t).split())
            
        count = Counter(all_words)
        trend_data[year] = count # 저장
        
        top_20 = count.most_common(20)
        
        print(f"\n📅 {int(year)}년 막대 그래프 생성")
        draw_bar_chart(int(year), top_20)

    # 마지막에 전체 트렌드 꺾은선 그래프 출력
    print("\n📈 연도별 트렌드(꺾은선) 그래프 생성")
    draw_line_chart(trend_data)

if __name__ == "__main__":
    main()