import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm  # 폰트 매니저를 fm으로 import
from wordcloud import WordCloud
import re
import os
import urllib.request
import ssl

# ==============================================================================
# 1. 파일 경로 설정 (사용자 경로)
# ==============================================================================
FILE_PATH = r'data/[수정본]교양교과목록_최종.xlsx'

# ==============================================================================
# 2. 폰트 파일 다운로드
# ==============================================================================
def download_font():
    font_filename = "NanumGothic.ttf"
    # GitHub Raw 링크 사용 (안정적)
    font_url = "https://github.com/google/fonts/raw/main/ofl/nanumgothic/NanumGothic-Regular.ttf"
    
    if os.path.exists(font_filename):
        print(f"✅ 폰트 파일이 이미 있습니다: {font_filename}")
        return font_filename
    
    print("⬇️ 한글 폰트(나눔고딕)를 다운로드 중입니다... 잠시만 기다려주세요.")
    try:
        context = ssl._create_unverified_context()
        with urllib.request.urlopen(font_url, context=context) as response, open(font_filename, 'wb') as out_file:
            data = response.read()
            out_file.write(data)
        print("✅ 폰트 다운로드 완료!")
        return font_filename
    except Exception as e:
        print(f"❌ 폰트 다운로드 실패: {e}")
        return None

# 폰트 준비
FONT_PATH = download_font()

# ==============================================================================
# 3. 텍스트 전처리 (불용어 및 어미 삭제)
# ==============================================================================
# ==============================================================================
# 3. 텍스트 전처리 (불용어 및 어미 삭제)
# ==============================================================================

# 1. 불용어(Stopwords): 분석 결과에서 아예 빼버릴 단어들
# (여기에 '필요 없는 단어'를 계속 추가하면 워드클라우드에서 사라집니다)
STOP_WORDS = set([
    # (1) 일반적인 교육 용어 (너무 흔해서 의미 없음)
    '교과목', '수업', '학습', '이해', '학생', '과목', '능력', '목표', '목적', '과정', '내용',
    '강의', '습득', '배양', '함양', '중심', '통한', '구성', '활용', '적용', '기반', '기초',
    '실제', '다양한', '지식', '대해', '관련', '제공', '향상', '강화', '소개', '개요',
    '연구', '개발', '문제', '해결', '분석', '방법', '이론', '실습', '살펴본다', '알아본다',
    '사고', '창의', '통찰', '역량', '대비', '자신', '표현', '바탕',
    
    # (2) 문법적 요소 & 접속사
    '및', '통해', '위한', '하는', '있는', '대한', '관한', '하여', '따라', '의해', '등',
    '수', '것', '것을', '등을', '경우', '함', '임', '됨', '그', '이', '저', '따라서',
    '한다', '합니다', '된다', '됩니다', '할', '될', '해', '된', '나', '와', '과',
    '위하여', '통하여', '대하여', '이러한', '있도록', '하고', '그리고', '그러나',
    ' , ', ' 관련하여', '각', '의', '이에', '을', '적절한', '...', '에', '직접', '나가' ,
    '하면', '으로는', '등에서', '현장', '바른', '정신을', '목적', '잇다', '이', '강좌는',
    '기반하여', '이는', '재고하는', '표현', '사용', '학생들', '필요한', '배양함', '위해', '있다'
    ])

# 2. 전처리 함수: 텍스트를 청소하고 '명사'만 남기는 로직
def preprocess_text(text):
    if pd.isna(text): return ""
    
    # [1단계] 특수문자 제거 (한글, 영문, 숫자만 남김)
    text = re.sub(r'[^가-힣a-zA-Z0-9\s]', ' ', str(text))
    words = text.split()
    meaningful_words = []
    
    # 제거할 꼬리말(Suffixes) 목록
    # (팁: 긴 단어부터 검사해야 정확도가 높습니다. 예: '합니다'를 먼저 검사하고 '다'를 검사)
    suffixes = [
        '합니다', '하는', '한다', '했다', '할', '해', '하니', '하고',
        '됩니다', '되는', '된다', '됐다', '될', '돼', '되니',
        '입니다', '이다', '이며', '인',
        '시키다', '시킨다', '시키는', '시켜',
        '가지다', '가진다', '가지는',
        '있습니다', '있다', '있는', '있어', '있도록',
        '없습니다', '없다', '없는', '없어',
        '으로', '에서', '보다', '부터', '까지',
        '이나', '에는', '과는', '와는', '의',
        '은', '는', '이', '가', '을', '를', '에', '로', '와', '과', '도'
    ]

    for word in words:
        if len(word) < 2: continue # 1글자 단어는 무시 (예: '그', '저')
        
        cleaned = word
        
        # [2단계] 꼬리 자르기 (Loop)
        # 예: "학습하는데" -> (1차) "학습하" -> (2차) "학습"
        # 꼬리가 여러 개 붙을 수 있으므로 3번 정도 반복해서 깎습니다.
        for _ in range(3): 
            for suffix in suffixes:
                if cleaned.endswith(suffix):
                    # 꼬리를 잘랐을 때 남은 단어가 2글자 이상이어야 자름 (예: '가는' -> '가'(X))
                    if len(cleaned) - len(suffix) >= 2:
                        cleaned = cleaned[:-len(suffix)]
                        break # 하나 자르면 루프 탈출 (다시 처음부터 검사)
        
        # [3단계] 불용어 필터링
        if len(cleaned) < 2: continue # 자르고 났더니 1글자면 버림
        if cleaned in STOP_WORDS: continue # 정리된 단어가 금지어 목록에 있으면 버림
        if word in STOP_WORDS: continue # 원본 단어가 금지어 목록에 있어도 버림
        
        meaningful_words.append(cleaned)
            
    return " ".join(meaningful_words)

# ==============================================================================
# 4. 메인 실행
# ==============================================================================
def main():
    if FONT_PATH is None:
        print("❌ 폰트 다운로드 실패. 인터넷 연결을 확인하거나 수동 다운로드를 시도하세요.")
        return

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
    
    # 폰트 속성 객체 미리 생성
    font_prop = fm.FontProperties(fname=FONT_PATH)
    
    print(f"📊 총 {len(years)}개 연도 워드클라우드 생성 시작")
    print(f"🔤 적용 폰트: {os.path.abspath(FONT_PATH)}")

    for year in years:
        print(f"\n⏳ {int(year)}년도 처리 중...")
        year_texts = df[df['학년도'] == year][target_col].tolist()
        
        cleaned_text_list = [preprocess_text(t) for t in year_texts]
        full_text = " ".join(cleaned_text_list)
        
        if not full_text.strip():
            print(f"  ⚠️ {year}년 텍스트 없음")
            continue
            
        wc = WordCloud(
            width=1000,
            height=600,
            background_color='white',
            font_path=FONT_PATH,
            max_words=80,
            min_font_size=10,
            colormap='Set2',
            font_step=1
        ).generate(full_text)
        
        plt.figure(figsize=(10, 6))
        plt.imshow(wc, interpolation='bilinear')
        plt.axis('off')
        
        # [수정됨] plt.font_manager 대신 fm.FontProperties 사용
        plt.title(f"{int(year)}년 핵심 키워드", fontsize=15, fontproperties=font_prop)
        plt.show()
        print(f"✅ {int(year)}년 완료 (그래프 창을 닫아야 다음 연도가 나옵니다)")

if __name__ == "__main__":
    main()