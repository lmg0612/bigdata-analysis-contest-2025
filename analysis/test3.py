import pandas as pd
from collections import Counter
import re
import os

# ==============================================================================
# 1. 파일 경로 설정
# ==============================================================================
FILE_PATH = r'data/[수정본]교양교과목록_최종.xlsx'

# ==============================================================================
# 2. 강력한 전처리 (불용어 제거)
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
        for _ in range(2):
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
# 3. 키워드 포함 과목 찾기 함수
# ==============================================================================
def find_courses_by_keyword(df, year, keyword):
    """
    교과목명 또는 교과설명에 해당 키워드가 포함된 과목들을 리스트로 반환
    """
    year_df = df[df['학년도'] == year]
    
    # 검색 조건: 교과목명 OR 교과설명에 키워드 포함
    mask = (year_df['교과목명'].astype(str).str.contains(keyword, na=False)) | \
           (year_df['교과설명'].astype(str).str.contains(keyword, na=False))
    
    found_courses = year_df[mask]['교과목명'].unique()
    return list(found_courses)

# ==============================================================================
# 4. 메인 실행
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

    print("="*60)
    print("📊 [연도별 핵심 키워드 및 관련 강좌 목록]")
    print("="*60)

    for year in years:
        # 1. 키워드 추출
        year_texts = df[df['학년도'] == year][target_col].tolist()
        all_words = []
        for t in year_texts:
            all_words.extend(preprocess_text(t).split())
            
        count = Counter(all_words)
        top_3 = count.most_common(3) # 상위 3개만 추출
        
        print(f"\n📅 {int(year)}학년도 분석 결과")
        print("-" * 60)
        
        # 2. 강좌 목록 출력
        for rank, (keyword, freq) in enumerate(top_3, 1):
            print(f"🏆 {rank}위 키워드: [{keyword}] (언급 빈도: {freq}회)")
            
            related_courses = find_courses_by_keyword(df, int(year), keyword)
            
            if related_courses:
                print(f"   📂 관련 강좌 ({len(related_courses)}개):")
                # 보기 좋게 5개씩 끊어서 출력하거나 리스트업
                for i, course in enumerate(related_courses):
                    print(f"      - {course}")
            else:
                print("   ⚠️ (해당 키워드가 포함된 정확한 과목명을 찾지 못함)")
            print("") # 줄바꿈
            
        print("=" * 60)

if __name__ == "__main__":
    main()