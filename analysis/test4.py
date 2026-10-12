import pandas as pd
import os
import time

# ==============================================================================
# 1. 파일 경로 설정 (사용자 경로 그대로 유지)
# ==============================================================================
FILE_PATH = r'data/[수정본]교양교과목록_최종.xlsx'
OUTPUT_FILENAME = r"data/IT_AI_교양과목_전체목록_수정.xlsx"

# ==============================================================================
# 2. 검색 키워드 수정 (오해의 소지가 있는 단어 전면 삭제)
# ==============================================================================
# [삭제된 키워드]: 기술, 정보, 디지털, 스마트, 통신, 지능, 데이터
# -> 이런 단어들은 '태권도 기술', '생활 정보', '정서 지능' 등에도 쓰이기 때문에 뺐습니다.

KEYWORDS_IT = [
    'SW', '소프트웨어', '컴퓨터', '코딩', '프로그래밍', 
    '네트워크', '웹', '앱', '메타버스', '사물인터넷', 'IoT', 
    '엑셀', '파이썬', 'C언어', 'Java', 'Linux', '리눅스',
    'VR', 'AR', '가상현실', '증강현실', '3D', '프린팅'
]

KEYWORDS_AI = [
    'AI', '인공지능', '머신러닝', '딥러닝', '빅데이터', '알고리즘', 
    '로봇', 'Chat', 'GPT', '4차산업', 'LLM', '생성형'
]

def load_data(path):
    if not os.path.exists(path):
        print(f"❌ 파일을 찾을 수 없습니다: {path}")
        return None
    try:
        if path.endswith(('.xlsx', '.xls')):
            return pd.read_excel(path)
        else:
            for enc in ['utf-8', 'cp949', 'euc-kr']:
                try:
                    return pd.read_csv(path, encoding=enc)
                except UnicodeDecodeError:
                    continue
    except Exception as e:
        print(f"❌ 파일 열기 실패: {e}")
        return None

def get_matching_courses(df, year, keywords):
    """해당 연도에서 키워드가 포함된 과목 리스트 반환"""
    year_df = df[df['학년도'] == year]
    mask = pd.Series(False, index=year_df.index)
    for keyword in keywords:
        # 대소문자 구분 없이 검색
        mask |= year_df['교과목명'].astype(str).str.contains(keyword, case=False)
        mask |= year_df['교과설명'].astype(str).str.contains(keyword, case=False)
    return sorted(year_df[mask]['교과목명'].unique())

def main():
    print("📂 데이터 로딩 중...")
    df = load_data(FILE_PATH)
    if df is None: return

    df['교과목명'] = df['교과목명'].fillna("")
    df['교과설명'] = df['교과설명'].fillna("")
    
    years = sorted(df['학년도'].dropna().unique())
    
    all_results = []

    print("\n" + "="*60)
    print("📊 [엄격 기준] IT 및 AI 관련 교양과목 목록")
    print("   (범용적인 단어 '기술', '정보' 등은 제외했습니다)")
    print("="*60)

    for year in years:
        print(f"\n📅 {int(year)}학년도")
        print("-" * 30)
        
        # 1. IT 과목 추출
        it_courses = get_matching_courses(df, year, KEYWORDS_IT)
        print(f"   💻 IT 관련 ({len(it_courses)}개):")
        for course in it_courses:
            print(f"      - {course}")
            all_results.append({
                '학년도': int(year),
                '구분': 'IT',
                '교과목명': course
            })

        print("") 

        # 2. AI 과목 추출
        ai_courses = get_matching_courses(df, year, KEYWORDS_AI)
        print(f"   🤖 AI 관련 ({len(ai_courses)}개):")
        for course in ai_courses:
            print(f"      - {course}")
            all_results.append({
                '학년도': int(year),
                '구분': 'AI',
                '교과목명': course
            })
            
    # 엑셀 파일로 저장 (파일 열림 오류 방지 로직 포함)
    if all_results:
        result_df = pd.DataFrame(all_results)
        result_df = result_df.sort_values(by=['학년도', '구분', '교과목명'])
        
        save_name = OUTPUT_FILENAME
        try:
            result_df.to_excel(save_name, index=False)
        except PermissionError:
            # 파일이 열려있으면 이름 뒤에 시간 붙여서 저장
            timestamp = time.strftime("%H%M%S")
            save_name = OUTPUT_FILENAME.replace(".xlsx", f"_{timestamp}.xlsx")
            print(f"\n⚠️ 기존 파일이 열려있어 '{save_name}'으로 저장합니다.")
            result_df.to_excel(save_name, index=False)
            
        print("\n" + "="*60)
        print(f"💾 저장 완료: {save_name}")
        print("="*60)
    else:
        print("\n⚠️ 검색된 과목이 없습니다.")

if __name__ == "__main__":
    main()