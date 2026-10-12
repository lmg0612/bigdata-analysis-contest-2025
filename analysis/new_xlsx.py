import pandas as pd
import os
import difflib

# ==============================================================================
# 1. 파일 경로 설정
# ==============================================================================
FILE_PATH = r'data/[수정본]교양교과목록_최종.xlsx'
OUTPUT_FILENAME = "교양_유지과목_최종분석결과.xlsx"

# ==============================================================================
# 2. 데이터 로드 및 헤더 정리
# ==============================================================================
def load_and_clean_data(path):
    if not os.path.exists(path):
        print(f"❌ 파일을 찾을 수 없습니다: {path}")
        return None
    
    try:
        if path.endswith(('.xlsx', '.xls')):
            df = pd.read_excel(path)
        else:
            df = pd.read_csv(path, encoding='utf-8')
    except:
        try:
            df = pd.read_csv(path, encoding='cp949')
        except:
            return None

    # [핵심] 컬럼명 공백 제거 (" 교과목코드 " -> "교과목코드")
    df.columns = df.columns.str.strip()
    
    # 데이터 문자열 변환 및 공백 제거
    df['교과목명_clean'] = df['교과목명'].astype(str).str.strip()
    
    if '교과목코드' in df.columns:
        df['교과목코드'] = df['교과목코드'].astype(str).str.strip()
        print("✅ '교과목코드' 컬럼을 정상적으로 찾았습니다!")
    else:
        print("⚠️ '교과목코드' 컬럼을 찾지 못했습니다. (과목명으로만 분석)")
        
    return df

# ==============================================================================
# 3. 메인 실행
# ==============================================================================
def main():
    df = load_and_clean_data(FILE_PATH)
    if df is None: return

    years = sorted(df['학년도'].dropna().unique())
    output_path = os.path.join(os.path.dirname(FILE_PATH), OUTPUT_FILENAME)
    
    print(f"\n📊 분석 시작 (기준: 이름 유사도 70% 이상 OR 코드 일치)")
    print("-" * 60)

    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        has_data = False
        
        for i in range(len(years) - 1):
            prev_year = years[i]
            curr_year = years[i+1]
            
            # 연도별 데이터 분리
            df_prev = df[df['학년도'] == prev_year].copy()
            df_curr = df[df['학년도'] == curr_year].copy()
            
            maintained_list = []
            
            # 1. 코드 기준 매칭 (가장 정확함)
            if '교과목코드' in df.columns:
                common_codes = set(df_prev['교과목코드']) & set(df_curr['교과목코드'])
                # 코드가 같은 과목들 추가
                code_matches = df_curr[df_curr['교과목코드'].isin(common_codes)]['교과목명'].tolist()
                maintained_list.extend(code_matches)
            
            # 2. 이름 유사도 기준 매칭 (코드가 달라도 이름이 비슷하면 인정)
            prev_names = df_prev['교과목명_clean'].tolist()
            curr_names = df_curr['교과목명_clean'].tolist()
            
            for p_name in prev_names:
                # 가장 비슷한 이름 찾기 (유사도 0.7 이상 = 70% 일치)
                matches = difflib.get_close_matches(p_name, curr_names, n=1, cutoff=0.6)
                if matches:
                    # 원본 이름(공백 포함)을 가져와서 리스트에 추가
                    original_name = df_curr[df_curr['교과목명_clean'] == matches[0]]['교과목명'].iloc[0]
                    maintained_list.append(original_name)

            # 중복 제거 (코드 매칭과 이름 매칭이 겹칠 수 있음)
            maintained_unique = list(set(maintained_list))
            
            print(f"🔍 [{prev_year} -> {curr_year}] 유지된 과목: {len(maintained_unique)}개 발견")
            
            if maintained_unique:
                # 엑셀 저장
                df_save = df_curr[df_curr['교과목명'].isin(maintained_unique)].drop_duplicates(subset=['교과목명'])
                sheet_name = f"{prev_year}-{curr_year}_유지"
                df_save.to_excel(writer, sheet_name=sheet_name, index=False)
                has_data = True
            else:
                print("   ⚠️ 유지된 과목이 없습니다.")

    if has_data:
        print("\n" + "="*60)
        print(f"💾 파일 저장 완료: {output_path}")
        print("이제 엑셀 파일을 열어보시면 훨씬 많은 과목이 들어있을 겁니다!")
        print("="*60)

if __name__ == "__main__":
    main()