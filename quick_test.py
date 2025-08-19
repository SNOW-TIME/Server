import pandas as pd
import os

def quick_test():
    """변환된 파일을 직접 테스트"""
    
    # 변환된 파일 경로
    converted_file = "data/프라임관301,수용인원 0070명,캡스톤디자인강의실(안유현강의실)_converted.xlsx"
    
    if not os.path.exists(converted_file):
        print(f"❌ 변환된 파일을 찾을 수 없습니다: {converted_file}")
        print("\n📂 data 디렉토리의 파일들:")
        if os.path.exists('data'):
            for file in os.listdir('data'):
                if '_converted.xlsx' in file:
                    print(f"  ✅ {file}")
        return
    
    try:
        print(f"📖 파일 로드 중: {os.path.basename(converted_file)}")
        
        # Excel 파일 읽기
        df = pd.read_excel(converted_file)
        
        print(f"✅ 파일 로드 성공!")
        print(f"📊 데이터 크기: {df.shape[0]}행 × {df.shape[1]}열")
        
        # 기본 정보 출력
        print(f"\n📋 컬럼 목록:")
        for i, col in enumerate(df.columns):
            print(f"  {i+1}. {col}")
        
        # 첫 몇 행 출력
        print(f"\n📄 데이터 미리보기:")
        print(df.head(3).to_string())
        
        # 시간대 컬럼 찾기
        import re
        time_pattern = r'\d{2}:\d{2}~'
        time_columns = [col for col in df.columns if isinstance(col, str) and re.match(time_pattern, col)]
        
        print(f"\n⏰ 시간대 컬럼 ({len(time_columns)}개):")
        print(f"  {time_columns[:5]}{'...' if len(time_columns) > 5 else ''}")
        
        # 날짜 정보 확인
        if '사용일자' in df.columns:
            dates = df['사용일자'].dropna().unique()
            print(f"\n📅 사용 가능한 날짜 ({len(dates)}개):")
            print(f"  {sorted(dates)[:5]}{'...' if len(dates) > 5 else ''}")
        
        # 샘플 데이터로 상태 확인
        if len(time_columns) > 0 and '사용일자' in df.columns:
            first_date = df['사용일자'].iloc[0]
            first_time = time_columns[0]
            sample_value = df[first_time].iloc[0]
            
            print(f"\n🔍 샘플 데이터:")
            print(f"  날짜: {first_date}")
            print(f"  시간: {first_time}")
            print(f"  상태: {'사용 가능' if pd.isna(sample_value) or str(sample_value).strip() == '' else f'사용중 ({sample_value})'}")
        
        print(f"\n✅ 변환된 파일이 정상적으로 작동합니다!")
        
    except Exception as e:
        print(f"❌ 오류 발생: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    quick_test()
