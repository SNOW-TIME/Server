import pandas as pd
import os
import re
from pathlib import Path
from typing import List, Dict

class HTMLToExcelConverter:
    def __init__(self, data_directory="data"):
        """
        HTML 파일을 Excel로 변환하는 클래스
        
        Args:
            data_directory: HTML 파일들이 있는 디렉토리
        """
        self.data_directory = data_directory
        self.converted_files = []
        
    def find_html_files(self) -> List[str]:
        """HTML 형식의 .XLS 파일들 찾기"""
        html_files = []
        
        if not os.path.exists(self.data_directory):
            print(f"❌ 디렉토리를 찾을 수 없습니다: {self.data_directory}")
            return html_files
        
        for file in os.listdir(self.data_directory):
            if file.endswith('.XLS') or file.endswith('.xls'):
                file_path = os.path.join(self.data_directory, file)
                
                # 파일이 HTML인지 확인
                try:
                    with open(file_path, 'rb') as f:
                        first_bytes = f.read(100)
                    
                    if first_bytes.startswith(b'<html'):
                        html_files.append(file_path)
                        print(f"📋 HTML 파일 발견: {file}")
                except Exception as e:
                    print(f"❌ 파일 읽기 오류 ({file}): {e}")
        
        return html_files
    
    def convert_single_file(self, html_file_path: str) -> str:
        """단일 HTML 파일을 Excel로 변환"""
        try:
            print(f"\n🔄 변환 중: {os.path.basename(html_file_path)}")
            
            # pandas로 HTML 테이블 읽기
            tables = pd.read_html(html_file_path, encoding='utf-8', header=0)
            
            if not tables:
                print(f"❌ 테이블을 찾을 수 없습니다: {html_file_path}")
                return None
            
            # 첫 번째 테이블 사용
            df = tables[0]
            
            # 출력 파일 경로 생성
            base_name = os.path.splitext(html_file_path)[0]
            output_path = f"{base_name}_converted.xlsx"
            
            # Excel 파일로 저장
            df.to_excel(output_path, index=False, engine='openpyxl')
            
            print(f"✅ 변환 완료: {os.path.basename(output_path)}")
            print(f"   📊 데이터 크기: {df.shape[0]}행 × {df.shape[1]}열")
            
            # 데이터 미리보기
            print(f"   📋 컬럼: {list(df.columns)[:5]}{'...' if len(df.columns) > 5 else ''}")
            
            self.converted_files.append(output_path)
            return output_path
            
        except Exception as e:
            print(f"❌ 변환 실패 ({os.path.basename(html_file_path)}): {e}")
            return None
    
    def convert_all_files(self) -> List[str]:
        """모든 HTML 파일을 Excel로 변환"""
        html_files = self.find_html_files()
        
        if not html_files:
            print("❌ 변환할 HTML 파일이 없습니다.")
            return []
        
        print(f"\n📂 {len(html_files)}개의 HTML 파일을 발견했습니다.")
        print("🔄 일괄 변환을 시작합니다...\n")
        
        converted_count = 0
        failed_count = 0
        
        for html_file in html_files:
            result = self.convert_single_file(html_file)
            if result:
                converted_count += 1
            else:
                failed_count += 1
        
        print(f"\n=== 변환 완료 ===")
        print(f"✅ 성공: {converted_count}개")
        print(f"❌ 실패: {failed_count}개")
        print(f"📁 변환된 파일들:")
        
        for file_path in self.converted_files:
            print(f"   - {os.path.basename(file_path)}")
        
        return self.converted_files
    
    def create_summary_report(self):
        """변환된 파일들의 요약 보고서 생성"""
        if not self.converted_files:
            print("❌ 변환된 파일이 없습니다.")
            return
        
        summary_data = []
        
        for file_path in self.converted_files:
            try:
                # 파일명에서 정보 추출
                file_name = os.path.basename(file_path)
                
                # 건물, 호실, 수용인원 추출
                building_match = re.search(r'^([가-힣]+관)', file_name)
                room_match = re.search(r'관(\d+)', file_name)
                capacity_match = re.search(r'수용인원\s*(\d+)명', file_name)
                
                building = building_match.group(1) if building_match else "Unknown"
                room = room_match.group(1) if room_match else "Unknown"
                capacity = int(capacity_match.group(1)) if capacity_match else 0
                floor = int(room) // 100 if room.isdigit() else 0
                
                # 파일 크기
                file_size = os.path.getsize(file_path)
                
                summary_data.append({
                    '파일명': file_name,
                    '건물': building,
                    '호실': room,
                    '층': floor,
                    '수용인원': capacity,
                    '파일크기(KB)': round(file_size / 1024, 1),
                    '파일경로': file_path
                })
                
            except Exception as e:
                print(f"❌ 파일 정보 추출 실패 ({file_path}): {e}")
        
        if summary_data:
            summary_df = pd.DataFrame(summary_data)
            summary_path = os.path.join(self.data_directory, "conversion_summary.xlsx")
            summary_df.to_excel(summary_path, index=False)
            
            print(f"\n📊 요약 보고서 생성: {summary_path}")
            print(summary_df.to_string(index=False))


class UniversalClassroomParserFixed:
    """HTML 변환을 지원하는 강의실 파서"""
    
    def __init__(self, file_path: str = None, data_directory: str = "data"):
        self.data_directory = data_directory
        self.converter = HTMLToExcelConverter(data_directory)
        
        # HTML 파일 자동 변환
        self._ensure_excel_files()
        
        # 파일 경로 설정
        if file_path:
            self.file_path = file_path
        else:
            self.file_path = self._find_converted_file()
        
        if self.file_path:
            self._load_excel_file()
    
    def _ensure_excel_files(self):
        """HTML 파일들을 Excel로 변환"""
        print("🔍 HTML 파일 변환 확인 중...")
        
        # 이미 변환된 파일들 확인
        converted_files = []
        if os.path.exists(self.data_directory):
            for file in os.listdir(self.data_directory):
                if file.endswith('_converted.xlsx'):
                    converted_files.append(file)
        
        if not converted_files:
            print("📥 HTML 파일을 Excel로 변환합니다...")
            self.converter.convert_all_files()
            self.converter.create_summary_report()
        else:
            print(f"✅ 이미 변환된 파일 {len(converted_files)}개를 발견했습니다.")
    
    def _find_converted_file(self) -> str:
        """변환된 Excel 파일 중 301호실 파일 찾기"""
        if not os.path.exists(self.data_directory):
            return None
        
        for file in os.listdir(self.data_directory):
            if file.endswith('_converted.xlsx') and '301' in file:
                file_path = os.path.join(self.data_directory, file)
                print(f"🎯 301호실 파일 발견: {file}")
                return file_path
        
        # 301호실이 없으면 첫 번째 변환된 파일 사용
        for file in os.listdir(self.data_directory):
            if file.endswith('_converted.xlsx'):
                file_path = os.path.join(self.data_directory, file)
                print(f"📋 첫 번째 변환 파일 사용: {file}")
                return file_path
        
        return None
    
    def _load_excel_file(self):
        """Excel 파일 로드"""
        try:
            print(f"📖 파일 로드 중: {os.path.basename(self.file_path)}")
            
            self.df = pd.read_excel(self.file_path)
            
            # 파일명에서 기본 정보 추출
            self._extract_basic_info()
            
            # 시간대 컬럼 추출
            self._extract_time_columns()
            
            print(f"✅ 파일 로드 완료 - {self.df.shape[0]}행 × {self.df.shape[1]}열")
            
        except Exception as e:
            print(f"❌ 파일 로드 실패: {e}")
            self.df = None
    
    def _extract_basic_info(self):
        """파일명에서 기본 정보 추출"""
        file_name = os.path.basename(self.file_path)
        
        # 건물명 추출
        building_match = re.search(r'^([가-힣]+관)', file_name)
        self.building = building_match.group(1) if building_match else None
        
        # 호실 번호 추출
        room_match = re.search(r'관(\d+)', file_name)
        self.room_number = room_match.group(1) if room_match else None
        
        # 수용인원 추출
        capacity_match = re.search(r'수용인원\s*(\d+)명', file_name)
        self.capacity = int(capacity_match.group(1)) if capacity_match else None
        
        # 층 수 계산
        if self.room_number:
            self.floor = int(self.room_number) // 100
        
        print(f"🏢 강의실 정보: {self.building} {self.room_number}호 ({self.floor}층, {self.capacity}명)")
    
    def _extract_time_columns(self):
        """시간대 컬럼 추출"""
        if self.df is None:
            return
        
        time_pattern = r'\d{2}:\d{2}~'
        self.time_columns = []
        
        for col in self.df.columns:
            if isinstance(col, str) and re.match(time_pattern, col):
                self.time_columns.append(col)
        
        self.time_columns.sort()
        print(f"⏰ 시간대: {len(self.time_columns)}개 ({self.time_columns[0] if self.time_columns else 'None'} ~ {self.time_columns[-1] if self.time_columns else 'None'})")
    
    def get_room_status_at_time(self, date: str, time: str) -> Dict:
        """특정 시간에 강의실의 상태 확인"""
        if self.df is None:
            return {'status': 'error', 'message': '데이터가 로드되지 않았습니다.'}
        
        # 날짜 형식 변환
        if isinstance(date, str):
            try:
                date = int(date)
            except ValueError:
                return {'status': 'error', 'message': '날짜 형식이 올바르지 않습니다.'}
        
        # 해당 날짜의 데이터 찾기
        date_data = self.df[self.df['사용일자'] == date]
        
        if date_data.empty:
            return {'status': 'no_data', 'message': f'{date} 날짜의 데이터를 찾을 수 없습니다.'}
        
        # 시간대 컬럼 찾기
        time_col = f"{time}~"
        if time_col not in self.time_columns:
            return {
                'status': 'invalid_time',
                'message': f'{time} 시간대를 찾을 수 없습니다.',
                'available_times': self.time_columns
            }
        
        # 해당 시간의 강의 정보 확인
        usage_info = date_data[time_col].iloc[0]
        
        if pd.isna(usage_info) or str(usage_info).strip() == '' or str(usage_info).strip() == ' ':
            status = '사용 가능'
            is_available = True
        else:
            status = f'사용중 ({usage_info})'
            is_available = False
        
        return {
            'status': 'success',
            'is_available': is_available,
            'usage_info': status,
            'date': date,
            'time': time,
            'day_of_week': date_data['요일'].iloc[0] if '요일' in date_data.columns else None
        }
    
    def print_room_info(self):
        """강의실 기본 정보 출력"""
        if self.df is None:
            print("❌ 데이터가 로드되지 않았습니다.")
            return
        
        print("\n=== 강의실 정보 ===")
        print(f"🏢 건물: {self.building}")
        print(f"🚪 호실: {self.room_number}")
        print(f"🏢 층: {self.floor}층")
        print(f"👥 수용인원: {self.capacity}명")
        print(f"⏰ 시간대: {len(self.time_columns)}개")
        print(f"📅 데이터 행 수: {len(self.df)}")


def main():
    print("=== HTML to Excel 자동 변환 및 강의실 파서 ===\n")
    
    try:
        # 필요한 라이브러리 확인
        import pandas as pd
        import openpyxl
    except ImportError as e:
        print("❌ 필요한 라이브러리가 설치되지 않았습니다.")
        print("다음 명령어로 설치하세요:")
        print("pip install pandas openpyxl lxml html5lib")
        return
    
    # 강의실 파서 초기화 (자동으로 HTML 변환 수행)
    parser = UniversalClassroomParserFixed()
    
    if parser.df is None:
        print("❌ 사용 가능한 데이터가 없습니다.")
        return
    
    # 기본 정보 출력
    parser.print_room_info()
    
    # 테스트
    print("\n=== 테스트 ===")
    
    # 첫 번째 날짜 찾기
    first_date = parser.df['사용일자'].iloc[0]
    print(f"📅 첫 번째 데이터 날짜: {first_date}")
    
    # 10:00 시간대 상태 확인
    status = parser.get_room_status_at_time(first_date, "10:00")
    print(f"🕙 10:00 상태: {status}")
    
    print(f"\n✅ HTML 파일들이 성공적으로 Excel로 변환되었습니다!")
    print(f"📁 변환된 파일들은 data 폴더에서 확인할 수 있습니다.")


if __name__ == "__main__":
    main()
