import pandas as pd  # 수정: pandas 사용으로 변경 (변환된 xlsx 파일 처리)
import re
import os
from datetime import datetime, timedelta
from typing import Dict, List, Tuple, Optional

class UniversalClassroomParser:
    def __init__(self, file_path: str):
        """
        강의실 스케줄 파서 초기화 (변환된 Excel 파일 지원)
        
        Args:
            file_path: 엑셀 파일 경로
        """
        self.file_path = file_path
        self.building = None
        self.room_number = None
        self.capacity = None
        self.floor = None
        self.time_columns = []
        self.df = None  # 수정: pandas DataFrame 사용
        
        self._load_and_parse()
    
    def _load_and_parse(self):
        """엑셀 파일 로드 및 기본 정보 파싱"""
        try:
            # 파일 확장자 확인
            file_ext = os.path.splitext(self.file_path)[1].lower()
            
            print(f"📖 파일 로드 중: {os.path.basename(self.file_path)}")
            
            # pandas로 Excel 파일 읽기
            if file_ext == '.xlsx':
                self.df = pd.read_excel(self.file_path, engine='openpyxl')
            elif file_ext == '.xls':
                self.df = pd.read_excel(self.file_path, engine='xlrd')
            else:
                # CSV나 다른 형식도 시도
                self.df = pd.read_excel(self.file_path)
            
            print(f"✅ 파일 로드 성공: {self.df.shape[0]}행 × {self.df.shape[1]}열")
            
            # 파일명에서 기본 정보 추출
            self._extract_basic_info()
            
            # 시간대 컬럼 추출
            self._extract_time_columns()
            
        except Exception as e:
            print(f"❌ 파일 로드 중 오류 발생: {e}")
            raise
    
    def _extract_basic_info(self):
        """파일명에서 건물, 호실, 수용인원 정보 추출"""
        file_name = os.path.basename(self.file_path)
        print(f"🔍 파일명 분석 중: {file_name}")
        
        # 건물명 추출 (예: 프라임관)
        building_match = re.search(r'^([가-힣]+관)', file_name)
        self.building = building_match.group(1) if building_match else None
        
        # 호실 번호 추출 (예: 301) - 쉼표 고려
        room_match = re.search(r'관(\d+)', file_name)
        self.room_number = room_match.group(1) if room_match else None
        
        # 수용인원 추출 (예: 0070명) - 쉼표와 공백 고려
        capacity_match = re.search(r'수용인원\s*(\d+)명', file_name)
        self.capacity = int(capacity_match.group(1)) if capacity_match else None
        
        # 층 수 계산
        if self.room_number:
            self.floor = int(self.room_number) // 100
        
        print(f"🏢 추출된 정보 - 건물: {self.building}, 호실: {self.room_number}, 수용인원: {self.capacity}, 층: {self.floor}")
    
    def _extract_time_columns(self):
        """시간대 컬럼 추출"""
        if self.df is None:
            return
        
        time_pattern = r'\d{2}:\d{2}~'
        self.time_columns = []
        
        for col in self.df.columns:
            if isinstance(col, str) and re.match(time_pattern, col):
                self.time_columns.append(col)
        
        self.time_columns.sort()  # 시간순 정렬
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
                return {
                    'status': 'error',
                    'message': '날짜 형식이 올바르지 않습니다. YYYYMMDD 형식을 사용하세요.'
                }
        
        # 해당 날짜의 데이터 찾기
        date_data = self.df[self.df['사용일자'] == date]
        
        if date_data.empty:
            return {
                'status': 'no_data',
                'message': f'{date} 날짜의 데이터를 찾을 수 없습니다.'
            }
        
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
        
        # pandas의 NaN 값도 고려
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
    
    def get_building_info(self) -> str:
        return self.building
    
    def get_room_info(self) -> str:
        return self.room_number
    
    def get_floor_info(self) -> int:
        return self.floor
    
    def get_capacity_info(self) -> int:
        return self.capacity
    
    def get_available_times_from(self, date: str, start_time: str) -> List[Tuple[str, str]]:
        """사용자가 입력한 시간으로부터 사용 가능한 시간대들 반환"""
        if self.df is None:
            return []
        
        if isinstance(date, str):
            try:
                date = int(date)
            except ValueError:
                return []
        
        # 해당 날짜의 데이터 찾기
        date_data = self.df[self.df['사용일자'] == date]
        
        if date_data.empty:
            return []
        
        available_times = []
        start_found = False
        
        for time_col in self.time_columns:
            current_time = time_col.replace('~', '')
            if current_time >= start_time:
                start_found = True
            
            if not start_found:
                continue
            
            usage_info = date_data[time_col].iloc[0]
            
            if pd.isna(usage_info) or str(usage_info).strip() == '' or str(usage_info).strip() == ' ':
                available_times.append((current_time, '사용 가능'))
            else:
                available_times.append((current_time, f'사용중 ({usage_info})'))
        
        return available_times
    
    def get_full_schedule_for_date(self, date: str) -> Dict:
        """특정 날짜의 전체 시간표 반환"""
        if self.df is None:
            return {'error': '데이터가 로드되지 않았습니다.'}
        
        if isinstance(date, str):
            try:
                date = int(date)
            except ValueError:
                return {'error': '날짜 형식이 올바르지 않습니다.'}
        
        # 해당 날짜의 데이터 찾기
        date_data = self.df[self.df['사용일자'] == date]
        
        if date_data.empty:
            return {'error': f'{date} 날짜의 데이터를 찾을 수 없습니다.'}
        
        schedule = {}
        for time_col in self.time_columns:
            usage_info = date_data[time_col].iloc[0]
            
            if pd.isna(usage_info) or str(usage_info).strip() == '' or str(usage_info).strip() == ' ':
                schedule[time_col] = '사용 가능'
            else:
                schedule[time_col] = f'사용중 ({usage_info})'
        
        return {
            'date': date,
            'day_of_week': date_data['요일'].iloc[0] if '요일' in date_data.columns else None,
            'building': self.building,
            'room': self.room_number,
            'floor': self.floor,
            'capacity': self.capacity,
            'schedule': schedule
        }
    
    def get_available_dates(self) -> List[int]:
        """사용 가능한 모든 날짜 반환"""
        if self.df is None:
            return []
        
        dates = self.df['사용일자'].dropna().unique().tolist()
        return sorted([int(date) for date in dates if not pd.isna(date)])
    
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
        
        available_dates = self.get_available_dates()
        if available_dates:
            print(f"📊 데이터 기간: {min(available_dates)} ~ {max(available_dates)}")

# classroom_parser.py의 main() 함수에서 파일 경로 찾는 부분을 이것으로 교체하세요

def main():
    import os
    
    try:
        print("=== 강의실 스케줄 파서 ===\n")
        
        # 수정: 변환된 파일을 최우선으로 찾기
        file_path = None
        
        # 1단계: 변환된 파일 직접 검색
        converted_paths = [
            "data/프라임관301,수용인원 0070명,캡스톤디자인강의실(안유현강의실)_converted.xlsx",
            "프라임관301,수용인원 0070명,캡스톤디자인강의실(안유현강의실)_converted.xlsx",
        ]
        
        for path in converted_paths:
            if os.path.exists(path):
                file_path = path
                print(f"🎯 변환된 Excel 파일 사용: {os.path.basename(path)}")
                break
        
        # 2단계: 변환된 파일이 없으면 자동 검색
        if not file_path:
            print("🔍 변환된 파일을 찾는 중...")
            search_dirs = ['data', '.']
            
            for search_dir in search_dirs:
                if os.path.exists(search_dir):
                    for file in os.listdir(search_dir):
                        if '_converted.xlsx' in file and '프라임관301' in file:
                            file_path = os.path.join(search_dir, file)
                            print(f"🎯 변환된 Excel 파일 발견: {file}")
                            break
                    if file_path:
                        break
        
        # 3단계: 그래도 없으면 안내 메시지
        if not file_path:
            print("❌ 변환된 Excel 파일을 찾을 수 없습니다.")
            print("\n💡 해결 방법:")
            print("1. 먼저 HTML 파일들을 변환하세요:")
            print("   python batch_html_converter.py")
            print("2. 또는 다음 명령으로 빠른 테스트:")
            print("   python quick_test.py")
            return
        
        print(f"✅ 사용할 파일: {file_path}")
        
        # 나머지 코드는 그대로 유지...
        parser = UniversalClassroomParser(file_path)
        parser.print_room_info()
        
        # 테스트 데이터 확인
        available_dates = parser.get_available_dates()
        if not available_dates:
            print("❌ 사용 가능한 날짜 데이터가 없습니다.")
            return
        
        test_date = available_dates[0]
        print(f"\n📅 테스트 날짜: {test_date}")
        
        status = parser.get_room_status_at_time(str(test_date), "10:00")
        print(f"🕙 10:00 상태: {status}")
        
        available_times = parser.get_available_times_from(str(test_date), "14:00")
        if available_times:
            print(f"\n⏰ {test_date} 14:00부터 사용 가능한 시간 (처음 10개):")
            for time, status in available_times[:10]:
                print(f"  {time}: {status}")
        
        print(f"\n✅ 강의실 파서 테스트 완료!")
        
    except ImportError as e:
        print(f"❌ 필요한 라이브러리가 설치되지 않았습니다: {e}")
        print("다음 명령어로 설치하세요:")
        print("pip install pandas openpyxl")
    except Exception as e:
        print(f"❌ 오류 발생: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()