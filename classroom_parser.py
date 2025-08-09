import pandas as pd
import re
import os
from typing import Dict, List, Tuple, Optional

class ClassroomParser:
    """강의실 시간표 분석 파서"""
    
    def __init__(self, file_path: str):
        self.file_path = file_path
        self.building = None
        self.room_number = None
        self.capacity = None
        self.floor = None
        self.time_columns = []
        self.df = None
        
        self._load_and_parse()
    
    def _load_and_parse(self):
        """파일 로드 및 기본 정보 파싱"""
        try:
            file_ext = os.path.splitext(self.file_path)[1].lower()
            
            if file_ext == '.xlsx':
                self.df = pd.read_excel(self.file_path, engine='openpyxl')
            elif file_ext == '.xls':
                self.df = pd.read_excel(self.file_path, engine='xlrd')
            else:
                self.df = pd.read_excel(self.file_path)
            
            self._extract_basic_info()
            self._extract_time_columns()
            
        except Exception:
            self.df = None
    
    def _extract_basic_info(self):
        """파일명에서 기본 정보 추출"""
        file_name = os.path.basename(self.file_path)
        
        building_match = re.search(r'^([가-힣]+관)', file_name)
        self.building = building_match.group(1) if building_match else None
        
        room_match = re.search(r'관(\d+)', file_name)
        self.room_number = room_match.group(1) if room_match else None
        
        capacity_match = re.search(r'수용인원\s*(\d+)명', file_name)
        self.capacity = int(capacity_match.group(1)) if capacity_match else None
        
        if self.room_number:
            self.floor = int(self.room_number) // 100
    
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
    
    def get_room_status_at_time(self, date: str, time: str) -> Dict:
        """특정 시간에 강의실 상태 확인"""
        if self.df is None:
            return {'status': 'error', 'message': '데이터 로드 실패'}
        
        if isinstance(date, str):
            try:
                date = int(date)
            except ValueError:
                return {'status': 'error', 'message': '날짜 형식 오류'}
        
        date_data = self.df[self.df['사용일자'] == date]
        
        if date_data.empty:
            return {'status': 'no_data', 'message': f'{date} 날짜 데이터 없음'}
        
        time_col = f"{time}~"
        if time_col not in self.time_columns:
            return {'status': 'invalid_time', 'message': f'{time} 시간대 없음'}
        
        usage_info = date_data[time_col].iloc[0]
        
        if pd.isna(usage_info) or str(usage_info).strip() in ['', ' ']:
            is_available = True
            status = '사용 가능'
        else:
            is_available = False
            status = f'사용중 ({usage_info})'
        
        return {
            'status': 'success',
            'is_available': is_available,
            'usage_info': status,
            'date': date,
            'time': time,
            'day_of_week': date_data['요일'].iloc[0] if '요일' in date_data.columns else None
        }
    
    def get_available_times_from(self, date: str, start_time: str) -> List[Tuple[str, str]]:
        """지정 시간부터 사용 가능한 시간대 반환"""
        if self.df is None:
            return []
        
        if isinstance(date, str):
            try:
                date = int(date)
            except ValueError:
                return []
        
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
            
            if pd.isna(usage_info) or str(usage_info).strip() in ['', ' ']:
                available_times.append((current_time, '사용 가능'))
            else:
                available_times.append((current_time, f'사용중 ({usage_info})'))
        
        return available_times
    
    def get_full_schedule_for_date(self, date: str) -> Dict:
        """특정 날짜 전체 시간표 반환"""
        if self.df is None:
            return {'error': '데이터 로드 실패'}
        
        if isinstance(date, str):
            try:
                date = int(date)
            except ValueError:
                return {'error': '날짜 형식 오류'}
        
        date_data = self.df[self.df['사용일자'] == date]
        if date_data.empty:
            return {'error': f'{date} 날짜 데이터 없음'}
        
        schedule = {}
        for time_col in self.time_columns:
            usage_info = date_data[time_col].iloc[0]
            
            if pd.isna(usage_info) or str(usage_info).strip() in ['', ' ']:
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
    
    def get_room_info(self) -> Dict:
        """강의실 기본 정보 반환"""
        return {
            'building': self.building,
            'room_number': self.room_number,
            'floor': self.floor,
            'capacity': self.capacity,
            'time_slots': len(self.time_columns),
            'data_rows': len(self.df) if self.df is not None else 0
        }
    
    def get_available_dates(self) -> List[int]:
        """사용 가능한 날짜 목록 반환"""
        if self.df is None:
            return []
        
        dates = self.df['사용일자'].dropna().unique().tolist()
        return sorted([int(date) for date in dates if not pd.isna(date)])


class ClassroomSearcher:
    """사용자 입력 기반 강의실 검색"""
    
    def __init__(self, data_directory: str = "data"):
        self.data_directory = data_directory
        self.available_files = self._scan_available_files()
    
    def _scan_available_files(self) -> List[Dict]:
        """사용 가능한 강의실 파일 스캔"""
        files = []
        
        if not os.path.exists(self.data_directory):
            return files
        
        for file in os.listdir(self.data_directory):
            if file.endswith('_converted.xlsx'):
                file_path = os.path.join(self.data_directory, file)
                
                # 파일명에서 정보 추출
                building_match = re.search(r'^([가-힣]+관)', file)
                room_match = re.search(r'관(\d+)', file)
                capacity_match = re.search(r'수용인원\s*(\d+)명', file)
                
                if building_match and room_match:
                    building = building_match.group(1)
                    room_number = room_match.group(1)
                    capacity = int(capacity_match.group(1)) if capacity_match else 0
                    floor = int(room_number) // 100
                    
                    files.append({
                        'file_path': file_path,
                        'building': building,
                        'room_number': room_number,
                        'floor': floor,
                        'capacity': capacity,
                        'filename': file
                    })
        
        return files
    
    def find_rooms_by_criteria(self, building: str = None, floor: int = None, 
                              min_capacity: int = None) -> List[Dict]:
        """조건에 맞는 강의실 찾기"""
        filtered_files = self.available_files.copy()
        
        if building:
            filtered_files = [f for f in filtered_files if f['building'] == building]
        
        if floor is not None:
            filtered_files = [f for f in filtered_files if f['floor'] == floor]
        
        if min_capacity is not None:
            filtered_files = [f for f in filtered_files if f['capacity'] >= min_capacity]
        
        return filtered_files
    
    def get_available_buildings(self) -> List[str]:
        """사용 가능한 건물 목록"""
        buildings = set(f['building'] for f in self.available_files)
        return sorted(list(buildings))
    
    def get_available_floors_in_building(self, building: str) -> List[int]:
        """특정 건물의 사용 가능한 층 목록"""
        floors = set(f['floor'] for f in self.available_files if f['building'] == building)
        return sorted(list(floors))
    
    def search_available_rooms(self, building: str, floor: int, date: str, 
                              start_time: str, duration_hours: int = 1) -> List[Dict]:
        """조건에 맞는 사용 가능한 강의실 검색"""
        # 해당 건물, 층의 강의실들 찾기
        candidate_rooms = self.find_rooms_by_criteria(building=building, floor=floor)
        
        available_rooms = []
        
        for room_info in candidate_rooms:
            try:
                parser = ClassroomParser(room_info['file_path'])
                
                # 시작 시간부터 연속된 시간 확인
                available_times = parser.get_available_times_from(date, start_time)
                
                if available_times:
                    # 연속된 사용 가능 시간 계산
                    consecutive_available = 0
                    for time, status in available_times:
                        if '사용 가능' in status:
                            consecutive_available += 1
                        else:
                            break
                    
                    # 요청한 시간만큼 사용 가능한지 확인
                    if consecutive_available >= duration_hours * 2:  # 30분 단위
                        room_info['consecutive_hours'] = consecutive_available / 2
                        room_info['parser'] = parser
                        available_rooms.append(room_info)
                        
            except Exception:
                continue
        
        return available_rooms


def search_classrooms(building: str, floor: int, date: str, start_time: str, 
                     duration_hours: int = 1, data_directory: str = "data") -> Dict:
    """강의실 검색 메인 함수"""
    searcher = ClassroomSearcher(data_directory)
    
    # 입력 검증
    available_buildings = searcher.get_available_buildings()
    if building not in available_buildings:
        return {
            'status': 'error',
            'message': f'건물 "{building}"을 찾을 수 없습니다.',
            'available_buildings': available_buildings
        }
    
    available_floors = searcher.get_available_floors_in_building(building)
    if floor not in available_floors:
        return {
            'status': 'error',
            'message': f'{building} {floor}층을 찾을 수 없습니다.',
            'available_floors': available_floors
        }
    
    # 강의실 검색
    available_rooms = searcher.search_available_rooms(
        building, floor, date, start_time, duration_hours
    )
    
    if not available_rooms:
        return {
            'status': 'no_rooms',
            'message': f'{building} {floor}층에 {start_time}부터 {duration_hours}시간 사용 가능한 강의실이 없습니다.'
        }
    
    # 결과 정리
    results = []
    for room in available_rooms:
        room_info = room['parser'].get_room_info()
        results.append({
            'building': room['building'],
            'room_number': room['room_number'],
            'capacity': room['capacity'],
            'consecutive_hours': room['consecutive_hours'],
            'room_info': room_info
        })
    
    return {
        'status': 'success',
        'search_criteria': {
            'building': building,
            'floor': floor,
            'date': date,
            'start_time': start_time,
            'duration_hours': duration_hours
        },
        'found_rooms': len(results),
        'rooms': results
    }


def get_user_input():
    """사용자로부터 검색 조건 입력받기"""
    print("=== 강의실 검색 시스템 ===\n")
    
    # 사용 가능한 건물 목록 표시
    searcher = ClassroomSearcher()
    available_buildings = searcher.get_available_buildings()
    
    if not available_buildings:
        print("❌ 사용 가능한 강의실 데이터가 없습니다.")
        print("먼저 batch_html_converter.py를 실행하여 데이터를 변환하세요.")
        return None
    
    print(f"📋 사용 가능한 건물: {', '.join(available_buildings)}")
    
    # 건물 입력
    while True:
        building = input("\n🏢 건물명을 입력하세요: ").strip()
        if building in available_buildings:
            break
        print(f"❌ '{building}'을 찾을 수 없습니다. 다음 중에서 선택하세요: {', '.join(available_buildings)}")
    
    # 해당 건물의 사용 가능한 층 표시
    available_floors = searcher.get_available_floors_in_building(building)
    print(f"📋 {building}의 사용 가능한 층: {', '.join(map(str, available_floors))}")
    
    # 층 입력
    while True:
        try:
            floor = int(input(f"\n🏢 층수를 입력하세요: ").strip())
            if floor in available_floors:
                break
            print(f"❌ {floor}층을 찾을 수 없습니다. 다음 중에서 선택하세요: {', '.join(map(str, available_floors))}")
        except ValueError:
            print("❌ 숫자를 입력해주세요.")
    
    # 날짜 입력
    while True:
        date = input("\n📅 날짜를 입력하세요 (YYYYMMDD, 예: 20250901): ").strip()
        if len(date) == 8 and date.isdigit():
            break
        print("❌ 날짜 형식이 올바르지 않습니다. YYYYMMDD 형식으로 입력하세요.")
    
    # 시간 입력
    while True:
        start_time = input("\n🕐 시작 시간을 입력하세요 (HH:MM, 예: 14:00): ").strip()
        time_pattern = r'^\d{2}:\d{2}$'
        if re.match(time_pattern, start_time):
            break
        print("❌ 시간 형식이 올바르지 않습니다. HH:MM 형식으로 입력하세요.")
    
    # 사용 시간 입력
    while True:
        try:
            duration = int(input("\n⏱️ 사용할 시간(시간 단위, 예: 2): ").strip())
            if duration > 0:
                break
            print("❌ 1 이상의 숫자를 입력해주세요.")
        except ValueError:
            print("❌ 숫자를 입력해주세요.")
    
    return {
        'building': building,
        'floor': floor,
        'date': date,
        'start_time': start_time,
        'duration_hours': duration
    }


def display_search_results(result):
    """검색 결과 출력"""
    print("\n" + "="*50)
    
    if result['status'] == 'success':
        criteria = result['search_criteria']
        print(f"🔍 검색 조건:")
        print(f"   건물: {criteria['building']}")
        print(f"   층: {criteria['floor']}층")
        print(f"   날짜: {criteria['date']}")
        print(f"   시작 시간: {criteria['start_time']}")
        print(f"   사용 시간: {criteria['duration_hours']}시간")
        
        print(f"\n✅ 검색 결과: {result['found_rooms']}개 강의실 발견")
        
        if result['found_rooms'] > 0:
            print("\n📍 사용 가능한 강의실:")
            for i, room in enumerate(result['rooms'], 1):
                print(f"   {i}. {room['building']} {room['room_number']}호")
                print(f"      - 수용인원: {room['capacity']}명")
                print(f"      - 연속 사용 가능: {room['consecutive_hours']}시간")
                print()
        
    elif result['status'] == 'error':
        print(f"❌ 오류: {result['message']}")
        if 'available_buildings' in result:
            print(f"   사용 가능한 건물: {', '.join(result['available_buildings'])}")
        if 'available_floors' in result:
            print(f"   사용 가능한 층: {', '.join(map(str, result['available_floors']))}")
    
    elif result['status'] == 'no_rooms':
        print(f"😔 {result['message']}")
        print("\n💡 다른 조건으로 다시 검색해보세요:")
        print("   - 다른 시간대")
        print("   - 다른 층")
        print("   - 더 짧은 사용 시간")


if __name__ == "__main__":
    try:
        # 사용자 입력 받기
        search_params = get_user_input()
        
        if search_params is None:
            exit(1)
        
        print("\n🔍 강의실을 검색하는 중...")
        
        # 강의실 검색
        result = search_classrooms(**search_params)
        
        # 결과 출력
        display_search_results(result)
        
        # 계속 검색할지 묻기
        while True:
            continue_search = input("\n🔄 다시 검색하시겠습니까? (y/n): ").strip().lower()
            if continue_search in ['y', 'yes', 'ㅇ']:
                print("\n" + "="*50)
                search_params = get_user_input()
                if search_params:
                    print("\n🔍 강의실을 검색하는 중...")
                    result = search_classrooms(**search_params)
                    display_search_results(result)
                else:
                    break
            elif continue_search in ['n', 'no', 'ㄴ']:
                print("\n👋 강의실 검색을 종료합니다.")
                break
            else:
                print("❌ 'y' 또는 'n'을 입력해주세요.")
                
    except KeyboardInterrupt:
        print("\n\n👋 사용자가 프로그램을 종료했습니다.")
    except Exception as e:
        print(f"\n❌ 예상하지 못한 오류가 발생했습니다: {e}")
        print("프로그램을 다시 실행해주세요.")