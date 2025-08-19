import os

def check_file_format(file_path):
    """
    파일의 실제 형식을 확인하는 함수
    """
    print(f"파일 검사 중: {file_path}")
    
    if not os.path.exists(file_path):
        print("❌ 파일이 존재하지 않습니다.")
        return False
    
    # 파일 크기 확인
    file_size = os.path.getsize(file_path)
    print(f"📁 파일 크기: {file_size:,} bytes")
    
    # 파일의 첫 몇 바이트 읽어서 형식 확인
    try:
        with open(file_path, 'rb') as f:
            first_bytes = f.read(100)
            
        print(f"🔍 파일 시작 바이트: {first_bytes[:50]}")
        
        # HTML 파일인지 확인
        if first_bytes.startswith(b'<html') or first_bytes.startswith(b'<!DOCTYPE'):
            print("❌ 이 파일은 HTML 파일입니다. Excel 파일이 아닙니다.")
            
            # HTML 내용 일부 출력
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    html_content = f.read(500)
                print("📄 HTML 내용 미리보기:")
                print(html_content[:300])
            except:
                pass
            return False
        
        # Excel 파일 시그니처 확인
        # .xls 파일: 0xD0CF11E0 (OLE Document)
        # .xlsx 파일: 0x504B0304 (ZIP archive)
        
        if first_bytes.startswith(b'\xd0\xcf\x11\xe0'):
            print("✅ 올바른 .xls (Excel 97-2003) 파일입니다.")
            return True
        elif first_bytes.startswith(b'PK'):
            print("✅ .xlsx/.xlsm (Excel 2007+) 파일일 가능성이 높습니다.")
            return True
        else:
            print("❓ 알 수 없는 파일 형식입니다.")
            print(f"   첫 4바이트: {first_bytes[:4].hex()}")
            return False
            
    except Exception as e:
        print(f"❌ 파일 읽기 오류: {e}")
        return False

def find_excel_files_in_directory(directory="."):
    """
    디렉토리에서 Excel 파일들을 찾고 검사
    """
    print(f"\n📂 디렉토리 검사: {os.path.abspath(directory)}")
    
    excel_extensions = ['.xls', '.xlsx', '.xlsm']
    found_files = []
    
    try:
        for file in os.listdir(directory):
            file_path = os.path.join(directory, file)
            if os.path.isfile(file_path):
                file_ext = os.path.splitext(file)[1].lower()
                if file_ext in excel_extensions:
                    found_files.append(file_path)
                    print(f"\n📋 Excel 파일 발견: {file}")
                    check_file_format(file_path)
                    
    except Exception as e:
        print(f"❌ 디렉토리 읽기 오류: {e}")
    
    return found_files

def main():
    print("=== Excel 파일 형식 검사 도구 ===\n")
    
    # 현재 디렉토리에서 Excel 파일 찾기
    found_files = find_excel_files_in_directory(".")
    
    # data 디렉토리도 확인
    if os.path.exists("data"):
        found_files.extend(find_excel_files_in_directory("data"))
    
    # 특정 파일 직접 검사
    target_file = "data/프라임관301,수용인원 0070명,캡스톤디자인강의실(안유현강의실).XLS"
    if os.path.exists(target_file):
        print(f"\n🎯 대상 파일 직접 검사:")
        check_file_format(target_file)
    
    print("\n=== 해결 방법 제안 ===")
    print("1. 파일이 HTML인 경우:")
    print("   - 원본 Excel 파일을 다시 다운로드하세요")
    print("   - 웹브라우저에서 '다른 이름으로 저장' 대신 직접 다운로드 링크를 사용하세요")
    print("\n2. 파일이 손상된 경우:")
    print("   - Excel에서 파일을 열어보세요")
    print("   - 열린다면 '다른 이름으로 저장'으로 새 파일을 만드세요")
    print("\n3. 파일 형식 변환:")
    print("   - .xls를 .xlsx로 변환해보세요")

if __name__ == "__main__":
    main()
