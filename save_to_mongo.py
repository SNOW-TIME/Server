# save_to_mongo.py
import os
import pandas as pd
from pymongo import MongoClient

# 1. MongoDB 연결
client = MongoClient("mongodb+srv://sm2005:MYxnbIdebLhvy2pp@cluster0.cxxkrzy.mongodb.net/")
db = client["snowtime"]       # DB 이름
collection = db["classrooms"] # 컬렉션 이름

# 2. data 폴더 경로
data_folder = os.path.join(os.path.dirname(__file__), "data")

# 3. 폴더 내 모든 xlsx 파일 가져오기
xlsx_files = [f for f in os.listdir(data_folder) if f.endswith(".xlsx")]

for xlsx_file in xlsx_files:
    file_path = os.path.join(data_folder, xlsx_file)
    
    # 4. pandas로 읽기
    df = pd.read_excel(file_path)
    
    # 5. 필요 시 컬럼 이름 표준화/선택 (엑셀마다 컬럼명 다를 수 있음)
    # 예: df = df[['건물', '층', '강의실', '날짜', '시작시간', '사용시간', '과목명', '교수명']]
    
    # 6. dict로 변환
    records = df.to_dict(orient='records')
    
    # 7. MongoDB 저장
    if records:  # 데이터가 있으면 insert
        collection.insert_many(records)
        print(f"[완료] {xlsx_file}: {len(records)}개 데이터 저장")

print("모든 xlsx 파일 MongoDB 저장 완료!")
