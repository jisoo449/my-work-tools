import os
from dotenv import load_dotenv

# DB 접속 정보는 프로젝트 루트의 .env 에서 읽는다 (git 제외, 예시: .env.example)
load_dotenv(os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), ".env"))

db = {
    'user': os.getenv('DB_USER', ''),
    'password': os.getenv('DB_PASS', ''),
    'host': os.getenv('DB_HOST', 'localhost'),
    'port': int(os.getenv('DB_PORT', '3306')),
    'database': os.getenv('DB_NAME', '')
}

db_url = f"mysql+mysqlconnector://{db['user']}:{db['password']}@" \
         f"{db['host']}:{db['port']}/{db['database']}?charset=utf8"