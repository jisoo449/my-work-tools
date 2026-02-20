db = {
    'user': 'root',
    'password': '517136',
    'host': 'localhost',
    'port': 3306,
    'database': 'msp_projects'
}

db_url = f"mysql+mysqlconnector://{db['user']}:{db['password']}@" \
         f"{db['host']}:{db['port']}/{db['database']}?charset=utf8"