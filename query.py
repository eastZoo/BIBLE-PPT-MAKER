import sqlite3

def insert(data):
  #테이블 생성 및 데이터 삽입
  conn = sqlite3.connect('history.db')

  # db 쿼리를 조작하기 위한 커서 객체 생성
  cur = conn.cursor()

  print(data)
  print(type(data))
  cur.execute('INSERT INTO history (history) VALUES (?)',  (data,) )
  conn.commit()

  # 마지막엔 무조건 close() 메소드로 db연결을 해제해야 한다.
  conn.close()
  return "insert SUCCESS"
  
def select():
  #데이터 조회
  conn = sqlite3.connect("history.db")

  cur = conn.cursor()
  cur.execute("SELECT * FROM history")

  historys = cur.fetchall()
  conn.close()
  return historys
  