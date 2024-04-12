import sqlite3

def initial():
  print("INITIAL")
  #테이블 생성 및 데이터 삽입
  history = sqlite3.connect('history.db')
  bible_index = sqlite3.connect('bible_index.db')
  # db 쿼리를 조작하기 위한 커서 객체 생성
  hisCur = history.cursor()
  bibleCur = bible_index.cursor()
  
  # ppt 생성 기록 테이블 생성
  hisCur.execute("""CREATE TABLE IF NOT EXISTS history(
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    history TEXT,
    createAt TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )""")
  # 성경 인덱스 테이블 생성
  bibleCur.execute("""CREATE TABLE IF NOT EXISTS bible_index(
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT,
    abbr TEXT
    )""")
  
  # 성경인덱스 데이터 있는지 확인
  index = bible_index_select()
  print(len(index))
  # 삽입할 데이터
  data = [
      ('창세기','창'),
      ('출애굽기','출'),
      ('레위기','레'),
      ('민수기','민'),
      ('신명기','신'),
      ('여호수아','수'),
      ('사사기','삿'),
      ('룻기','롯'),
      ('사무엘상','삼상'),
      ('사무엘하','삼하'),
      ('열왕기상','왕상'),
      ('열왕기하','왕하'),
      ('역대상','대상'),
      ('역대하','대하'),
      ('에스라','스'),
      ('느헤미야','느'),
      ('에스더','에'),
      ('욥기','욥'),
      ('시편','시'),
      ('잠언','잠'),
      ('전도서','전'),
      ('아가','아'),
      ('이사야','사'),
      ('예레미야','렘'),
      ('예레미야애가','애'),
      ('에스겔','겔'),
      ('다니엘','단'),
      ('호세아','호'),
      ('요엘','욜'),
      ('아모스','암'),
      ('오바댜','옵'),
      ('요나','욘'),
      ('미가','미'),
      ('나훔','나'),
      ('하박국','합'),
      ('스바냐','습'),
      ('학개','학'),
      ('스가랴','슥'),
      ('말라기','말'),
      ('마태복음','마'),
      ('마가복음','막'),
      ('누가복음','눅'),
      ('요한복음','요'),
      ('사도행전','행'),
      ('로마서','롬'),
      ('고린도전서','고전'),
      ('고린도후서','고후'),
      ('갈라디아서','갈'),
      ('에베소서','엡'),
      ('빌립보서','빌'),
      ('골로새서','골'),
      ('데살로니가전서','살전'),
      ('데살로니가후서','살후'),
      ('디모데전서','딤전'),
      ('디모데후서','딤후'),
      ('디도서','딛'),
      ('빌레몬서','몬'),
      ('히브리서','히'),
      ('야고보서','약'),
      ('베드로전서','벧전'),
      ('베드로후서','벧후'),
      ('요한1서','요1'),
      ('요한2서','요2'),
      ('요한3서','요3'),
      ('유다서','유'),
      ('요한계시록','계')
  ]
  
  # 위의 성경 인덱스 디비가 없을때만 추가 해주는 분기
  if(len(index) == 0):
    print("len == 0")
    # 여러 개의 레코드 삽입
    bibleCur.executemany('INSERT INTO bible_index (name, abbr) VALUES (?, ?)', data)


  # 변경사항 저장
  history.commit()
  history.close()
  # 변경사항 저장
  bible_index.commit()
  bible_index.close()
  
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
  cur.execute("SELECT * FROM history ORDER BY createAt DESC")

  historys = cur.fetchall()
  conn.close()
  return historys


def delete(id):
  print(id)
  print(type(id))
  #데이터 조회
  conn = sqlite3.connect("history.db")

  cur = conn.cursor()
  cur.execute("DELETE FROM history WHERE id = ?", (id,))
  
  conn.commit()
  conn.close()

def bible_index_select():
  #데이터 조회
  conn = sqlite3.connect("bible_index.db")

  cur = conn.cursor()
  cur.execute("SELECT * FROM bible_index")

  bible_index = cur.fetchall()
  conn.close()
  return bible_index