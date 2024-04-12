# python 필요 라이브러리 설치

```bash
pip install -r requirements.txt
```

# python sqllite

https://itstory1592.tistory.com/37

# python slide 레이아웃 관리

https://ai-creator.tistory.com/208

# 첫번째 슬라이드 복사하는 코드 from chatGPT

python-pptx 라이브러리에서는 슬라이드를 복사하는 기능이 내장되어 있지 않습니다. 대신에 첫 번째 슬라이드를 생성한 후에 이를 복사하여 새로운 슬라이드를 만들어야 합니다. 이를 위해서는 첫 번째 슬라이드를 생성한 후에 해당 슬라이드의 내용을 복사하여 새로운 슬라이드에 붙여넣는 방식으로 작업할 수 있습니다.

다음은 이러한 접근 방식을 사용하여 첫 번째 슬라이드를 복사하여 새로운 슬라이드를 생성하는 코드입니다.

```python
from pptx import Presentation

# 기존 프레젠테이션 열기
prs = Presentation('existing_presentation.pptx')

# 첫 번째 슬라이드 가져오기
first_slide = prs.slides[0]

# 새로운 슬라이드 생성
slide_layout = first_slide.slide_layout
new_slide = prs.slides.add_slide(slide_layout)

# 첫 번째 슬라이드의 내용 복사하여 새로운 슬라이드에 붙여넣기
for shape in first_slide.shapes:
    new_shape = new_slide.shapes.add_shape(shape.auto_shape_type, shape.left, shape.top, shape.width, shape.height)
    if hasattr(shape, 'text'):
        new_shape.text = shape.text

# 프레젠테이션 저장
prs.save('output.pptx')
```

# 알게된것

1.  아래와 같은 함수가 있을때
    indicate 함수에 전달하는 page 함수가 아래와 같이 전달되면 함수 호출에 의해 즉시 실행되기 때문에

```python
    indicate("history",history_indicate, history.history_page(main_frame))
```

```python
  # 페이지 별 라우팅
    if(type == "history"):
        print("HISTORY")
        print(page)
        page
```

이 호출은 indicate() 함수를 호출하는 시점에서 한 번만 실행되며, 결과 값으로 반환된 페이지 객체가 page 변수에 할당됩니다. 하지만 history.history_page(main_frame)이 반환하는 값이 None이므로 page 변수에는 None이 할당됩니다.

해결 방법은 history.history_page(main_frame) 대신에 함수 참조만 전달하여 람다 함수가 필요할 때마다 호출되도록 하는 것입니다. 이를 위해 람다 함수 안에서 페이지를 생성하도록 변경해야 합니다.

```python
import pages.history as history

def indicate(type,lb, page):
    hide_indicators()
    lb.config(bg='#F15642')
    delete_page()

    # 페이지 별 라우팅
    if(type == "history"):
        print("HISTORY")
        print(page)
        page
    if(type == "home"):
        page()

history_btn = tk.Button(options_frame, text='history', font=('Bold', 15),
                        image=historyImg, height=50, width=50, command=lambda: indicate("history",history_indicate, lambda: history.history_page(main_frame)))
```
