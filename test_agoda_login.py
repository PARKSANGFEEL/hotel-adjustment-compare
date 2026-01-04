"""
아고다 로그인 페이지 구조 확인 (디버그용)
"""
import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options

options = Options()
options.add_argument('--no-sandbox')
options.add_argument('--disable-dev-shm-usage')

driver = webdriver.Chrome(
    service=Service(ChromeDriverManager().install()),
    options=options
)

try:
    # 로그인 페이지 방문
    print("아고다 로그인 페이지 방문 중...")
    driver.get('http://ycs.agoda.com/mldc/en-us/public/login')
    time.sleep(5)
    
    # 페이지 소스 저장
    with open('agoda_login_page.html', 'w', encoding='utf-8') as f:
        f.write(driver.page_source)
    
    print("페이지 소스 저장 완료: agoda_login_page.html")
    
    # 이메일 입력 필드 찾기 시도
    print("\n입력 필드 찾기:")
    
    # NAME으로 찾기
    inputs = driver.find_elements("tag name", "input")
    print(f"총 input 필드: {len(inputs)}")
    
    for i, inp in enumerate(inputs):
        name = inp.get_attribute('name')
        id_attr = inp.get_attribute('id')
        type_attr = inp.get_attribute('type')
        placeholder = inp.get_attribute('placeholder')
        print(f"  [{i}] name={name}, id={id_attr}, type={type_attr}, placeholder={placeholder}")
    
    # 버튼 찾기
    print("\n버튼 찾기:")
    buttons = driver.find_elements("tag name", "button")
    for i, btn in enumerate(buttons):
        text = btn.text
        id_attr = btn.get_attribute('id')
        name = btn.get_attribute('name')
        print(f"  [{i}] text={text}, id={id_attr}, name={name}")
    
    # form 찾기
    print("\nForm 찾기:")
    forms = driver.find_elements("tag name", "form")
    print(f"총 form: {len(forms)}")
    for i, form in enumerate(forms):
        id_attr = form.get_attribute('id')
        name = form.get_attribute('name')
        print(f"  [{i}] id={id_attr}, name={name}")
    
    print("\n" + "="*50)
    print("브라우저 창을 열어두고 있습니다.")
    print("HTML 구조를 확인하신 후 Enter를 누르세요.")
    print("="*50)
    input()
    
finally:
    driver.quit()
    print("\n드라이버 종료")
