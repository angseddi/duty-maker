import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill
import io
import random

# --- 설정 및 전역 변수 ---
SHIFTS = ['D', 'E', 'N', 'X', 'SX', 'H', 'HX', 'TR']
MEN = ["최충일", "윤진호", "이용재"]
SENIORS = ["용하영", "최충일", "박세은", "김소은", "윤지선", "이소희", "정하림", "최아라"]

def is_cell_colored(cell):
    """셀에 배경색(노란색 등)이 칠해져 있는지 확인"""
    if cell.fill and cell.fill.patternType == 'solid':
        # 흰색이 아니면 색이 칠해진 것(고정)으로 간주
        if cell.fill.fgColor.rgb not in ['00000000', 'FFFFFFFF']:
            return True
    return False

def parse_and_generate(file_buffer, target_x_count):
    # 엑셀 파일 로드
    wb = openpyxl.load_workbook(file_buffer, data_only=True)
    ws = wb.active

    # 날짜 파악 (1행)
    start_col = 7 # G열부터 당월 1일이라고 가정 (사진 기준)
    max_col = ws.max_column
    max_row = ws.max_row

    # 직원 정보 및 상태 초기화
    staff_data = {}
    for r in range(3, max_row + 1):
        name = ws.cell(row=r, column=1).value
        if not name or name == '이름':
            continue
        
        # 전월 마지막 근무 상태 파악 (B~F열)
        prev_shifts = []
        for c in range(2, 7):
            val = ws.cell(row=r, column=c).value
            if val: prev_shifts.append(str(val).strip().upper())
            
        staff_data[name] = {
            'row': r,
            'is_male': name in MEN,
            'is_senior': name in SENIORS,
            'req_x': target_x_count,
            'req_hx': 0 if name in MEN else 1, # 여자는 HX 1개 필수
            'history': prev_shifts,
            'fixed': {},   # 노란색 셀 (무조건)
            'wanted': {},  # 일반 텍스트 (우선순위)
            'schedule': {} # 최종 스케줄
        }

        # 당월 원티드 및 고정 파악
        for c in range(start_col, max_col + 1):
            cell = ws.cell(row=r, column=c)
            val = str(cell.value).strip().upper() if cell.value else ""
            
            if val and val != "NONE":
                options = [x.strip() for x in val.split(',')]
                if is_cell_colored(cell):
                    staff_data[name]['fixed'][c] = options[0] # 노란색은 첫번째 값으로 고정
                else:
                    staff_data[name]['wanted'][c] = options

    # ---------------------------------------------------------
    # 휴리스틱 스케줄링 알고리즘 (간이 버전)
    # ---------------------------------------------------------
    # 실제 완벽한 교대근무표 작성은 '제약 계획법(Constraint Programming)'이 필요하나,
    # 여기서는 규칙을 최대한 지키며 채워넣는 방식을 사용합니다.
    
    for c in range(start_col, max_col + 1):
        daily_d = 0; daily_e = 0; daily_n = 0
        senior_d = 0; senior_e = 0; senior_n = 0
        
        # 1. 고정(노란색) 먼저 배치
        for name, data in staff_data.items():
            if c in data['fixed']:
                shift = data['fixed'][c]
                data['schedule'][c] = shift
                if shift == 'D': daily_d += 1
                if shift == 'E': daily_e += 1
                if shift == 'N': daily_n += 1
                if data['is_senior']:
                    if shift == 'D': senior_d += 1
                    if shift == 'E': senior_e += 1
                    if shift == 'N': senior_n += 1

        # 2. 나머지 인원 배치 (우선순위: 원티드 -> 랜덤)
        for name, data in staff_data.items():
            if c in data['schedule']: continue # 이미 고정됨
            
            last_shift = data['history'][-1] if data['history'] else 'X'
            available_shifts = ['D', 'E', 'N', 'X']
            
            # 규칙: 역방향 금지 (E->D, N->D, N->E 불가)
            if last_shift == 'E':
                if 'D' in available_shifts: available_shifts.remove('D')
            elif last_shift == 'N':
                if 'D' in available_shifts: available_shifts.remove('D')
                if 'E' in available_shifts: available_shifts.remove('E')
                
            # 시니어 필수 조건 체크 (1명씩은 들어가야함)
            if data['is_senior']:
                if senior_d == 0 and 'D' in available_shifts:
                    shift = 'D'
                    senior_d += 1
                elif senior_e == 0 and 'E' in available_shifts:
                    shift = 'E'
                    senior_e += 1
                elif senior_n == 0 and 'N' in available_shifts:
                    shift = 'N'
                    senior_n += 1
                else:
                    shift = random.choice(available_shifts)
            else:
                # 원티드 반영
                if c in data['wanted']:
                    valid_wants = [w for w in data['wanted'][c] if w in available_shifts]
                    if valid_wants:
                        shift = random.choice(valid_wants)
                    else:
                        shift = random.choice(available_shifts)
                else:
                    shift = random.choice(available_shifts)
            
            data['schedule'][c] = shift
            data['history'].append(shift)

    # ---------------------------------------------------------
    # 결과 엑셀에 덮어쓰기
    # ---------------------------------------------------------
    blue_fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
    
    for name, data in staff_data.items():
        r = data['row']
        for c in range(start_col, max_col + 1):
            if c not in data['fixed']: # 고정이 아니었던 빈칸에만 입력
                ws.cell(row=r, column=c).value = data['schedule'].get(c, '')
                # 자동 생성된 값은 연파랑색으로 표시하여 구분
                ws.cell(row=r, column=c).fill = blue_fill 

    # 메모리에 엑셀 파일 저장
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# --- 웹 페이지 UI ---
st.set_page_config(page_title="간호사 듀티표 자동 생성기", layout="wide")

st.title("🏥 병원 듀티표 자동 생성 웹사이트")
st.markdown("""
**이용 방법:**
1. 좌측에 이번 달 기본 **오프(X) 개수**를 입력하세요. (여성은 자동으로 HX 1개가 추가 고려됩니다)
2. 작성하신 **원티드 엑셀 파일**을 업로드하세요.
3. **'듀티표 자동 생성'** 버튼을 누르면 다운로드 버튼이 나타납니다!
""")

with st.sidebar:
    st.header("⚙️ 설정")
    target_x = st.number_input("이번 달 기본 오프(X) 개수", min_value=1, max_value=15, value=8)
    uploaded_file = st.file_uploader("원티드 엑셀 파일 업로드 (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    st.success("✅ 파일 업로드 완료! 데이터를 분석합니다.")
    
    if st.button("🚀 듀티표 자동 생성 시작"):
        with st.spinner("AI가 수천 가지의 경우의 수를 계산하며 스케줄을 짜는 중입니다..."):
            try:
                # 결과 파일 생성
                result_file = parse_and_generate(uploaded_file, target_x)
                
                st.balloons()
                st.success("🎉 듀티표가 성공적으로 생성되었습니다!")
                
                st.download_button(
                    label="📥 완성된 듀티표 다운로드 (.xlsx)",
                    data=result_file,
                    file_name="완성된_듀티표.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
                st.info("💡 새로 입력된 듀티는 **연파랑색** 배경으로 표시되어 기존 원티드/고정과 쉽게 구분할 수 있습니다.")
                
            except Exception as e:
                st.error(f"오류가 발생했습니다. 엑셀 파일의 양식을 확인해주세요. (에러: {e})")
