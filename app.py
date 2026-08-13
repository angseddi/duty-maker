import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill
import io
import random
from ortools.sat.python import cp_model

# --- 설정 및 전역 변수 ---
SHIFTS = ['D', 'E', 'N', 'O'] # O는 Off(휴무)를 의미 (X, SX, H, HX, TR 통합)
SHIFT_IDX = {'D': 0, 'E': 1, 'N': 2, 'O': 3}
REV_SHIFT_IDX = {0: 'D', 1: 'E', 2: 'N', 3: 'X'}

MEN = ["최충일", "윤진호", "이용재"]
# 6번 조건: 리더와 보조리더 분리
LEADER = ["용하영", "최충일", "박세은", "김소은", "윤지선", "이소희", "정하림", "최아라"]
SUB_LEADER = ["김민지", "박우영", "오지은"]

def is_cell_colored(cell):
    """셀에 배경색(노란색 등)이 칠해져 있는지 확인"""
    if cell.fill and cell.fill.patternType == 'solid':
        if cell.fill.fgColor.rgb not in ['00000000', 'FFFFFFFF', None]:
            return True
    return False

def solve_schedule(staff_data, num_days, base_x_count, min_d, min_e, min_n):
    model = cp_model.CpModel()
    num_staff = len(staff_data)
    
    # 1. 변수 생성: shifts[직원인덱스, 날짜인덱스, 근무형태인덱스] = 1(배정됨) or 0(배정안됨)
    shifts = {}
    for e in range(num_staff):
        for d in range(num_days):
            for s in range(4): # 0:D, 1:E, 2:N, 3:O
                shifts[(e, d, s)] = model.NewBoolVar(f'shift_e{e}_d{d}_s{s}')

    # --- 제약 조건 (하드 제약) 적용 ---
    
    for e in range(num_staff):
        # 규칙 1: 하루에 무조건 1개의 근무(또는 휴무)만 배정
        for d in range(num_days):
            model.AddExactlyOne(shifts[(e, d, s)] for s in range(4))

        # 규칙 3: 역방향 교대 금지 (E->D, N->D, N->E 불가)
        for d in range(num_days - 1):
            model.AddImplication(shifts[(e, d, 1)], shifts[(e, d+1, 0)].Not())
            model.AddImplication(shifts[(e, d, 2)], shifts[(e, d+1, 0)].Not())
            model.AddImplication(shifts[(e, d, 2)], shifts[(e, d+1, 1)].Not())

        # 규칙 4: 연속 5일 근무 시 반드시 다음날은 오프(O) (최대 5일 연속 근무)
        for d in range(num_days - 5):
            model.Add(sum(shifts[(e, d+i, 3)] for i in range(6)) >= 1)

        # 규칙 7: 나이트(N)는 최대 3일까지만 연속 가능
        for d in range(num_days - 3):
            model.Add(sum(shifts[(e, d+i, 2)] for i in range(4)) <= 3)
            
        # 규칙 7-2: N 근무 후에는 최소 2일 이상 휴무
        for d in range(num_days - 2):
            model.Add(shifts[(e, d, 2)] + shifts[(e, d+1, 3)] + (1 - shifts[(e, d+2, 3)]) <= 2)

        # 규칙 8: 오프(X) 개수 맞추기 (여자는 HX때문에 +1)
        target_offs = base_x_count if staff_data[e]['is_male'] else base_x_count + 1
        model.Add(sum(shifts[(e, d, 3)] for d in range(num_days)) == target_offs)

        # 규칙 9: 퐁당퐁당 휴무 금지 (일 -> O -> 일 금지)
        for d in range(1, num_days - 1):
            model.Add((1 - shifts[(e, d-1, 3)]) + shifts[(e, d, 3)] + (1 - shifts[(e, d+1, 3)]) <= 2)

    # 규칙 5: 노란색 고정 근무는 무조건 반영
    for e, staff in enumerate(staff_data):
        for d, shift_str in staff['fixed'].items():
            if shift_str in SHIFT_IDX:
                model.Add(shifts[(e, d, SHIFT_IDX[shift_str])] == 1)
            elif shift_str in ['X', 'SX', 'H', 'HX', 'TR']:
                model.Add(shifts[(e, d, 3)] == 1)

    # 기본 근무자 수 보장
    for d in range(num_days):
        model.Add(sum(shifts[(e, d, 0)] for e in range(num_staff)) >= min_d)
        model.Add(sum(shifts[(e, d, 1)] for e in range(num_staff)) >= min_e)
        model.Add(sum(shifts[(e, d, 2)] for e in range(num_staff)) >= min_n)

    # --- 리더/보조리더 조건 (소프트 제약 및 목표 함수 설정) ---
    
    leader_indices = [i for i, s in enumerate(staff_data) if s['name'] in LEADER]
    sub_leader_indices = [i for i, s in enumerate(staff_data) if s['name'] in SUB_LEADER]
    
    objective_terms = []
    
    # 1. 원티드(일반 글씨) 반영 보상 (가중치 10)
    for e, staff in enumerate(staff_data):
        for d, wanted_shifts in staff['wanted'].items():
            for w_shift in wanted_shifts:
                if w_shift in SHIFT_IDX:
                    objective_terms.append(10 * shifts[(e, d, SHIFT_IDX[w_shift])])
                elif w_shift in ['X', 'SX', 'H', 'HX', 'TR']:
                    objective_terms.append(10 * shifts[(e, d, 3)])

    # 2. 리더/보조리더 배치 로직
    for d in range(num_days):
        for s in [0, 1, 2]: # D, E, N
            # 해당 근무에 리더가 1명 이상 있는지 나타내는 boolean 변수
            has_leader = model.NewBoolVar(f'has_leader_d{d}_s{s}')
            # 리더 합계가 1 이상이면 has_leader는 1
            model.Add(sum(shifts[(e, d, s)] for e in leader_indices) >= 1).OnlyEnforceIf(has_leader)
            model.Add(sum(shifts[(e, d, s)] for e in leader_indices) == 0).OnlyEnforceIf(has_leader.Not())
            
            # 해당 근무에 보조리더가 1명 이상 있는지 나타내는 boolean 변수
            has_sub_leader = model.NewBoolVar(f'has_sub_leader_d{d}_s{s}')
            model.Add(sum(shifts[(e, d, s)] for e in sub_leader_indices) >= 1).OnlyEnforceIf(has_sub_leader)
            model.Add(sum(shifts[(e, d, s)] for e in sub_leader_indices) == 0).OnlyEnforceIf(has_sub_leader.Not())
            
            # 필수 조건: 리더나 보조리더 중 최소 1명은 있어야 함 (하드 제약)
            model.AddBoolOr([has_leader, has_sub_leader])
            
            # 목표 함수: 리더가 배치되는 것을 강력히 선호 (가중치 100)
            # 보조리더 배치는 허용하되 보상은 주지 않음 (또는 음수 페널티를 줄 수 있음)
            objective_terms.append(100 * has_leader)
            
    # 모델의 목표: 보상(objective_terms)을 최대화
    model.Maximize(sum(objective_terms))

    # 솔버 실행
    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = 60.0 # 복잡한 조건이므로 60초로 설정
    status = solver.Solve(model)

    if status == cp_model.OPTIMAL or status == cp_model.FEASIBLE:
        # 결과 저장
        for e, staff in enumerate(staff_data):
            for d in range(num_days):
                for s in range(4):
                    if solver.Value(shifts[(e, d, s)]) == 1:
                        staff['final_schedule'][d] = REV_SHIFT_IDX[s]
        return True
    else:
        return False


def process_excel(file_content, target_x_count, min_d, min_e, min_n):
    wb = openpyxl.load_workbook(io.BytesIO(file_content), data_only=True)
    ws = wb.active

    # 날짜 컬럼 파악
    start_col = None
    num_days = 0
    for col in range(1, ws.max_column + 1):
        val = ws.cell(row=1, column=col).value
        if val and str(val).isdigit(): # 1, 2, 3... 등 날짜가 시작되는 열
            if start_col is None:
                start_col = col
            num_days += 1
            
    if start_col is None:
        start_col = 2 # 기본값
        num_days = 31

    staff_data = []
    # 3행부터 마지막 행까지 읽기
    for r in range(3, ws.max_row + 1):
        name = ws.cell(row=r, column=1).value
        if not name or name == '이름':
            continue
            
        staff_dict = {
            'row': r,
            'name': str(name).strip(),
            'is_male': str(name).strip() in MEN,
            'fixed': {},
            'wanted': {},
            'final_schedule': {}
        }
        
        # 당월 원티드 및 고정 파악
        for d in range(num_days):
            cell = ws.cell(row=r, column=start_col + d)
            val = str(cell.value).strip().upper() if cell.value else ""
            
            if val and val != "NONE":
                options = [x.strip() for x in val.split(',')]
                if is_cell_colored(cell):
                    staff_dict['fixed'][d] = options[0]
                else:
                    staff_dict['wanted'][d] = options
                    
        staff_data.append(staff_dict)

    # AI 알고리즘 실행
    success = solve_schedule(staff_data, num_days, target_x_count, min_d, min_e, min_n)

    if not success:
        return None

    # 성공 시 엑셀에 데이터 쓰기
    blue_fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
    
    for staff in staff_data:
        r = staff['row']
        for d in range(num_days):
            col = start_col + d
            if d not in staff['fixed']: # 노란색 고정이 아니었던 곳만 새로 채움
                val = staff['final_schedule'].get(d, '')
                # 여자의 경우 추가된 오프 1개를 'HX'로 변환 (간단한 후처리)
                if val == 'X' and not staff['is_male'] and 'HX' not in staff['final_schedule'].values():
                    val = 'HX'
                
                ws.cell(row=r, column=col).value = val
                ws.cell(row=r, column=col).fill = blue_fill 

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# --- Streamlit UI ---
st.set_page_config(page_title="스마트 듀티표 생성기", layout="wide")

st.title("🏥 스마트 듀티표 자동 생성기")
st.markdown("입력하신 **제약조건**과 **리더/보조리더** 규칙을 모두 반영한 최적화 모델입니다.")

with st.sidebar:
    st.header("⚙️ 기본 설정")
    target_x = st.number_input("이번 달 기본 오프(X) 개수", min_value=1, max_value=15, value=8)
    
    st.header("👥 일별 최소 필요 인원")
    st.markdown("*(조건이 빡빡해서 에러가 날 경우 이 숫자를 조절해보세요)*")
    min_d = st.number_input("Day 최소 인원", min_value=1, max_value=10, value=3)
    min_e = st.number_input("Evening 최소 인원", min_value=1, max_value=10, value=3)
    min_n = st.number_input("Night 최소 인원", min_value=1, max_value=10, value=3)
    
    st.header("📂 파일 업로드")
    uploaded_file = st.file_uploader("원티드 엑셀 파일 업로드 (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    st.success("✅ 파일 업로드 완료! 데이터를 분석합니다.")
    
    if st.button("🚀 듀티표 AI 생성 시작"):
        with st.spinner("AI가 수천만 가지의 경우의 수를 계산하여 최적의 스케줄을 찾고 있습니다. (최대 60초 소요)"):
            try:
                # 결과 파일 생성
                result_file = process_excel(uploaded_file.getvalue(), target_x, min_d, min_e, min_n)
                
                if result_file:
                    st.balloons()
                    st.success("🎉 모든 조건이 만족되는 완벽한 듀티표가 생성되었습니다!")
                    
                    st.download_button(
                        label="📥 완성된 듀티표 다운로드 (.xlsx)",
                        data=result_file,
                        file_name="완성된_듀티표.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    st.info("💡 새로 입력된 듀티는 **연파랑색** 배경으로 표시됩니다. 원티드 요청 및 리더 배정은 조건이 허락하는 한 최대한 반영되었습니다.")
                else:
                    st.error("🚨 제약 조건이 너무 빡빡하여 해답을 찾을 수 없습니다! 노란색(고정) 휴무가 너무 많거나 조건이 충돌하는지 확인해주세요. (사이드바의 '일별 최소 필요 인원'을 줄여보세요)")
            except Exception as e:
                st.error(f"오류가 발생했습니다. 엑셀 파일의 양식을 확인해주세요. (에러: {e})")
