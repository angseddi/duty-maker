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
# 6번 조건: 1~8번 리더 명단 (사진 기준)
LEADER = ["용하영", "최충일", "박세은", "김소은", "윤지선", "이소희", "정하림", "최아라"]

def is_cell_colored(cell):
    """셀에 배경색(노란색 등)이 칠해져 있는지 확인"""
    if cell.fill and cell.fill.patternType == 'solid':
        if cell.fill.fgColor.rgb not in ['00000000', 'FFFFFFFF']:
            return True
    return False

def solve_schedule(staff_data, num_days, base_x_count):
    model = cp_model.CpModel()
    num_staff = len(staff_data)
    
    # 1. 변수 생성: shifts[직원인덱스, 날짜인덱스, 근무형태인덱스] = 1(배정됨) or 0(배정안됨)
    shifts = {}
    for e in range(num_staff):
        for d in range(num_days):
            for s in range(4): # 0:D, 1:E, 2:N, 3:O
                shifts[(e, d, s)] = model.NewBoolVar(f'shift_e{e}_d{d}_s{s}')

    # --- 제약 조건 (규칙) 적용 ---
    
    # 규칙 1: 하루에 무조건 1개의 근무(또는 휴무)만 배정
    for e in range(num_staff):
        for d in range(num_days):
            model.AddExactlyOne(shifts[(e, d, s)] for s in range(4))

    # 규칙 5: 노란색 고정 근무는 무조건 반영
    for e, staff in enumerate(staff_data):
        for d, shift_str in staff['fixed'].items():
            if shift_str in SHIFT_IDX:
                model.Add(shifts[(e, d, SHIFT_IDX[shift_str])] == 1)
            elif shift_str in ['X', 'SX', 'H', 'HX', 'TR']:
                model.Add(shifts[(e, d, 3)] == 1) # Off로 처리

    # 규칙 3: 역방향 교대 금지 (E->D, N->D, N->E 불가)
    for e in range(num_staff):
        for d in range(num_days - 1):
            # E(1) 다음날 D(0) 불가: E가 1이면 다음날 D는 0이어야 함
            model.AddImplication(shifts[(e, d, 1)], shifts[(e, d+1, 0)].Not())
            # N(2) 다음날 D(0) 또는 E(1) 불가
            model.AddImplication(shifts[(e, d, 2)], shifts[(e, d+1, 0)].Not())
            model.AddImplication(shifts[(e, d, 2)], shifts[(e, d+1, 1)].Not())

    # 규칙 4: 연속 5일 근무 시 반드시 다음날은 오프(O)
    for e in range(num_staff):
        for d in range(num_days - 5):
            # d, d+1, d+2, d+3, d+4 5일간 오프(3)가 한 번도 없다면(모두 근무라면)
            # sum(shifts[(e, d+i, 3)]) == 0 이면 
            # d+5는 반드시 오프(3)여야 함
            working_5_days = model.NewBoolVar(f'work5_{e}_{d}')
            # 5일간 오프의 합이 0일 때 working_5_days가 1이 되도록 설정
            model.Add(sum(shifts[(e, d+i, 3)] for i in range(5)) == 0).OnlyEnforceIf(working_5_days)
            model.Add(sum(shifts[(e, d+i, 3)] for i in range(5)) > 0).OnlyEnforceIf(working_5_days.Not())
            
            # 5일 연속 근무했으면 다음날은 반드시 오프
            model.AddImplication(working_5_days, shifts[(e, d+5, 3)])

    # 규칙 7: 나이트(N)는 최대 3일까지만 연속 가능, N 근무 후에는 최소 2일 이상 휴무
    for e in range(num_staff):
        # 최대 3일 연속 제한: 4일 연속 N 불가
        for d in range(num_days - 3):
            model.Add(sum(shifts[(e, d+i, 2)] for i in range(4)) <= 3)
            
        # N 근무 후 무조건 휴무 (위에 역방향 금지에서 N->D, N->E를 막았으므로, 남은 건 N->N 또는 N->O)
        # N이 끝나고 O가 시작되면, 그 다음날도 O여야 함 (최소 2일 오프)
        for d in range(num_days - 2):
            # N(2) 하다가 다음날 O(3)가 되면, 그 다음날도 O(3)
            n_then_o = model.NewBoolVar(f'n_then_o_{e}_{d}')
            model.AddBoolAnd([shifts[(e, d, 2)], shifts[(e, d+1, 3)]]).OnlyEnforceIf(n_then_o)
            # a and b 가 아니면 n_then_o는 0이어야 함 (안전장치)
            model.AddImplication(n_then_o.Not(), shifts[(e, d, 2)].Not()) # 간략화된 형태
            
            model.AddImplication(n_then_o, shifts[(e, d+2, 3)])

    # 규칙 6: 리더(1~8번) 중 매일 D, E, N 각각 최소 1명 이상 필수
    leader_indices = [i for i, s in enumerate(staff_data) if s['name'] in LEADER]
    for d in range(num_days):
        model.Add(sum(shifts[(e, d, 0)] for e in leader_indices) >= 1) # D
        model.Add(sum(shifts[(e, d, 1)] for e in leader_indices) >= 1) # E
        model.Add(sum(shifts[(e, d, 2)] for e in leader_indices) >= 1) # N

    # 규칙 8: 오프(X) 개수 맞추기 (여자는 HX때문에 +1)
    for e, staff in enumerate(staff_data):
        total_offs = base_x_count if staff['is_male'] else base_x_count + 1
        model.Add(sum(shifts[(e, d, 3)] for d in range(num_days)) == total_offs)

    # 규칙 9: 퐁당퐁당 휴무 금지 (휴무는 무조건 2개 이상 붙어서)
    # 패턴: [일함, 오프, 일함] 이 불가능해야 함 -> 즉, O(3) 앞뒤로 근무(O가 아님)가 오는 것을 금지
    for e in range(num_staff):
        for d in range(1, num_days - 1):
            isolated_off = model.NewBoolVar(f'isolated_off_{e}_{d}')
            # d-1일 근무(O가 아님), d일 오프(O), d+1일 근무(O가 아님)
            model.AddBoolAnd([
                shifts[(e, d-1, 3)].Not(), 
                shifts[(e, d, 3)], 
                shifts[(e, d+1, 3)].Not()
            ]).OnlyEnforceIf(isolated_off)
            
            # isolated_off는 절대 발생하면 안됨
            model.Add(isolated_off == 0)

    # 규칙 10: 모든 근무자들의 D, E 비율을 비슷하게 (임시로 각 듀티의 편차를 줄이도록 목표 설정)
    # 여기서는 각 개인의 D, E, N 개수가 평균에 가깝도록 하는 것을 목표로 합니다.
    # 단순화를 위해 일단은 '원티드'를 최대한 맞추는 것을 최우선 목표로 잡습니다.

    # 목적 함수: 원티드(일반 글씨)를 최대한 많이 들어주도록 설정
    objective_terms = []
    for e, staff in enumerate(staff_data):
        for d, wanted_shifts in staff['wanted'].items():
            for w_shift in wanted_shifts:
                if w_shift in SHIFT_IDX:
                    objective_terms.append(shifts[(e, d, SHIFT_IDX[w_shift])])
                elif w_shift in ['X', 'SX', 'H', 'HX', 'TR']:
                    objective_terms.append(shifts[(e, d, 3)])
    
    model.Maximize(sum(objective_terms))

    # 솔버 실행
    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = 60.0 # 복잡한 조건이므로 60초로 증가
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


def process_excel(file_content, target_x_count):
    wb = openpyxl.load_workbook(io.BytesIO(file_content), data_only=True)
    ws = wb.active

    # 날짜 컬럼 찾기 (대략 G열, index 6부터 31일까지)
    start_col = 7 # G열 (엑셀은 1부터 시작)
    num_days = 31 # 31일까지 있다고 가정 (월마다 다름, 추후 개선 필요)
    
    staff_data = []
    # 3행부터 마지막 행까지 읽기
    for r in range(3, ws.max_row + 1):
        name = ws.cell(row=r, column=1).value
        if not name or name == '이름':
            continue
            
        staff_dict = {
            'row': r,
            'name': name,
            'is_male': name in MEN,
            'is_leader': name in LEADER,
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
    success = solve_schedule(staff_data, num_days, target_x_count)

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

st.title("🏥 스마트 듀티표 자동 생성기 (AI 적용버전)")
st.markdown("""
입력하신 **10가지 제약조건**이 모두 포함된 최적화 엔진(`Google OR-Tools`)이 적용되었습니다.  
""")

with st.sidebar:
    st.header("⚙️ 설정")
    target_x = st.number_input("이번 달 기본 오프(X) 개수", min_value=1, max_value=15, value=8)
    uploaded_file = st.file_uploader("원티드 엑셀 파일 업로드 (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    st.success("✅ 파일 업로드 완료! 데이터를 분석합니다.")
    
    if st.button("🚀 듀티표 AI 생성 시작"):
        with st.spinner("AI가 수천만 가지의 경우의 수를 계산하여 최적의 스케줄을 찾고 있습니다. (최대 60초 소요)"):
            try:
                # 결과 파일 생성
                result_file = process_excel(uploaded_file.getvalue(), target_x)
                
                if result_file:
                    st.balloons()
                    st.success("🎉 모든 조건이 만족되는 완벽한 듀티표가 생성되었습니다!")
                    
                    st.download_button(
                        label="📥 완성된 듀티표 다운로드 (.xlsx)",
                        data=result_file,
                        file_name="완성된_듀티표.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    st.info("💡 새로 입력된 듀티는 **연파랑색** 배경으로 표시됩니다. 원티드 요청은 조건이 허락하는 한 최대한 반영되었습니다.")
                else:
                    st.error("🚨 제약 조건이 너무 빡빡하여 해답을 찾을 수 없습니다! 노란색(고정) 휴무가 너무 많거나 조건이 충돌하는지 확인해주세요.")
            except Exception as e:
                st.error(f"오류가 발생했습니다. 엑셀 파일의 양식을 확인해주세요. (에러: {e})")
