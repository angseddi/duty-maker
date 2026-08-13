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
# 리더와 보조리더 명단
LEADER = ["용하영", "최충일", "박세은", "김소은", "윤지선", "이소희", "정하림", "최아라"]
SUB_LEADER = ["김민지", "박우영", "오지은"]

def is_cell_colored(cell):
    """셀에 배경색(노란색 등)이 칠해져 있는지 확인"""
    if cell.fill and cell.fill.patternType == 'solid':
        if cell.fill.fgColor.rgb not in ['00000000', 'FFFFFFFF', None]:
            return True
    return False

def solve_schedule(staff_data, num_days, base_x_count, min_d, max_d, min_e, max_e, min_n, max_n):
    model = cp_model.CpModel()
    num_staff = len(staff_data)
    
    # 1. 변수 생성: shifts[직원인덱스, 날짜인덱스, 근무형태인덱스] = 1(배정됨) or 0(배정안됨)
    shifts = {}
    for e in range(num_staff):
        for d in range(num_days):
            for s in range(4): # 0:D, 1:E, 2:N, 3:O
                shifts[(e, d, s)] = model.NewBoolVar(f'shift_e{e}_d{d}_s{s}')

    objective_terms = []

    # --- 제약 조건 (하드 제약) 적용 ---
    
    for e in range(num_staff):
        # 규칙 1: 하루에 무조건 1개의 근무(또는 휴무)만 배정
        for d in range(num_days):
            model.AddExactlyOne(shifts[(e, d, s)] for s in range(4))

        # 규칙 2: 역방향 교대 금지 (E->D, N->D, N->E 불가)
        for d in range(num_days - 1):
            model.AddImplication(shifts[(e, d, 1)], shifts[(e, d+1, 0)].Not()) # E->D 불가
            model.AddImplication(shifts[(e, d, 2)], shifts[(e, d+1, 0)].Not()) # N->D 불가
            model.AddImplication(shifts[(e, d, 2)], shifts[(e, d+1, 1)].Not()) # N->E 불가

        # 규칙 3: 최대 5일 연속 근무 허용 (6일 중 최소 1일은 휴무)
        for d in range(num_days - 5):
            model.Add(sum(shifts[(e, d+i, 3)] for i in range(6)) >= 1)

        # 규칙 4: 나이트(N)는 최대 3일 연속까지만
        for d in range(num_days - 3):
            model.Add(sum(shifts[(e, d+i, 2)] for i in range(4)) <= 3)
            
        # 규칙 5: N 근무 후 최소 2일 오프 (N -> O -> 일 금지)
        for d in range(num_days - 2):
            model.Add(shifts[(e, d, 2)] + shifts[(e, d+1, 3)] + (1 - shifts[(e, d+2, 3)]) <= 2)

        # 규칙 6: 오프(X) 개수 맞추기 (여유분 ±1 허용하여 에러 방지)
        target_offs = base_x_count if staff_data[e]['is_male'] else base_x_count + 1
        off_count = sum(shifts[(e, d, 3)] for d in range(num_days))
        model.Add(off_count >= target_offs - 1)
        model.Add(off_count <= target_offs + 1)
        
        # 정확히 오프 개수를 맞추면 점수 부여 (소프트 제약)
        is_perfect_off = model.NewBoolVar(f'perf_off_{e}')
        model.Add(off_count == target_offs).OnlyEnforceIf(is_perfect_off)
        objective_terms.append(50 * is_perfect_off)

        # 규칙 7: 퐁당퐁당(Single X) 휴무 조건
        for d in range(1, num_days - 1):
            # E-X-D, N-X-D, N-X-E 절대 금지 (하드 제약)
            model.AddBoolOr([shifts[(e, d-1, 1)].Not(), shifts[(e, d, 3)].Not(), shifts[(e, d+1, 0)].Not()]) # E-X-D 금지
            model.AddBoolOr([shifts[(e, d-1, 2)].Not(), shifts[(e, d, 3)].Not(), shifts[(e, d+1, 0)].Not()]) # N-X-D 금지
            model.AddBoolOr([shifts[(e, d-1, 2)].Not(), shifts[(e, d, 3)].Not(), shifts[(e, d+1, 1)].Not()]) # N-X-E 금지
            
            # 그 외 Single X 는 가능하면 피하도록 페널티 부과 (어쩔 수 없을 때만 허용)
            single_x_penalty = model.NewBoolVar(f'single_x_pen_{e}_{d}')
            model.Add(single_x_penalty >= (1 - shifts[(e, d-1, 3)]) + shifts[(e, d, 3)] + (1 - shifts[(e, d+1, 3)]) - 2)
            objective_terms.append(-30 * single_x_penalty) 

        # 규칙 8: 각 사람당 D와 E 갯수는 최대한 비슷하게
        d_count = sum(shifts[(e, d, 0)] for d in range(num_days))
        e_count = sum(shifts[(e, d, 1)] for d in range(num_days))
        de_diff = model.NewIntVar(0, num_days, f'de_diff_{e}')
        model.Add(de_diff >= d_count - e_count)
        model.Add(de_diff >= e_count - d_count)
        objective_terms.append(-5 * de_diff) # 편차가 클수록 페널티

    # 규칙 9: 노란색 고정 근무는 무조건 반영
    for e, staff in enumerate(staff_data):
        for d, shift_str in staff['fixed'].items():
            if shift_str in SHIFT_IDX:
                model.Add(shifts[(e, d, SHIFT_IDX[shift_str])] == 1)
            elif shift_str in ['X', 'SX', 'H', 'HX', 'TR']:
                model.Add(shifts[(e, d, 3)] == 1)

    # 규칙 10: 일별 D, E, N 최소/최대 인원 보장
    for d in range(num_days):
        model.Add(sum(shifts[(e, d, 0)] for e in range(num_staff)) >= min_d)
        model.Add(sum(shifts[(e, d, 0)] for e in range(num_staff)) <= max_d)
        
        model.Add(sum(shifts[(e, d, 1)] for e in range(num_staff)) >= min_e)
        model.Add(sum(shifts[(e, d, 1)] for e in range(num_staff)) <= max_e)
        
        model.Add(sum(shifts[(e, d, 2)] for e in range(num_staff)) >= min_n)
        model.Add(sum(shifts[(e, d, 2)] for e in range(num_staff)) <= max_n)

    # 규칙 11: 리더 + 보조리더 배치
    leader_and_sub_indices = [i for i, s in enumerate(staff_data) if s['is_leader'] or s['is_sub_leader']]
    leader_indices = [i for i, s in enumerate(staff_data) if s['is_leader']]
    
    if len(leader_and_sub_indices) > 0:
        for d in range(num_days):
            for s in [0, 1, 2]: # D, E, N
                # 하드 제약: 리더+보조리더 그룹에서 최소 1명은 무조건 있어야 함
                model.Add(sum(shifts[(e, d, s)] for e in leader_and_sub_indices) >= 1)
                
                # 소프트 제약: '찐' 리더가 들어가면 가산점 부여 (보조리더보다 찐 리더를 우선 배치)
                leader_count = sum(shifts[(e, d, s)] for e in leader_indices)
                objective_terms.append(10 * leader_count)

    # 규칙 12: 모든 사람의 N(나이트) 개수는 최대 1개 차이만 나도록
    min_n_shifts = model.NewIntVar(0, num_days, 'min_n_shifts')
    max_n_shifts = model.NewIntVar(0, num_days, 'max_n_shifts')
    model.Add(max_n_shifts - min_n_shifts <= 1)
    for e in range(num_staff):
        n_count = sum(shifts[(e, d, 2)] for d in range(num_days))
        model.Add(n_count >= min_n_shifts)
        model.Add(n_count <= max_n_shifts)

    # 규칙 13: 원티드(일반 글씨) 최대한 반영 (가산점)
    for e, staff in enumerate(staff_data):
        for d, wanted_shifts in staff['wanted'].items():
            for w_shift in wanted_shifts:
                if w_shift in SHIFT_IDX:
                    objective_terms.append(20 * shifts[(e, d, SHIFT_IDX[w_shift])])
                elif w_shift in ['X', 'SX', 'H', 'HX', 'TR']:
                    objective_terms.append(20 * shifts[(e, d, 3)])
    
    # 모델 풀이
    model.Maximize(sum(objective_terms))
    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = 60.0 # 탐색 시간 60초
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


def process_excel(file_content, target_x_count, min_d, max_d, min_e, max_e, min_n, max_n):
    wb = openpyxl.load_workbook(io.BytesIO(file_content), data_only=True)
    ws = wb.active

    # 날짜 컬럼 파악
    start_col = None
    num_days = 0
    for col in range(1, ws.max_column + 1):
        val = ws.cell(row=1, column=col).value
        if val and str(val).isdigit():
            if start_col is None:
                start_col = col
            num_days += 1
            
    if start_col is None:
        start_col = 2
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
            'is_leader': str(name).strip() in LEADER,
            'is_sub_leader': str(name).strip() in SUB_LEADER,
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
    success = solve_schedule(staff_data, num_days, target_x_count, min_d, max_d, min_e, max_e, min_n, max_n)

    if not success:
        return None

    # 성공 시 엑셀에 데이터 쓰기
    blue_fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
    
    for staff in staff_data:
        r = staff['row']
        for d in range(num_days):
            col = start_col + d
            if d not in staff['fixed']: 
                val = staff['final_schedule'].get(d, '')
                # 여자의 경우 추가된 오프 1개를 'HX'로 변환
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
st.markdown("입력하신 **제약조건**과 **리더/보조리더** 규칙을 모두 반영한 인공지능 최적화 모델입니다.")

with st.sidebar:
    st.header("⚙️ 기본 설정")
    target_x = st.number_input("이번 달 기본 오프(X) 개수", min_value=1, max_value=15, value=8)
    
    st.header("👥 일별 근무자 인원수")
    st.markdown("*(조건이 빡빡해서 에러가 날 경우 이 숫자를 조절해보세요)*")
    
    col1, col2 = st.columns(2)
    with col1:
        min_d = st.number_input("Day 최소", min_value=1, max_value=10, value=4)
        min_e = st.number_input("Evening 최소", min_value=1, max_value=10, value=4)
        min_n = st.number_input("Night 최소", min_value=1, max_value=10, value=4)
    with col2:
        max_d = st.number_input("Day 최대", min_value=1, max_value=15, value=6)
        max_e = st.number_input("Evening 최대", min_value=1, max_value=15, value=6)
        max_n = st.number_input("Night 최대", min_value=1, max_value=15, value=6)
    
    st.header("📂 파일 업로드")
    uploaded_file = st.file_uploader("원티드 엑셀 파일 업로드 (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    st.success("✅ 파일 업로드 완료! 데이터를 분석합니다.")
    
    if st.button("🚀 듀티표 AI 생성 시작"):
        with st.spinner("AI가 수천만 가지의 경우의 수를 계산하여 최적의 스케줄을 찾고 있습니다. (최대 60초 소요)"):
            try:
                # 결과 파일 생성
                result_file = process_excel(uploaded_file.getvalue(), target_x, min_d, max_d, min_e, max_e, min_n, max_n)
                
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
                    st.error("🚨 제약 조건이 너무 빡빡하여 해답을 찾을 수 없습니다! 설정하신 최소/최대 인원을 조절하거나 원티드를 확인해주세요.")
            except Exception as e:
                st.error(f"오류가 발생했습니다. 엑셀 파일의 양식을 확인해주세요. (에러: {e})")
