import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Alignment
import io
import random
from ortools.sat.python import cp_model

# --- 설정 및 전역 변수 ---
SHIFTS = ['D', 'E', 'N', 'O'] # O는 Off(휴무)를 의미 (X, SX, H, HX, TR 통합)
SHIFT_IDX = {'D': 0, 'E': 1, 'N': 2, 'O': 3}
REV_SHIFT_IDX = {0: 'D', 1: 'E', 2: 'N', 3: 'X'}

MEN = ["최충일", "윤진호", "이용재"]
LEADER = ["용하영", "최충일", "박세은", "김소은", "윤지선", "이소희", "정하림", "최아라"]
SUB_LEADER = ["김민지", "박우영", "오지은"]

def is_cell_colored(cell):
    """셀에 배경색(노란색 등)이 칠해져 있는지 확인"""
    if cell.fill and cell.fill.patternType == 'solid':
        if cell.fill.fgColor.rgb not in ['00000000', 'FFFFFFFF', None]:
            return True
    return False

def solve_schedule(staff_data, num_days, num_history, base_x_count, min_d, max_d, min_e, max_e, min_n, max_n):
    model = cp_model.CpModel()
    num_staff = len(staff_data)
    total_days = num_history + num_days
    
    # 1. 변수 생성: shifts[직원인덱스, 날짜인덱스, 근무형태인덱스] = 1(배정됨) or 0(배정안됨)
    shifts = {}
    for e in range(num_staff):
        for d in range(total_days):
            for s in range(4): # 0:D, 1:E, 2:N, 3:O
                shifts[(e, d, s)] = model.NewBoolVar(f'shift_e{e}_d{d}_s{s}')

    objective_terms = []

    # --- 제약 조건 (하드 및 소프트 제약) 적용 ---
    
    for e, staff in enumerate(staff_data):
        # [A] 이전달 기록 고정
        for d in range(num_history):
            s_val = staff['history'][d]
            s_idx = SHIFT_IDX.get(s_val, 3) # 기본값 O
            model.Add(shifts[(e, d, s_idx)] == 1)

        # [B] 당월 노란색 고정 근무 반영
        for d, shift_str in staff['fixed'].items():
            abs_d = num_history + d
            if 'D' in shift_str: s_idx = 0
            elif 'E' in shift_str: s_idx = 1
            elif 'N' in shift_str: s_idx = 2
            else: s_idx = 3
            model.Add(shifts[(e, abs_d, s_idx)] == 1)

        # [C] 기본 규칙 (전체 기간에 대해 적용)
        for d in range(total_days):
            model.AddExactlyOne(shifts[(e, d, s)] for s in range(4))

        for d in range(total_days - 1):
            # 역방향 교대 금지 (E->D, N->D, N->E 불가)
            model.AddImplication(shifts[(e, d, 1)], shifts[(e, d+1, 0)].Not())
            model.AddImplication(shifts[(e, d, 2)], shifts[(e, d+1, 0)].Not())
            model.AddImplication(shifts[(e, d, 2)], shifts[(e, d+1, 1)].Not())

        for d in range(total_days - 5):
            # 최대 5일 연속 근무 허용 (6일 중 최소 1일은 휴무)
            model.Add(sum(shifts[(e, d+i, 3)] for i in range(6)) >= 1)

        for d in range(total_days - 3):
            # 나이트(N)는 최대 3일 연속까지만
            model.Add(sum(shifts[(e, d+i, 2)] for i in range(4)) <= 3)

        for d in range(total_days - 2):
            # N 근무 후 최소 2일 오프 (N -> O -> 일 금지)
            model.Add(shifts[(e, d, 2)] + shifts[(e, d+1, 3)] + (1 - shifts[(e, d+2, 3)]) <= 2)

        for d in range(1, total_days - 1):
            # 독성 퐁당퐁당(Single X) 패턴 절대 금지: E-X-D, N-X-D, N-X-E
            model.AddBoolOr([shifts[(e, d-1, 1)].Not(), shifts[(e, d, 3)].Not(), shifts[(e, d+1, 0)].Not()])
            model.AddBoolOr([shifts[(e, d-1, 2)].Not(), shifts[(e, d, 3)].Not(), shifts[(e, d+1, 0)].Not()])
            model.AddBoolOr([shifts[(e, d-1, 2)].Not(), shifts[(e, d, 3)].Not(), shifts[(e, d+1, 1)].Not()])
            
            # 기타 Single X는 가급적 피하도록 페널티 부과
            single_x = model.NewBoolVar(f'single_x_{e}_{d}')
            model.Add(single_x >= (1 - shifts[(e, d-1, 3)]) + shifts[(e, d, 3)] + (1 - shifts[(e, d+1, 3)]) - 2)
            objective_terms.append(-20 * single_x) 

        # [D] 당월 월간 목표 (당월 기간만 계산)
        target_offs = int(base_x_count) if staff['is_male'] else int(base_x_count) + 1
        off_count = sum(shifts[(e, num_history + d, 3)] for d in range(num_days))
        
        model.Add(off_count >= target_offs - 1)
        model.Add(off_count <= target_offs + 1)
        
        diff_off = model.NewIntVar(0, num_days, f'diff_off_{e}')
        model.Add(diff_off >= off_count - target_offs)
        model.Add(diff_off >= target_offs - off_count)
        objective_terms.append(-30 * diff_off)

        # D와 E 갯수 비슷하게 맞추기
        d_count = sum(shifts[(e, num_history + d, 0)] for d in range(num_days))
        e_count = sum(shifts[(e, num_history + d, 1)] for d in range(num_days))
        de_diff = model.NewIntVar(0, num_days, f'de_diff_{e}')
        model.Add(de_diff >= d_count - e_count)
        model.Add(de_diff >= e_count - d_count)
        objective_terms.append(-5 * de_diff)

    # [E] 전역 제약 조건
    leader_and_sub_indices = [i for i, s in enumerate(staff_data) if s['is_leader'] or s['is_sub_leader']]
    leader_indices = [i for i, s in enumerate(staff_data) if s['is_leader']]
    
    for d in range(num_days):
        abs_d = num_history + d
        for s in [0, 1, 2]: # D, E, N
            # 하드 제약: 리더+보조리더 그룹에서 최소 1명은 무조건 필수
            if leader_and_sub_indices:
                model.Add(sum(shifts[(e, abs_d, s)] for e in leader_and_sub_indices) >= 1)
            
            # 소프트 제약: '찐' 리더가 들어가면 가산점 부여 (리더 우선 배치)
            if leader_indices:
                leader_count = sum(shifts[(e, abs_d, s)] for e in leader_indices)
                objective_terms.append(25 * leader_count)

        # 일별 D, E, N 인원 제한
        model.Add(sum(shifts[(e, abs_d, 0)] for e in range(num_staff)) >= int(min_d))
        model.Add(sum(shifts[(e, abs_d, 0)] for e in range(num_staff)) <= int(max_d))
        
        model.Add(sum(shifts[(e, abs_d, 1)] for e in range(num_staff)) >= int(min_e))
        model.Add(sum(shifts[(e, abs_d, 1)] for e in range(num_staff)) <= int(max_e))
        
        model.Add(sum(shifts[(e, abs_d, 2)] for e in range(num_staff)) >= int(min_n))
        model.Add(sum(shifts[(e, abs_d, 2)] for e in range(num_staff)) <= int(max_n))

    # 모든 사람의 N(나이트) 개수는 최대 1개 차이만 나도록
    n_counts = []
    for e in range(num_staff):
        n_count = model.NewIntVar(0, num_days, f'n_count_{e}')
        model.Add(n_count == sum(shifts[(e, num_history + d, 2)] for d in range(num_days)))
        n_counts.append(n_count)
        
    max_n_count = model.NewIntVar(0, num_days, 'max_n_count')
    min_n_count = model.NewIntVar(0, num_days, 'min_n_count')
    model.AddMaxEquality(max_n_count, n_counts)
    model.AddMinEquality(min_n_count, n_counts)
    model.Add(max_n_count - min_n_count <= 1)

    # [F] 원티드(일반 글씨) 최대한 반영 (가산점)
    for e, staff in enumerate(staff_data):
        for d, wanted_shifts in staff['wanted'].items():
            abs_d = num_history + d
            for w_shift in wanted_shifts:
                if w_shift in SHIFT_IDX:
                    objective_terms.append(20 * shifts[(e, abs_d, SHIFT_IDX[w_shift])])
                elif w_shift in ['X', 'SX', 'H', 'HX', 'TR']:
                    objective_terms.append(20 * shifts[(e, abs_d, 3)])
    
    model.Maximize(sum(objective_terms))
    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = 120.0
    status = solver.Solve(model)

    if status == cp_model.OPTIMAL or status == cp_model.FEASIBLE:
        for e, staff in enumerate(staff_data):
            for d in range(num_days):
                for s in range(4):
                    if solver.Value(shifts[(e, num_history + d, s)]) == 1:
                        staff['final_schedule'][d] = REV_SHIFT_IDX[s]
        return True
    else:
        return False


def process_excel(file_content, target_x_count, min_d, max_d, min_e, max_e, min_n, max_n):
    wb = openpyxl.load_workbook(io.BytesIO(file_content))
    ws = wb.active

    days_of_week = ['월', '화', '수', '목', '금', '토', '일']
    cal_cols = []
    summary_cols = {}
    
    for col in range(2, ws.max_column + 1):
        v1 = str(ws.cell(row=1, column=col).value).strip()
        v2 = str(ws.cell(row=2, column=col).value).strip()
        
        if v2 in days_of_week:
            cal_cols.append(col)
        elif v1 in ['D', 'E', 'N', 'X', 'H', 'HX', 'SX', 'I', '총합']:
            summary_cols[v1] = col

    history_cols = cal_cols[:5]
    current_cols = cal_cols[5:]
    num_history = len(history_cols)
    num_days = len(current_cols)

    staff_data = []
    for r in range(3, ws.max_row + 1):
        name = ws.cell(row=r, column=1).value
        if not name or str(name).strip() == '이름':
            continue
            
        name = str(name).strip()
        staff_dict = {
            'row': r,
            'name': name,
            'is_male': name in MEN,
            'is_leader': name in LEADER,
            'is_sub_leader': name in SUB_LEADER,
            'history': [],
            'fixed': {},
            'wanted': {},
            'final_schedule': {}
        }
        
        for col in history_cols:
            val = str(ws.cell(row=r, column=col).value).strip().upper()
            if val in ['D', 'E', 'N']:
                staff_dict['history'].append(val)
            else:
                staff_dict['history'].append('O')
                
        for d, col in enumerate(current_cols):
            cell = ws.cell(row=r, column=col)
            val = str(cell.value).strip().upper() if cell.value else ""
            
            if val and val != "NONE":
                options = [x.strip() for x in val.split(',')]
                if is_cell_colored(cell):
                    staff_dict['fixed'][d] = options[0]
                else:
                    staff_dict['wanted'][d] = options
                    
        staff_data.append(staff_dict)

    success = solve_schedule(staff_data, num_days, num_history, target_x_count, min_d, max_d, min_e, max_e, min_n, max_n)

    if not success:
        return None

    blue_fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
    center_align = Alignment(horizontal='center', vertical='center')
    
    for staff in staff_data:
        r = staff['row']
        counts = {'D':0, 'E':0, 'N':0, 'X':0, 'H':0, 'HX':0, 'SX':0, 'I':0}
        
        for d in range(num_days):
            col = current_cols[d]
            is_fixed = d in staff['fixed']
            
            if is_fixed:
                val = staff['fixed'][d]
            else:
                val = staff['final_schedule'].get(d, '')
                if val == 'X' and not staff['is_male'] and 'HX' not in staff['fixed'].values() and counts['HX'] == 0:
                    val = 'HX'
                
                cell = ws.cell(row=r, column=col)
                cell.value = val
                cell.fill = blue_fill 
                
            cell = ws.cell(row=r, column=col)
            cell.alignment = center_align
                
            if val in counts:
                counts[val] += 1
            elif val in ['O', 'TR']:
                counts['X'] += 1
                
        for key, col in summary_cols.items():
            if key == '총합':
                ws.cell(row=r, column=col).value = sum(counts.values())
            elif key in counts:
                ws.cell(row=r, column=col).value = counts[key]
            ws.cell(row=r, column=col).alignment = center_align

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
    target_x = st.number_input("이번 달 기본 오프(X) 개수", min_value=1, max_value=15, value=8, step=1)
    
    st.header("👥 일별 근무자 인원수")
    st.markdown("*(조건이 빡빡해서 에러가 날 경우 이 숫자를 조절해보세요)*")
    
    col1, col2 = st.columns(2)
    with col1:
        min_d = st.number_input("Day 최소", min_value=1, max_value=10, value=4, step=1)
        min_e = st.number_input("Evening 최소", min_value=1, max_value=10, value=4, step=1)
        min_n = st.number_input("Night 최소", min_value=1, max_value=10, value=4, step=1)
    with col2:
        max_d = st.number_input("Day 최대", min_value=1, max_value=15, value=6, step=1)
        max_e = st.number_input("Evening 최대", min_value=1, max_value=15, value=6, step=1)
        max_n = st.number_input("Night 최대", min_value=1, max_value=15, value=6, step=1)
    
    st.header("📂 파일 업로드")
    uploaded_file = st.file_uploader("원티드 엑셀 파일 업로드 (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    st.success("✅ 파일 업로드 완료! 데이터를 분석합니다.")
    
    if st.button("🚀 듀티표 AI 생성 시작"):
        with st.spinner("AI가 수천만 가지의 경우의 수를 계산하여 최적의 스케줄을 찾고 있습니다. (최대 120초 소요)"):
            try:
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
                    st.info("💡 새로 입력된 듀티는 **연파랑색** 배경으로 표시됩니다. 오른쪽에 근무 갯수 통계도 자동으로 입력되었습니다.")
                else:
                    st.error("🚨 제약 조건이 너무 빡빡하여 해답을 찾을 수 없습니다! 설정하신 최소/최대 인원을 조절하거나 원티드를 확인해주세요.")
            except Exception as e:
                st.error(f"오류가 발생했습니다. 엑셀 파일의 양식을 확인해주세요. (에러: {e})")
