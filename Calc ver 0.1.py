import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import calendar
import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
import io

# 페이지 설정
st.set_page_config(
    page_title="재고 관리 시스템",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded"
)

def is_weekend(date):
    """주말인지 확인하는 함수 (토요일: 5, 일요일: 6)"""
    return date.weekday() >= 5

def get_next_monday(date):
    """다음 월요일 날짜를 반환하는 함수"""
    days_until_monday = (7 - date.weekday()) % 7
    if days_until_monday == 0:
        days_until_monday = 7
    return date + timedelta(days=days_until_monday)

def parse_date(date_str):
    """날짜 문자열을 datetime 객체로 변환"""
    if isinstance(date_str, datetime):
        return date_str.replace(hour=0, minute=0, second=0, microsecond=0)
    if isinstance(date_str, str):
        return datetime.strptime(date_str.strip().split()[0], '%Y-%m-%d')
    return None

def display_calendar(results_data):
    """스트림릿에서 캘린더 형식으로 입고 일정을 표시하는 함수"""
    try:
        # 모든 입고 날짜 수집
        all_dates = []
        item_dates = {}  # {날짜: [품목명 리스트]}

        for item in results_data:
            if item['입고 필요일'] != '입고 필요 없음':
                dates = item['입고 필요일'].split(', ')
                for date in dates:
                    try:
                        parsed_date = datetime.strptime(date, '%Y-%m-%d')
                        all_dates.append(parsed_date)
                        if date not in item_dates:
                            item_dates[date] = []
                        item_dates[date].append(item['품목명'])
                    except ValueError:
                        continue

        if not all_dates:
            st.info("📅 입고 일정이 없습니다.")
            return

        # 캘린더에 표시할 기간 결정
        start_date = min(all_dates)
        end_date = max(all_dates)

        st.info(f"📅 입고 일정 기간: {start_date.strftime('%Y-%m-%d')} ~ {end_date.strftime('%Y-%m-%d')}")

        # 각 월별로 캘린더 표시
        current_date = start_date.replace(day=1)
        end_month = end_date.replace(day=1)

        while current_date <= end_month:
            # 월 제목
            st.markdown(f"### 📅 {current_date.strftime('%Y년 %m월')}")

            # 요일 헤더
            col_headers = st.columns(7)
            days = ['월', '화', '수', '목', '금', '토', '일']
            for i, day in enumerate(days):
                with col_headers[i]:
                    st.markdown(f"**{day}**")

            # 해당 월의 캘린더 생성
            cal = calendar.monthcalendar(current_date.year, current_date.month)

            for week in cal:
                cols = st.columns(7)
                for day_idx, day in enumerate(week):
                    with cols[day_idx]:
                        if day == 0:
                            # 빈 날짜
                            st.markdown("<div style='height: 80px;'></div>", unsafe_allow_html=True)
                        else:
                            # 날짜가 있는 경우
                            date_str = f"{current_date.year}-{current_date.month:02d}-{day:02d}"

                            # 주말 여부 확인
                            is_weekend_day = day_idx >= 5

                            # 입고 일정이 있는지 확인
                            has_schedule = date_str in item_dates

                            # 스타일 적용
                            if has_schedule:
                                # 입고 일정이 있는 날
                                bg_color = "#FFE6E6"  # 연한 빨간색
                                border_color = "#FF4444"
                                text_color = "#000000"
                            elif is_weekend_day:
                                # 주말
                                bg_color = "#F0F0F0"  # 연한 회색
                                border_color = "#CCCCCC"
                                text_color = "#666666"
                            else:
                                # 평일
                                bg_color = "#FFFFFF"  # 흰색
                                border_color = "#DDDDDD"
                                text_color = "#000000"

                            # 날짜 셀 내용 구성
                            cell_content = f"""
                            <div style='
                                background-color: {bg_color};
                                border: 2px solid {border_color};
                                border-radius: 5px;
                                padding: 5px;
                                margin: 2px;
                                min-height: 70px;
                                color: {text_color};
                                font-size: 12px;
                            '>
                                <div style='font-weight: bold; margin-bottom: 3px;'>{day}</div>
                            """

                            # 입고 품목이 있는 경우 추가
                            if has_schedule:
                                for item_name in item_dates[date_str]:
                                    # 품목명이 너무 길면 줄임
                                    display_name = item_name if len(item_name) <= 8 else item_name[:8] + "..."
                                    cell_content += f"<div style='background-color: #FF6666; color: white; padding: 1px 3px; margin: 1px 0; border-radius: 3px; font-size: 10px;'>{display_name}</div>"

                            cell_content += "</div>"

                            st.markdown(cell_content, unsafe_allow_html=True)

            # 월 간의 간격
            st.markdown("---")

            # 다음 달로 이동
            if current_date.month == 12:
                current_date = current_date.replace(year=current_date.year + 1, month=1)
            else:
                current_date = current_date.replace(month=current_date.month + 1)

        # 범례 표시
        st.markdown("### 📖 범례")
        legend_cols = st.columns(3)

        with legend_cols[0]:
            st.markdown("""
            <div style='background-color: #FFE6E6; border: 2px solid #FF4444; padding: 10px; border-radius: 5px; text-align: center;'>
                <strong>🔴 입고 일정 있음</strong>
            </div>
            """, unsafe_allow_html=True)

        with legend_cols[1]:
            st.markdown("""
            <div style='background-color: #F0F0F0; border: 2px solid #CCCCCC; padding: 10px; border-radius: 5px; text-align: center;'>
                <strong>⚪ 주말</strong>
            </div>
            """, unsafe_allow_html=True)

        with legend_cols[2]:
            st.markdown("""
            <div style='background-color: #FFFFFF; border: 2px solid #DDDDDD; padding: 10px; border-radius: 5px; text-align: center;'>
                <strong>⚫ 평일</strong>
            </div>
            """, unsafe_allow_html=True)

        # 입고 일정 상세 정보
        st.markdown("### 📋 입고 일정 상세")

        # 날짜별로 정렬하여 표시
        sorted_dates = sorted(item_dates.items())

        for date_str, items in sorted_dates:
            date_obj = datetime.strptime(date_str, '%Y-%m-%d')
            weekday = ['월', '화', '수', '목', '금', '토', '일'][date_obj.weekday()]

            with st.expander(f"📅 {date_str} ({weekday}) - {len(items)}개 품목"):
                for i, item in enumerate(items, 1):
                    st.write(f"{i}. {item}")

    except Exception as e:
        st.error(f"캘린더 표시 중 오류 발생: {str(e)}")
        st.info("입고 일정을 목록 형태로 표시합니다.")

        # 대안: 간단한 목록 형태로 표시
        all_purchase_dates = []
        for item in results_data:
            if item['입고 필요일'] != '입고 필요 없음':
                dates = item['입고 필요일'].split(', ')
                for date in dates:
                    all_purchase_dates.append({
                        '날짜': date,
                        '품목명': item['품목명']
                    })

        if all_purchase_dates:
            schedule_df = pd.DataFrame(all_purchase_dates)
            st.dataframe(schedule_df, use_container_width=True)

def calculate_purchase_date(df, web_exclude_dates=None, web_include_dates=None):
    """데이터프레임에서 안전재고 도달일과 입고 필요일을 계산하는 함수"""
    try:
        results = []

        # 웹에서 입력받은 날짜를 datetime 객체로 변환
        web_exclude_datetimes = []
        if web_exclude_dates:
            if isinstance(web_exclude_dates, (list, tuple)):
                web_exclude_datetimes = [datetime.combine(date, datetime.min.time()) for date in web_exclude_dates]
            else:
                web_exclude_datetimes = [datetime.combine(web_exclude_dates, datetime.min.time())]

        web_include_datetimes = []
        if web_include_dates:
            if isinstance(web_include_dates, (list, tuple)):
                web_include_datetimes = [datetime.combine(date, datetime.min.time()) for date in web_include_dates]
            else:
                web_include_datetimes = [datetime.combine(web_include_dates, datetime.min.time())]

        for _, row in df.iterrows():
            # 기본 데이터 추출
            item_code = row['원료코드명']
            item_name = row['원료명']
            current_inventory = float(row['현재 재고'])
            daily_usage = (float(row['최소 사용량']) + float(row['최대 사용량'])) / 2
            calculation_days = int(row['계산 일자'])
            safety_stock = float(row['안전 재고'])
            purchase_amount = float(row['1회 구매량'])

            # 시작일 설정
            start_date = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
            end_date = start_date + timedelta(days=calculation_days)

            # Excel 파일의 제외 날짜 처리 (열이 없어도 동작)
            excel_exclude_dates = []
            if pd.notna(row.get('제외 날짜')):
                date_strings = str(row.get('제외 날짜')).split(',')
                excel_exclude_dates = [parse_date(date_str) for date_str in date_strings if parse_date(date_str)]

            # Excel 파일의 포함 날짜 처리 (열이 없어도 동작)
            excel_include_dates = []
            if pd.notna(row.get('포함 날짜')):
                date_strings = str(row.get('포함 날짜')).split(',')
                excel_include_dates = [parse_date(date_str) for date_str in date_strings if parse_date(date_str)]

            # 최종 제외/포함 날짜 결합 (웹 입력이 우선, Excel 데이터와 합침)
            final_exclude_dates = web_exclude_datetimes + excel_exclude_dates
            final_include_dates = web_include_datetimes + excel_include_dates

            # 안전재고 도달일 계산
            current_date = start_date
            remaining_inventory = current_inventory
            safety_stock_info = []

            while current_date <= end_date:
                # 주말이면서 포함 날짜에 없는 경우 스킵
                is_included_date = any(include_date and include_date.date() == current_date.date()
                                    for include_date in final_include_dates)
                if is_weekend(current_date) and not is_included_date:
                    current_date += timedelta(days=1)
                    continue

                # 제외일인 경우 스킵
                if any(exclude_date and exclude_date.date() == current_date.date()
                      for exclude_date in final_exclude_dates):
                    current_date += timedelta(days=1)
                    continue

                # 재고가 안전재고 이하가 되는지 확인
                if remaining_inventory <= safety_stock:
                    # 포함일(주말)인 경우 다음 월요일을 입고일로 설정
                    actual_purchase_date = (
                        get_next_monday(current_date) if is_included_date
                        else current_date
                    )

                    safety_stock_info.append({
                        'reach_date': current_date.strftime('%Y-%m-%d'),
                        'purchase_date': actual_purchase_date.strftime('%Y-%m-%d'),
                        'stock_before': f"{remaining_inventory:.2f}",
                        'stock_after': f"{remaining_inventory + purchase_amount:.2f}"
                    })
                    remaining_inventory += purchase_amount

                # 재고 계산
                remaining_inventory -= daily_usage
                current_date += timedelta(days=1)

            # 결과 저장
            if safety_stock_info:
                results.append({
                    '품목코드': item_code,
                    '품목명': item_name,
                    '안전재고 도달 일자': ', '.join(info['reach_date'] for info in safety_stock_info),
                    '도달시 재고량': ', '.join(info['stock_before'] for info in safety_stock_info),
                    '입고후 재고량': ', '.join(info['stock_after'] for info in safety_stock_info),
                    '입고 필요일': ', '.join(info['purchase_date'] for info in safety_stock_info)
                })
            else:
                results.append({
                    '품목코드': item_code,
                    '품목명': item_name,
                    '안전재고 도달 일자': '계산기간 내 안전재고 미도달',
                    '도달시 재고량': '-',
                    '입고후 재고량': '-',
                    '입고 필요일': '입고 필요 없음'
                })

        return results

    except Exception as e:
        st.error(f"계산 중 에러 발생: {str(e)}")
        return None

def calculate_monthly_purchase(results_data, df):
    """월별 구매 필요량을 계산하는 함수"""
    monthly_data = []

    for _, row in df.iterrows():
        item_code = row['원료코드명']
        item_name = row['원료명']

        # 해당 품목의 결과 데이터 찾기
        item_result = next((r for r in results_data if r['품목코드'] == item_code), None)

        if item_result and item_result['입고 필요일'] != '입고 필요 없음':
            # 입고일, 도달시 재고량, 입고후 재고량을 가져옴
            purchase_dates = item_result['입고 필요일'].split(', ')
            stock_before = [float(x) for x in item_result['도달시 재고량'].split(', ')]
            stock_after = [float(x) for x in item_result['입고후 재고량'].split(', ')]

            # 각 입고시점의 실제 구매량 계산
            actual_purchases = [after - before for before, after in zip(stock_before, stock_after)]

            monthly_counts = {}

            for date_str, purchase_amount in zip(purchase_dates, actual_purchases):
                date = datetime.strptime(date_str, '%Y-%m-%d')
                year = date.year
                month = date.month
                month_key = f"{year}-{month:02d}"
                display_key = f"{month}월({year}년)"

                if display_key not in monthly_counts:
                    monthly_counts[display_key] = {
                        'count': 0,
                        'amount': 0,
                        'sort_key': month_key
                    }
                monthly_counts[display_key]['count'] += 1
                monthly_counts[display_key]['amount'] += purchase_amount

            # 품목별 데이터 생성
            item_data = {
                '품목코드': item_code,
                '품목명': item_name
            }

            # 정렬된 순서대로 데이터 추가
            sorted_months = sorted(monthly_counts.items(),
                                 key=lambda x: x[1]['sort_key'])

            for month_display, data in sorted_months:
                item_data[f'{month_display} 입고 횟수'] = data['count']
                item_data[f'{month_display} 총 구매량'] = f"{data['amount']:.2f}"

            monthly_data.append(item_data)

    if not monthly_data:
        return pd.DataFrame(columns=['품목코드', '품목명'])

    return pd.DataFrame(monthly_data)

def create_purchase_order_df(results_data, df):
    """계산 결과를 발주서 생성창 규격으로 변환하는 함수
    입고일·품목명·입고량(KG)은 계산 결과에서, 나머지는 업로드 Excel의 발주 정보에서 가져옴
    총액(원) = 입고량(KG) × 단가(원)  ※ 단가는 KG당 단가 기준
    """
    order_columns = ['입고일', '품목명', '부가세 구분', '제품 구분', '업체명',
                     '단량(KG)', '입고량(KG)', '단가(원)', '총액(원)', '담당자 이메일']
    order_rows = []

    for _, row in df.iterrows():
        item_result = next((r for r in results_data if r['품목코드'] == row['원료코드명']), None)
        if not item_result or item_result['입고 필요일'] == '입고 필요 없음':
            continue

        purchase_amount = float(row['1회 구매량'])  # 입고량(KG)
        unit_price = pd.to_numeric(row.get('단가(원)', None), errors='coerce')
        total_price = round(purchase_amount * unit_price) if pd.notna(unit_price) else ''

        for date_str in item_result['입고 필요일'].split(', '):
            order_rows.append({
                '입고일': date_str,
                '품목명': row['원료명'],
                '부가세 구분': '' if pd.isna(row.get('부가세 구분')) else row.get('부가세 구분'),
                '제품 구분': '' if pd.isna(row.get('제품 구분')) else row.get('제품 구분'),
                '업체명': '' if pd.isna(row.get('업체명')) else row.get('업체명'),
                '단량(KG)': '' if pd.isna(row.get('단량(KG)')) else row.get('단량(KG)'),
                '입고량(KG)': purchase_amount,
                '단가(원)': unit_price if pd.notna(unit_price) else '',
                '총액(원)': total_price,
                '담당자 이메일': '' if pd.isna(row.get('담당자 이메일')) else str(row.get('담당자 이메일')).strip(),
            })

    if not order_rows:
        return pd.DataFrame(columns=order_columns)

    return pd.DataFrame(order_rows, columns=order_columns).sort_values(['입고일', '품목명']).reset_index(drop=True)

def create_calendar_sheet_safe(writer, results_data):
    """달력 형식으로 입고 일정을 표시하는 함수 (안전한 버전)"""
    try:
        # 모든 입고 날짜 수집
        all_dates = []
        item_dates = {}

        for item in results_data:
            if item['입고 필요일'] != '입고 필요 없음':
                dates = item['입고 필요일'].split(', ')
                for date in dates:
                    try:
                        parsed_date = datetime.strptime(date, '%Y-%m-%d')
                        all_dates.append(parsed_date)
                        if date not in item_dates:
                            item_dates[date] = []
                        item_dates[date].append(item['품목명'])
                    except ValueError:
                        continue

        if not all_dates:
            # 달력 시트 생성하되 메시지만 표시
            workbook = writer.book
            calendar_sheet = workbook.create_sheet('입고일정달력')
            calendar_sheet.cell(row=1, column=1, value="입고 일정이 없습니다.")
            return

        # 달력에 표시할 기간 결정
        start_date = min(all_dates)
        end_date = max(all_dates)

        # 각 월별로 달력 생성
        current_date = start_date.replace(day=1)
        end_month = end_date.replace(day=1)

        # 엑셀 워크북 가져오기
        workbook = writer.book
        calendar_sheet = workbook.create_sheet('입고일정달력')

        # 스타일 설정
        header_fill = PatternFill(start_color='CCE5FF', end_color='CCE5FF', fill_type='solid')
        weekend_fill = PatternFill(start_color='F2F2F2', end_color='F2F2F2', fill_type='solid')
        border = Border(left=Side(style='thin'), right=Side(style='thin'),
                       top=Side(style='thin'), bottom=Side(style='thin'))

        # 현재 행 위치
        current_row = 1

        while current_date <= end_month:
            # 월 제목 추가
            month_title = current_date.strftime('%Y년 %m월')
            calendar_sheet.merge_cells(f'A{current_row}:G{current_row}')
            title_cell = calendar_sheet.cell(row=current_row, column=1, value=month_title)
            title_cell.font = Font(bold=True, size=12)
            title_cell.alignment = Alignment(horizontal='center')

            # 요일 헤더 추가
            days = ['월', '화', '수', '목', '금', '토', '일']
            header_row = current_row + 1
            for col, day in enumerate(days, 1):
                cell = calendar_sheet.cell(row=header_row, column=col, value=day)
                cell.fill = header_fill
                cell.border = border
                cell.alignment = Alignment(horizontal='center')
                calendar_sheet.column_dimensions[get_column_letter(col)].width = 15

            # 달력 날짜 채우기
            cal = calendar.monthcalendar(current_date.year, current_date.month)
            for week_idx, week in enumerate(cal):
                row = current_row + 2 + week_idx
                for day_idx, day in enumerate(week):
                    cell = calendar_sheet.cell(row=row, column=day_idx + 1)
                    cell.border = border
                    cell.alignment = Alignment(horizontal='center', vertical='top', wrap_text=True)

                    if day != 0:
                        date_str = f"{current_date.year}-{current_date.month:02d}-{day:02d}"
                        cell_text = str(day)

                        # 입고 품목이 있는 경우 추가
                        if date_str in item_dates:
                            cell_text += "\n" + "\n".join(item_dates[date_str])

                        cell.value = cell_text

                        # 주말인 경우 배경색 지정
                        if day_idx >= 5:
                            cell.fill = weekend_fill

            # 다음 달력을 위한 간격 추가
            current_row += len(cal) + 4
            current_date = (current_date + timedelta(days=32)).replace(day=1)

        # 전체 셀 높이 조정
        for row_num in range(1, current_row):
            calendar_sheet.row_dimensions[row_num].height = 60

    except Exception as e:
        st.warning(f"달력 생성 중 오류: {str(e)}")
        # 기본 메시지 시트라도 생성
        workbook = writer.book
        calendar_sheet = workbook.create_sheet('입고일정달력')
        calendar_sheet.cell(row=1, column=1, value=f"달력 생성 중 오류 발생: {str(e)}")

def create_excel_file(results_data, monthly_df, po_df):
    """엑셀 파일을 메모리에 생성하여 반환"""
    output = io.BytesIO()

    try:
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            # 발주서 시트
            if not po_df.empty:
                po_df.to_excel(writer, sheet_name='발주서', index=False)
            else:
                pd.DataFrame({'메시지': ['계산 기간 내 발주가 필요한 품목이 없습니다.']}).to_excel(
                    writer, sheet_name='발주서', index=False)

            # 입고계획 시트
            output_df = pd.DataFrame(results_data)
            output_df.to_excel(writer, sheet_name='입고계획', index=False)

            # 월별구매량 시트
            if not monthly_df.empty:
                monthly_df.to_excel(writer, sheet_name='월별구매량', index=False)
            else:
                # 빈 데이터프레임이라도 시트는 생성
                pd.DataFrame({'메시지': ['계산 기간 내 구매가 필요한 품목이 없습니다.']}).to_excel(
                    writer, sheet_name='월별구매량', index=False)

            # 달력 시트 생성 (안전하게)
            try:
                create_calendar_sheet_safe(writer, results_data)
            except Exception as e:
                st.warning(f"달력 시트 생성 중 오류 발생: {str(e)}")

        output.seek(0)
        return output
    except Exception as e:
        st.error(f"Excel 파일 생성 중 오류: {str(e)}")
        return None

# Streamlit 앱 시작
def main():
    st.title("📦 재고 관리 시스템")
    st.markdown("---")

    # 사이드바
    st.sidebar.header("📝 사용 방법")
    st.sidebar.markdown("""
    1. **Excel 파일 업로드**: 재고 데이터가 포함된 Excel 파일을 업로드하세요.
    2. **데이터 확인**: 업로드된 데이터를 확인하세요.
    3. **계산 실행**: '계산 시작' 버튼을 클릭하세요.
    4. **결과 확인**: 입고 계획과 발주서, 월별 구매량을 확인하세요.
    5. **복사/다운로드**: 발주서 탭에서 복사하거나 Excel로 다운로드하세요.
    """)

    st.sidebar.markdown("---")
    st.sidebar.header("📋 필요한 열 정보")
    st.sidebar.markdown("""
    **필수 열 (재고 계산용):**
    - 원료코드명
    - 원료명
    - 현재 재고
    - 최소 사용량
    - 최대 사용량
    - 계산 일자
    - 안전 재고
    - 1회 구매량
    - 제외 날짜 (선택사항)
    - 포함 날짜 (선택사항)

    **발주서 정보 열 (선택사항):**
    - 부가세 구분
    - 제품 구분
    - 업체명
    - 단량(KG)
    - 단가(원)
    - 담당자 이메일
    """)

    # 파일 업로드
    uploaded_file = st.file_uploader(
        "Excel 파일을 업로드하세요",
        type=['xlsx', 'xls'],
        help="재고 데이터가 포함된 Excel 파일을 선택하세요."
    )

    if uploaded_file is not None:
        try:
            # 파일 읽기
            df = pd.read_excel(uploaded_file)

            # 다른 파일이 업로드되면 이전 계산 결과 초기화
            if st.session_state.get('results_file') != uploaded_file.name:
                st.session_state.pop('results', None)

            # 데이터 미리보기
            st.subheader("📊 업로드된 데이터 미리보기")
            st.dataframe(df.head(), use_container_width=True)

            # 데이터 정보
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("총 품목 수", len(df))
            with col2:
                st.metric("총 열 수", len(df.columns))
            with col3:
                st.metric("데이터 크기", f"{df.shape[0]} × {df.shape[1]}")

            # 필수 열 확인
            required_columns = [
                '원료코드명', '원료명', '현재 재고', '최소 사용량',
                '최대 사용량', '계산 일자', '안전 재고', '1회 구매량'
            ]

            missing_columns = [col for col in required_columns if col not in df.columns]

            if missing_columns:
                st.error(f"❌ 필수 열이 누락되었습니다: {', '.join(missing_columns)}")
                st.stop()
            else:
                st.success("✅ 모든 필수 열이 확인되었습니다!")

            # 발주서용 정보 컬럼 확인 (없으면 빈 값으로 생성)
            order_info_columns = ['부가세 구분', '제품 구분', '업체명', '단량(KG)', '단가(원)', '담당자 이메일']
            missing_order_cols = [col for col in order_info_columns if col not in df.columns]
            if missing_order_cols:
                st.warning(f"⚠️ 발주 정보 열이 없어 빈 값으로 처리됩니다: {', '.join(missing_order_cols)}")
                for col in missing_order_cols:
                    df[col] = ''

            # 계산 버튼
            st.markdown("---")
            st.subheader("📅 날짜 설정 (선택사항)")

            # 세션 상태 초기화
            if 'exclude_dates_list' not in st.session_state:
                st.session_state.exclude_dates_list = []
            if 'include_dates_list' not in st.session_state:
                st.session_state.include_dates_list = []

            # 날짜 설정을 위한 컬럼
            date_col1, date_col2 = st.columns(2)

            with date_col1:
                st.markdown("**🚫 제외할 날짜들**")
                st.caption("계산에서 제외할 날짜를 추가하세요 (휴일, 비가동일 등)")

                # 날짜 범위 설정
                today = datetime.now().date()
                max_date = today + timedelta(days=365)

                # 제외 날짜 추가
                exclude_date_input = st.date_input(
                    "제외할 날짜 선택",
                    value=today,
                    min_value=today,
                    max_value=max_date,
                    key="exclude_date_input"
                )

                col1, col2 = st.columns(2)
                with col1:
                    if st.button("➕ 제외 날짜 추가", use_container_width=True):
                        if exclude_date_input not in st.session_state.exclude_dates_list:
                            st.session_state.exclude_dates_list.append(exclude_date_input)
                            st.success(f"제외 날짜 추가: {exclude_date_input.strftime('%Y-%m-%d')}")
                        else:
                            st.warning("이미 추가된 날짜입니다.")

                with col2:
                    if st.button("🗑️ 전체 삭제", use_container_width=True):
                        st.session_state.exclude_dates_list = []
                        st.success("모든 제외 날짜가 삭제되었습니다.")

                # 추가된 제외 날짜 목록 표시
                if st.session_state.exclude_dates_list:
                    st.write("**추가된 제외 날짜들:**")
                    for i, date in enumerate(sorted(st.session_state.exclude_dates_list)):
                        col_date, col_del = st.columns([3, 1])
                        with col_date:
                            weekday = ['월', '화', '수', '목', '금', '토', '일'][date.weekday()]
                            st.write(f"• {date.strftime('%Y-%m-%d')} ({weekday})")
                        with col_del:
                            if st.button("❌", key=f"del_exclude_{i}", help="이 날짜 삭제"):
                                st.session_state.exclude_dates_list.remove(date)
                                st.rerun()
                else:
                    st.info("추가된 제외 날짜가 없습니다.")

            with date_col2:
                st.markdown("**✅ 포함할 날짜들**")
                st.caption("주말이지만 가동하는 날짜를 추가하세요")

                # 포함 날짜 추가
                include_date_input = st.date_input(
                    "포함할 날짜 선택",
                    value=today,
                    min_value=today,
                    max_value=max_date,
                    key="include_date_input"
                )

                col1, col2 = st.columns(2)
                with col1:
                    if st.button("➕ 포함 날짜 추가", use_container_width=True):
                        if include_date_input not in st.session_state.include_dates_list:
                            st.session_state.include_dates_list.append(include_date_input)
                            weekday = ['월', '화', '수', '목', '금', '토', '일'][include_date_input.weekday()]
                            st.success(f"포함 날짜 추가: {include_date_input.strftime('%Y-%m-%d')} ({weekday})")
                        else:
                            st.warning("이미 추가된 날짜입니다.")

                with col2:
                    if st.button("🗑️ 전체 삭제 ", use_container_width=True, key="clear_include"):
                        st.session_state.include_dates_list = []
                        st.success("모든 포함 날짜가 삭제되었습니다.")

                # 추가된 포함 날짜 목록 표시
                if st.session_state.include_dates_list:
                    st.write("**추가된 포함 날짜들:**")
                    for i, date in enumerate(sorted(st.session_state.include_dates_list)):
                        col_date, col_del = st.columns([3, 1])
                        with col_date:
                            weekday = ['월', '화', '수', '목', '금', '토', '일'][date.weekday()]
                            is_weekend_day = date.weekday() >= 5
                            weekend_text = " 🟡" if is_weekend_day else " 🔵"
                            st.write(f"• {date.strftime('%Y-%m-%d')} ({weekday}){weekend_text}")
                        with col_del:
                            if st.button("❌", key=f"del_include_{i}", help="이 날짜 삭제"):
                                st.session_state.include_dates_list.remove(date)
                                st.rerun()
                else:
                    st.info("추가된 포함 날짜가 없습니다.")

            # 날짜 설정 요약
            total_dates = len(st.session_state.exclude_dates_list) + len(st.session_state.include_dates_list)
            if total_dates > 0:
                st.success(f"📝 총 {total_dates}개의 날짜가 설정되었습니다. (제외: {len(st.session_state.exclude_dates_list)}개, 포함: {len(st.session_state.include_dates_list)}개)")
                st.caption("🔵 = 평일, 🟡 = 주말")

            if st.button("🚀 계산 시작", type="primary", use_container_width=True):
                with st.spinner("계산 중입니다... 잠시만 기다려주세요."):
                    # 계산 실행 후 세션에 저장 (복사·체크박스 조작 등 rerun에도 결과 유지)
                    results = calculate_purchase_date(df, st.session_state.exclude_dates_list, st.session_state.include_dates_list)
                    if results:
                        st.session_state['results'] = results
                        st.session_state['results_file'] = uploaded_file.name
                    else:
                        st.session_state.pop('results', None)
                        st.error("❌ 계산 중 오류가 발생했습니다.")

            # 계산 결과 표시
            results = st.session_state.get('results')
            if results:
                st.success("✅ 계산이 완료되었습니다!")

                # 발주서/월별 데이터 준비
                po_df = create_purchase_order_df(results, df)
                monthly_df = calculate_monthly_purchase(results, df)

                # 탭으로 결과 구분
                tab1, tab2, tab3, tab4, tab5 = st.tabs(
                    ["🧾 발주서", "📋 입고 계획", "📊 월별 구매량", "📅 입고 일정 요약", "🗓️ 캘린더 보기"])

                with tab1:
                    st.subheader("🧾 발주서 출력 (생성창 규격)")

                    if po_df.empty:
                        st.info("계산 기간 내 발주가 필요한 품목이 없습니다.")
                    else:
                        st.dataframe(po_df, use_container_width=True)

                        # 단가 미입력 안내
                        missing_price = int(po_df['총액(원)'].astype(str).eq('').sum())
                        if missing_price:
                            st.warning(f"⚠️ 단가 미입력으로 총액을 계산하지 못한 행이 {missing_price}건 있습니다.")

                        # 발주 총액 요약
                        total_sum = pd.to_numeric(po_df['총액(원)'], errors='coerce').sum()
                        sum_col1, sum_col2 = st.columns(2)
                        with sum_col1:
                            st.metric("발주 건수", len(po_df))
                        with sum_col2:
                            st.metric("발주 총액(원)", f"{total_sum:,.0f}")

                        st.markdown("**📋 아래 박스 우측 상단의 복사 버튼을 눌러 발주서 생성창에 그대로 붙여넣으세요.** (탭 구분 형식이라 Excel/그리드에 열이 맞춰 들어갑니다)")
                        include_header = st.checkbox("헤더(열 이름) 포함", value=True, key="po_header")
                        tsv_text = po_df.to_csv(sep='\t', index=False, header=include_header)
                        st.code(tsv_text, language=None)

                with tab2:
                    st.subheader("📋 입고 계획")
                    output_df = pd.DataFrame(results)
                    st.dataframe(output_df, use_container_width=True)

                    # 적용된 날짜 설정 표시
                    if st.session_state.exclude_dates_list or st.session_state.include_dates_list:
                        with st.expander("📅 적용된 날짜 설정 확인"):
                            col1, col2 = st.columns(2)
                            with col1:
                                if st.session_state.exclude_dates_list:
                                    st.write("**🚫 제외된 날짜들:**")
                                    for date in sorted(st.session_state.exclude_dates_list):
                                        weekday = ['월', '화', '수', '목', '금', '토', '일'][date.weekday()]
                                        st.write(f"• {date.strftime('%Y-%m-%d')} ({weekday})")
                                else:
                                    st.write("**🚫 제외된 날짜:** 없음")

                            with col2:
                                if st.session_state.include_dates_list:
                                    st.write("**✅ 포함된 날짜들:**")
                                    for date in sorted(st.session_state.include_dates_list):
                                        weekday = ['월', '화', '수', '목', '금', '토', '일'][date.weekday()]
                                        is_weekend_day = date.weekday() >= 5
                                        weekend_text = " 🟡" if is_weekend_day else " 🔵"
                                        st.write(f"• {date.strftime('%Y-%m-%d')} ({weekday}){weekend_text}")
                                else:
                                    st.write("**✅ 포함된 날짜:** 없음")

                    # 통계 정보
                    purchase_needed = len([r for r in results if r['입고 필요일'] != '입고 필요 없음'])
                    no_purchase = len(results) - purchase_needed

                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("입고 필요 품목", purchase_needed)
                    with col2:
                        st.metric("입고 불필요 품목", no_purchase)
                    with col3:
                        if len(results) > 0:
                            ratio = (purchase_needed / len(results)) * 100
                            st.metric("입고 필요 비율", f"{ratio:.1f}%")

                with tab3:
                    st.subheader("📊 월별 구매량")
                    if not monthly_df.empty:
                        st.dataframe(monthly_df, use_container_width=True)
                    else:
                        st.info("계산 기간 내 구매가 필요한 품목이 없습니다.")

                with tab4:
                    st.subheader("📅 입고 일정 요약")

                    # 입고 일정 달력 형태로 표시
                    all_purchase_dates = []
                    for item in results:
                        if item['입고 필요일'] != '입고 필요 없음':
                            dates = item['입고 필요일'].split(', ')
                            for date in dates:
                                all_purchase_dates.append({
                                    '날짜': date,
                                    '품목명': item['품목명'],
                                    '품목코드': item['품목코드']
                                })

                    if all_purchase_dates:
                        schedule_df = pd.DataFrame(all_purchase_dates)
                        schedule_df['날짜'] = pd.to_datetime(schedule_df['날짜'])
                        schedule_df = schedule_df.sort_values('날짜')
                        schedule_df['날짜'] = schedule_df['날짜'].dt.strftime('%Y-%m-%d')

                        st.dataframe(schedule_df, use_container_width=True)

                        # 날짜별 입고 품목 수
                        date_counts = schedule_df['날짜'].value_counts().sort_index()
                        st.bar_chart(date_counts)
                    else:
                        st.info("계산 기간 내 입고 일정이 없습니다.")

                with tab5:
                    st.subheader("🗓️ 입고 일정 캘린더")
                    display_calendar(results)

                # 다운로드 버튼
                st.markdown("---")
                st.subheader("💾 결과 다운로드")

                # Excel 파일 생성
                try:
                    excel_file = create_excel_file(results, monthly_df, po_df)

                    if excel_file is not None:
                        st.download_button(
                            label="📥 Excel 파일 다운로드 (발주서 포함)",
                            data=excel_file,
                            file_name=f"입고계획_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
                    else:
                        st.error("Excel 파일 생성에 실패했습니다.")
                except Exception as e:
                    st.error(f"다운로드 파일 생성 중 오류: {str(e)}")
                    # 대안: CSV 파일 다운로드 제공
                    st.info("대신 CSV 파일로 다운로드하시겠습니까?")
                    csv = po_df.to_csv(index=False, encoding='utf-8-sig')
                    st.download_button(
                        label="📄 CSV 파일 다운로드 (발주서)",
                        data=csv,
                        file_name=f"발주서_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                        mime="text/csv",
                        use_container_width=True
                    )

        except Exception as e:
            st.error(f"❌ 파일 처리 중 오류가 발생했습니다: {str(e)}")

    else:
        st.info("👆 Excel 파일을 업로드해주세요.")

        # 샘플 데이터 표시
        st.subheader("📋 샘플 데이터 형식")
        sample_data = {
            '원료코드명': ['1010101', '1010111'],
            '원료명': ['닭고기 MDCM', '닭고기가수분해단백질분말'],
            '현재 재고': [100, 50],
            '최소 사용량': [5, 3],
            '최대 사용량': [10, 8],
            '계산 일자': [30, 30],
            '안전 재고': [20, 15],
            '1회 구매량': [100, 80],
            '제외 날짜': ['2026-07-15', ''],
            '포함 날짜': ['', '2026-07-18'],
            '부가세 구분': ['면세', '과세'],
            '제품 구분': ['원료', '원료'],
            '업체명': ['주식회사 도명트레이딩', '(주)가온트레이딩'],
            '단량(KG)': [20, 25],
            '단가(원)': [3500, 12000],
            '담당자 이메일': ['dohmyung2022@naver.com', 'donghyun0910@gaontrading.com']
        }
        sample_df = pd.DataFrame(sample_data)
        st.dataframe(sample_df, use_container_width=True)

if __name__ == "__main__":
    main()
