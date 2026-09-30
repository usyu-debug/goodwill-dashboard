from fastapi import FastAPI, UploadFile, File
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
import pandas as pd
import json
import glob
import os
import time
import re
import shutil

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

cached_data = None
last_file_mod_time = None

def get_latest_excel_file():
    files = glob.glob("*기증데이터*.xlsx")
    return max(files, key=os.path.getmtime) if files else None

def load_historical_data():
    file_name = "기증량분석 고정데이터_2023-2025년 요약본_2.xlsx"
    empty_df = pd.DataFrame(columns=['기증방법', '연도', '월', '모점포구분', '물품수량'])
    try:
        if not os.path.exists(file_name): return empty_df
        xls = pd.ExcelFile(file_name)
        records = []
        for sheet in xls.sheet_names:
            method = None
            if '고정데이터6' in sheet: method = '전체합계'
            elif '고정데이터7' in sheet or '기증센터 기증량2' in sheet: method = '직접기증'
            elif '고정데이터8' in sheet or '방문수거' in sheet: method = '방문수거'
            elif '고정데이터9' in sheet or '택배수거' in sheet: method = '택배수거'
            elif '고정데이터10' in sheet or '부서제출' in sheet: method = '부서제출'
            elif '고정데이터11' in sheet or '기업기증' in sheet: method = '기업기증'
            
            if method:
                df_sheet = pd.read_excel(xls, sheet_name=sheet, header=None)
                for i in range(len(df_sheet)):
                    if df_sheet.iloc[i, 1] == '월 구분':
                        branches = [(df_sheet.iloc[i, col_idx], col_idx) for col_idx in range(2, len(df_sheet.columns), 3) if pd.notna(df_sheet.iloc[i, col_idx]) and '총계' not in str(df_sheet.iloc[i, col_idx])]
                        row = i + 2
                        while row < len(df_sheet) and pd.notna(df_sheet.iloc[row, 1]) and str(df_sheet.iloc[row, 1]).isdigit():
                            month = int(df_sheet.iloc[row, 1])
                            for branch, col_idx in branches:
                                for year_offset, year_val in enumerate([2023, 2024, 2025]):
                                    if col_idx + year_offset < len(df_sheet.columns):
                                        qty = df_sheet.iloc[row, col_idx + year_offset]
                                        if pd.notna(qty) and str(qty).replace('.', '', 1).isdigit() and float(qty) > 0:
                                            records.append({'기증방법': method, '연도': f"{year_val}년", '월': month, '모점포구분': branch, '물품수량': float(qty)})
                            row += 1
        if records:
            df_res = pd.DataFrame(records)
            df_res['모점포구분'] = df_res['모점포구분'].astype(str).str[:2]
            return df_res
        return empty_df
    except Exception as e:
        return empty_df

def load_and_calculate_data():
    global cached_data, last_file_mod_time
    latest_file = get_latest_excel_file()
    if not latest_file: return {"status": "error", "message": "기증데이터 엑셀 파일을 찾을 수 없습니다."}
        
    current_mod_time = os.path.getmtime(latest_file)
    if cached_data is not None and last_file_mod_time == current_mod_time: return cached_data
        
    print(f"🔄 데이터 연산 시작...")
    start_time = time.time()
    try:
        filename = os.path.basename(latest_file)
        date_str = "당월"
        
        match_explicit = re.search(r'(\d+월\s*\d+일)', filename)
        if match_explicit:
            date_str = match_explicit.group(1)
        else:
            match_full = re.search(r'(20\d{2})[\.\-\_ ]+(\d{1,2})[\.\-\_ ]+(\d{1,2})', filename)
            month, day = 0, 0
            if match_full: 
                month = int(match_full.group(2))
                day = int(match_full.group(3))
            else:
                match_short = re.search(r'(?<!\d)(\d{1,2})[\.\-\_ ]+(\d{1,2})(?!\d)', filename.replace("기증데이터", ""))
                if match_short:
                    month = int(match_short.group(1))
                    day = int(match_short.group(2))
                    
            if month > 0 and day > 0:
                if month > 12: month, day = day, month
                date_str = f"{month}월 {day}일"
            
        df = pd.read_excel(latest_file, sheet_name=0)
        df_hist = load_historical_data()
        
        df['모점포구분'] = df['모점포구분'].fillna("알수없음").astype(str).str[:2]
        if '기증 건수' not in df.columns and '기증건수' not in df.columns:
            if '기증번호' in df.columns: df['기증 건수'] = (~df.duplicated(subset=['기증번호'])).astype(int)
            elif '접수번호' in df.columns: df['기증 건수'] = (~df.duplicated(subset=['접수번호'])).astype(int)
            else: df['기증 건수'] = 1 
        elif '기증건수' in df.columns: df.rename(columns={'기증건수': '기증 건수'}, inplace=True)
             
        if df['물품수량'].dtype == object: df['물품수량'] = df['물품수량'].astype(str).str.replace(',', '')
        if df['총금액'].dtype == object: df['총금액'] = df['총금액'].astype(str).str.replace(',', '')
        df['물품수량'] = pd.to_numeric(df['물품수량'], errors='coerce').fillna(0)
        df['총금액'] = pd.to_numeric(df['총금액'], errors='coerce').fillna(0)

        if '기증연월' not in df.columns and '기증 연월' not in df.columns:
            date_cols = [col for col in df.columns if any(kw in col for kw in ['일자', '날짜', '기증일', '접수일'])]
            if date_cols: df['기증연월'] = pd.to_datetime(df[date_cols[0]], errors='coerce').dt.month
            else: return {"status": "error", "message": "엑셀 파일에서 날짜를 알 수 있는 열을 찾지 못했습니다."}
        elif '기증 연월' in df.columns: df.rename(columns={'기증 연월': '기증연월'}, inplace=True)
            
        df['기증연월'] = df['기증연월'].fillna(0).astype(int)
        current_month_num = df['기증연월'].max()
        
        curr_df = df[df['기증연월'] == current_month_num]
        prev_df = df[df['기증연월'] == current_month_num - 1]
        prev_y_df = df_hist[(df_hist['연도'] == '2025년') & (df_hist['월'] == current_month_num)] if not df_hist.empty else pd.DataFrame(columns=['기증방법', '모점포구분', '물품수량'])

        methods = ['전체합계', '방문수거', '직접기증', '택배수거', '부서제출', '기업기증']
        branches = ['전체'] + sorted(curr_df['모점포구분'].unique().tolist())
        
        # 💡 [추가] 0번 대시보드 (굿윌스토어 전체 기증 추이)
        methods_5 = ['방문수거', '직접기증', '택배수거', '부서제출', '기업기증']
        trend_all_26_methods = df[(df['기증연월'] < current_month_num) & (df['기증방법'].isin(methods_5))]
        trend_all_26_grouped = trend_all_26_methods.groupby(['기증방법', '기증연월'])['물품수량'].sum().reset_index()
        trend_all_26_grouped.rename(columns={'기증연월': '월'}, inplace=True)
        trend_all_26_grouped['연도'] = '2026년'

        if not df_hist.empty:
            trend_hist_methods = df_hist[df_hist['기증방법'].isin(methods_5)]
            trend_hist_grouped = trend_hist_methods.groupby(['기증방법', '연도', '월'])['물품수량'].sum().reset_index()
        else:
            trend_hist_grouped = pd.DataFrame(columns=['기증방법', '연도', '월', '물품수량'])

        trend_master_0 = pd.concat([trend_hist_grouped, trend_all_26_grouped], ignore_index=True)
        grouped_0 = trend_master_0.groupby(['연도', '월', '기증방법'])['물품수량'].sum().unstack(fill_value=0).reset_index()
        
        for m in methods_5:
            if m not in grouped_0.columns: grouped_0[m] = 0
                
        grouped_0['연도_num'] = grouped_0['연도'].str.replace('년', '').astype(int)
        grouped_0 = grouped_0.sort_values(['연도_num', '월'])
        
        dashboard_0_data = {
            "years": grouped_0['연도'].astype(str).tolist(),
            "months": grouped_0['월'].tolist(),
            "total": grouped_0[methods_5].sum(axis=1).tolist(),
            "methods": {m: grouped_0[m].tolist() for m in methods_5}
        }

        stats_by_branch = {}
        for branch in branches:
            branch_stats = []
            b_curr = curr_df if branch == '전체' else curr_df[curr_df['모점포구분'] == branch]
            b_prev = prev_df if branch == '전체' else prev_df[prev_df['모점포구분'] == branch]
            b_prev_y = prev_y_df if branch == '전체' else prev_y_df[prev_y_df['모점포구분'] == branch]
            
            for m in methods:
                if m == '전체합계':
                    c_qty = b_curr['물품수량'].sum() if not b_curr.empty else 0
                    p_qty = b_prev['물품수량'].sum() if not b_prev.empty else 0
                    py_qty = b_prev_y['물품수량'].sum() if not b_prev_y.empty else 0
                else:
                    c_qty = b_curr[b_curr['기증방법'] == m]['물품수량'].sum() if not b_curr.empty else 0
                    p_qty = b_prev[b_prev['기증방법'] == m]['물품수량'].sum() if not b_prev.empty else 0
                    py_qty = b_prev_y[b_prev_y['기증방법'] == m]['물품수량'].sum() if not b_prev_y.empty else 0

                mom = round(((c_qty - p_qty) / p_qty * 100), 1) if p_qty > 0 else 0
                yoy = round(((c_qty - py_qty) / py_qty * 100), 1) if py_qty > 0 else 0
                branch_stats.append({"method": m, "qty": int(c_qty), "mom": mom, "yoy": yoy})
            stats_by_branch[branch] = branch_stats

        pivot_count = pd.pivot_table(df, index='기증방법', columns='모점포구분', values='기증 건수', aggfunc='sum', fill_value=0)
        pivot_quantity = pd.pivot_table(df, index='기증방법', columns='모점포구분', values='물품수량', aggfunc='sum', fill_value=0)
        pivot_amount = pd.pivot_table(df, index='기증방법', columns='모점포구분', values='총금액', aggfunc='sum', fill_value=0)
        pivot_count = pivot_count[pivot_count.sum().sort_values(ascending=False).index]
        pivot_quantity = pivot_quantity[pivot_quantity.sum().sort_values(ascending=False).index]
        pivot_amount = pivot_amount[pivot_amount.sum().sort_values(ascending=False).index]
        
        chart_data = df.groupby(['모점포구분', '기증방법'])['물품수량'].sum().reset_index()
        branch_order_chart = chart_data.groupby('모점포구분')['물품수량'].sum().sort_values(ascending=False).index.tolist()
        
        corp_df = df[df['기증방법'] == '기업기증']
        dashboard_3_data = {"overall": None, "items": {}}
        if not corp_df.empty:
            corp_grouped = corp_df.groupby('모점포구분')[['기증 건수', '물품수량', '총금액']].sum().sort_values(by='물품수량', ascending=False)
            dashboard_3_data["overall"] = {"branches": corp_grouped.index.tolist(), "count": corp_grouped['기증 건수'].tolist(), "quantity": corp_grouped['물품수량'].tolist(), "amount": corp_grouped['총금액'].tolist()}
            for item in corp_df['물품명'].dropna().unique().tolist():
                item_df = corp_df[corp_df['물품명'] == item].groupby('모점포구분')[['물품수량', '총금액']].sum().sort_values(by='물품수량', ascending=False)
                dashboard_3_data["items"][item] = {"branches": item_df.index.tolist(), "quantity": item_df['물품수량'].tolist(), "amount": item_df['총금액'].tolist()}

        trend_26 = df[df['기증방법'] == '직접기증'].groupby('기증연월')['물품수량'].sum().reset_index()
        trend_26.rename(columns={'기증연월': '월'}, inplace=True)
        trend_26 = trend_26[trend_26['월'] < current_month_num]
        trend_26['연도'] = '2026년'
        hist_direct = df_hist[df_hist['기증방법'] == '직접기증'] if not df_hist.empty else pd.DataFrame(columns=['연도', '월', '물품수량'])
        trend_hist = hist_direct.groupby(['연도', '월'])['물품수량'].sum().reset_index()
        trend_all = pd.concat([trend_hist, trend_26[['연도', '월', '물품수량']]], ignore_index=True).sort_values(by=['연도', '월'])
        dashboard_4_data = {"years": trend_all['연도'].astype(str).tolist(), "months": trend_all['월'].tolist(), "quantity": trend_all['물품수량'].tolist()}

        jan_25_stores = hist_direct[(hist_direct['연도'] == '2025년') & (hist_direct['월'] == 1) & (hist_direct['물품수량'] > 0)]['모점포구분'].unique()
        sum_25_ytd = hist_direct[(hist_direct['연도'] == '2025년') & (hist_direct['월'] < current_month_num)].groupby('모점포구분')['물품수량'].sum()
        sum_26_branch = df[(df['기증방법'] == '직접기증') & (df['기증연월'] < current_month_num)].groupby('모점포구분')['물품수량'].sum()
        
        growth_df = pd.DataFrame({'2025_YTD': sum_25_ytd, '2026': sum_26_branch}).fillna(0)
        growth_df = growth_df[growth_df.index.isin(jan_25_stores)]
        growth_df['growth'] = (growth_df['2026'] - growth_df['2025_YTD']) / growth_df['2025_YTD'].replace(0, 1) * 100
        top5_branches = growth_df.sort_values('growth', ascending=False).head(5).index.tolist()
        
        top5_data = []
        for br in top5_branches:
            br_hist = hist_direct[(hist_direct['모점포구분'] == br) & (hist_direct['연도'].str.contains('2025'))]
            br_curr = df[(df['기증방법'] == '직접기증') & (df['모점포구분'] == br)]
            br_curr = br_curr[br_curr['기증연월'] < current_month_num]
            br_curr_grp = br_curr.groupby('기증연월')['물품수량'].sum().reset_index()
            br_curr_grp.rename(columns={'기증연월': '월'}, inplace=True)
            br_curr_grp['연도'] = '2026년'
            br_trend = pd.concat([br_hist[['연도', '월', '물품수량']], br_curr_grp], ignore_index=True).sort_values(['연도', '월'])
            top5_data.append({"branch": br, "growth": round(growth_df.loc[br, 'growth'], 1), "years": br_trend['연도'].astype(str).tolist(), "months": br_trend['월'].tolist(), "quantity": br_trend['물품수량'].tolist()})
            
        all_trend_26 = df.groupby(['기증방법', '모점포구분', '기증연월'])['물품수량'].sum().reset_index()
        all_trend_26.rename(columns={'기증연월': '월'}, inplace=True)
        all_trend_26 = all_trend_26[all_trend_26['월'] < current_month_num]
        all_trend_26['연도'] = '2026년'
        all_trend_hist = df_hist.groupby(['기증방법', '모점포구분', '연도', '월'])['물품수량'].sum().reset_index() if not df_hist.empty else pd.DataFrame()
        all_trend_data = pd.concat([all_trend_hist, all_trend_26], ignore_index=True)
        dashboard_5_data = {"top5": top5_data, "all_trends": all_trend_data.to_dict(orient='records')}

        item_df = df.groupby('물품명')[['기증 건수', '물품수량', '총금액']].sum().reset_index().sort_values('물품수량', ascending=False)
        item_method_df = df.groupby(['기증방법', '물품명'])['물품수량'].sum().reset_index()
        method_items = {}
        for m in df['기증방법'].unique():
            m_data = item_method_df[item_method_df['기증방법'] == m].sort_values('물품수량', ascending=False)
            if not m_data.empty: method_items[m] = m_data.to_dict(orient='records')
        dashboard_6_data = {"overall": item_df.to_dict(orient='records'), "by_method": method_items}

        cached_data = {
            "status": "success",
            "file_name": os.path.basename(latest_file),
            "dashboard_0": dashboard_0_data,
            "dashboard_1": {"date_str": date_str, "stats_by_branch": stats_by_branch, "branch_list": branches},
            "dashboard_2": {
                "count": json.loads(pivot_count.to_json(orient='index')), "quantity": json.loads(pivot_quantity.to_json(orient='index')), "amount": json.loads(pivot_amount.to_json(orient='index')), 
                "ordered_branches_count": pivot_count.columns.tolist(), "ordered_branches_quantity": pivot_quantity.columns.tolist(), "ordered_branches_amount": pivot_amount.columns.tolist(),
                "chart": chart_data.to_dict(orient='records'), "chart_branch_order": branch_order_chart
            },
            "dashboard_3": dashboard_3_data, "dashboard_4": dashboard_4_data, "dashboard_5": dashboard_5_data, "dashboard_6": dashboard_6_data
        }
        last_file_mod_time = current_mod_time
        print(f"✅ 연산 완료! 소요 시간: {time.time() - start_time:.2f}초")
        return cached_data

    except Exception as e:
        import traceback
        traceback.print_exc()
        return {"status": "error", "message": f"데이터 처리 중 오류 발생: {str(e)}"}

@app.get("/api/dashboard-data")
def get_dashboard_data():
    return load_and_calculate_data()

@app.post("/api/upload-data")
async def upload_data(file: UploadFile = File(...)):
    try:
        safe_filename = file.filename
        if not safe_filename.endswith('.xlsx'): return {"status": "error", "message": "엑셀(.xlsx) 파일만 업로드 가능합니다."}
        file_path = os.path.join(os.getcwd(), safe_filename)
        with open(file_path, "wb") as buffer: shutil.copyfileobj(file.file, buffer)
        return {"status": "success", "message": "파일 업로드 및 서버 반영 완료"}
    except Exception as e:
        return {"status": "error", "message": f"업로드 실패: {str(e)}"}
@app.get("/")
def read_index():
    return FileResponse("index.html")

@app.get("/dashboard.html")
def read_dashboard():
    return FileResponse("dashboard.html")