from fastapi import FastAPI, UploadFile, File, HTTPException, Query
from fastapi.middleware.cors import CORSMiddleware
import pandas as pd
import sqlite3
import io
import re

app = FastAPI(title="Tax Invoice Management API")

# CORS setup for Frontend integration
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

DB_PATH = "invoices.db"

def init_db():
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS invoices (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            approval_num TEXT UNIQUE,
            issue_date TEXT,
            supplier_biz_num TEXT,
            supplier_name TEXT,
            supply_value INTEGER,
            tax_amount INTEGER,
            total_amount INTEGER,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    conn.commit()
    conn.close()

init_db()

def extract_field(row, target_keywords):
    """
    row(행 데이터)에서 띄어쓰기나 특수문자를 제거한 후 
    target_keywords 중 하나라도 포함된 열의 값을 찾아 반환합니다.
    """
    for col_name, val in row.items():
        clean_col = re.sub(r'[\s\n\(\)]', '', str(col_name))
        for key in target_keywords:
            clean_key = re.sub(r'[\s\n\(\)]', '', key)
            if clean_key in clean_col and pd.notna(val):
                return str(val).strip()
    return ""

@app.post("/api/upload")
async def upload_excel(file: UploadFile = File(...)):
    if not file.filename.endswith(('.xlsx', '.xls')):
        raise HTTPException(status_code=400, detail="엑셀(.xlsx, .xls) 파일만 업로드 가능합니다.")

    contents = await file.read()
    
    try:
        raw_df = pd.read_excel(io.BytesIO(contents), header=None)
        
        # '작성일자' 또는 '승인번호'가 있는 헤더 행 찾기
        header_idx = None
        for i, row in raw_df.iterrows():
            row_str = " ".join([str(val) for val in row.values])
            if "작성일자" in row_str or "승인번호" in row_str:
                header_idx = i
                break
                
        if header_idx is None:
            raise HTTPException(status_code=400, detail="올바른 세금계산서 엑셀 양식이 아닙니다.")

        df = pd.read_excel(io.BytesIO(contents), skiprows=header_idx)
        
        inserted_cnt = 0
        duplicate_cnt = 0

        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()

        for _, row in df.iterrows():
            # 유연한 열이름 매핑 탐색
            date_val = extract_field(row, ['작성일자', '발급일자'])
            if not date_val:
                continue # 작성일자가 없는 합계/빈 행은 제외

            app_num = extract_field(row, ['승인번호'])
            
            # 공급자 상호 (매출일 경우 공급받는자 상호까지 탐색)
            name_val = extract_field(row, ['공급자상호', '공급자상호성명', '상호', '공급받는자상호'])
            
            # 사업자등록번호
            biz_num = extract_field(row, ['공급자사업자등록번호', '공급자등록번호', '사업자등록번호', '공급받는자사업자등록번호'])
            
            # 공급가액 및 세액 추출
            supply_raw = extract_field(row, ['공급가액', '공급가']).replace(',', '').split('.')[0]
            tax_raw = extract_field(row, ['세액', '부가세']).replace(',', '').split('.')[0]
            
            supply_val = int(supply_raw) if supply_raw.lstrip('-').isdigit() else 0
            tax_val = int(tax_raw) if tax_raw.lstrip('-').isdigit() else 0
            total_val = supply_val + tax_val

            if not app_num or app_num in ('nan', 'None', ''):
                app_num = f"AUTO-{date_val}-{name_val}-{supply_val}"

            try:
                cursor.execute('''
                    INSERT INTO invoices (approval_num, issue_date, supplier_biz_num, supplier_name, supply_value, tax_amount, total_amount)
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                ''', (app_num, date_val, biz_num, name_val, supply_val, tax_val, total_val))
                inserted_cnt += 1
            except sqlite3.IntegrityError:
                duplicate_cnt += 1

        conn.commit()
        conn.close()

        return {
            "message": "파일 업로드 완료",
            "inserted": inserted_cnt,
            "duplicates": duplicate_cnt
        }

    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"엑셀 파싱 중 오류 발생: {str(e)}")

@app.get("/api/invoices")
def get_invoices(search: str = Query(None)):
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()

    if search:
        cursor.execute('''
            SELECT * FROM invoices 
            WHERE supplier_name LIKE ? OR supplier_biz_num LIKE ? OR approval_num LIKE ?
            ORDER BY issue_date DESC
        ''', (f"%{search}%", f"%{search}%", f"%{search}%"))
    else:
        cursor.execute('SELECT * FROM invoices ORDER BY issue_date DESC')
        
    rows = cursor.fetchall()
    conn.close()

    items = [dict(row) for row in rows]
    
    total_supply = sum(item['supply_value'] for item in items)
    total_tax = sum(item['tax_amount'] for item in items)
    total_amount = sum(item['total_amount'] for item in items)

    return {
        "summary": {
            "total_count": len(items),
            "total_supply": total_supply,
            "total_tax": total_tax,
            "total_amount": total_amount
        },
        "items": items
    }

@app.delete("/api/invoices")
def clear_invoices():
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute('DELETE FROM invoices')
    conn.commit()
    conn.close()
    return {"message": "전체 데이터 삭제 완료"}
