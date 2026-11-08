from fastapi import FastAPI, UploadFile, File, HTTPException, Query, Form
from fastapi.middleware.cors import CORSMiddleware
import pandas as pd
import sqlite3
import io
import re

app = FastAPI(title="Tax Invoice Management API")

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
            is_tax_exempt INTEGER DEFAULT 0,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    try:
        cursor.execute("ALTER TABLE invoices ADD COLUMN is_tax_exempt INTEGER DEFAULT 0")
    except sqlite3.OperationalError:
        pass # 이미 컬럼이 존재하는 경우
        
    conn.commit()
    conn.close()

init_db()

def extract_field(row, target_keywords):
    for col_name, val in row.items():
        clean_col = re.sub(r'[\s\n\(\)]', '', str(col_name))
        for key in target_keywords:
            clean_key = re.sub(r'[\s\n\(\)]', '', key)
            if clean_key in clean_col and pd.notna(val):
                return str(val).strip()
    return ""

@app.post("/api/upload")
async def upload_excel(
    file: UploadFile = File(...),
    is_tax_exempt: bool = Form(False)
):
    if not file.filename.endswith(('.xlsx', '.xls')):
        raise HTTPException(status_code=400, detail="엑셀(.xlsx, .xls) 파일만 업로드 가능합니다.")

    contents = await file.read()
    
    try:
        raw_df = pd.read_excel(io.BytesIO(contents), header=None)
        
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

        tax_exempt_flag = 1 if is_tax_exempt else 0

        for _, row in df.iterrows():
            date_val = extract_field(row, ['작성일자', '발급일자'])
            if not date_val:
                continue

            app_num = extract_field(row, ['승인번호'])
            name_val = extract_field(row, ['공급자상호', '공급자상호성명', '상호', '공급받는자상호'])
            biz_num = extract_field(row, ['공급자사업자등록번호', '공급자등록번호', '사업자등록번호', '공급받는자사업자등록번호'])
            
            supply_raw = extract_field(row, ['공급가액', '공급가']).replace(',', '').split('.')[0]
            tax_raw = extract_field(row, ['세액', '부가세']).replace(',', '').split('.')[0]
            
            supply_val = int(supply_raw) if supply_raw.lstrip('-').isdigit() else 0
            tax_val = int(tax_raw) if tax_raw.lstrip('-').isdigit() else 0
            
            # 면세계산서 체크 시 세액은 0원 처리
            if tax_exempt_flag:
                tax_val = 0
                
            total_val = supply_val + tax_val

            if not app_num or app_num in ('nan', 'None', ''):
                app_num = f"AUTO-{date_val}-{name_val}-{supply_val}"

            try:
                cursor.execute('''
                    INSERT INTO invoices (approval_num, issue_date, supplier_biz_num, supplier_name, supply_value, tax_amount, total_amount, is_tax_exempt)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                ''', (app_num, date_val, biz_num, name_val, supply_val, tax_val, total_val, tax_exempt_flag))
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