from fastapi import FastAPI, UploadFile, File, HTTPException, Query, Form
from fastapi.middleware.cors import CORSMiddleware
import pandas as pd
import sqlite3
import io
import re
import os
import uuid

app = FastAPI(title="Tax Invoice Management API")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, "invoices.db")

def init_db():
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS invoices (
            approval_num TEXT PRIMARY KEY,
            issue_date TEXT,
            send_date TEXT,
            
            supplier_biz_num TEXT,
            supplier_sub_num TEXT,
            supplier_name TEXT,
            supplier_ceo TEXT,
            supplier_addr TEXT,
            supplier_biz_type TEXT,
            supplier_biz_item TEXT,
            supplier_email TEXT,
            
            receiver_biz_num TEXT,
            receiver_sub_num TEXT,
            receiver_name TEXT,
            receiver_ceo TEXT,
            receiver_addr TEXT,
            receiver_biz_type TEXT,
            receiver_biz_item TEXT,
            receiver_email TEXT,
            
            supply_value INTEGER,
            tax_amount INTEGER,
            total_amount INTEGER,
            is_tax_exempt INTEGER DEFAULT 0,
            
            item_date TEXT,
            item_name TEXT,
            item_spec TEXT,
            item_qty INTEGER,
            item_price INTEGER,
            remarks TEXT,
            claim_or_receipt TEXT DEFAULT '청구',
            
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    
    # 기존 DB 테이블 마이그레이션 (컬럼 자동 추가)
    columns_to_add = [
        ("send_date", "TEXT"), ("supplier_sub_num", "TEXT"), ("supplier_ceo", "TEXT"),
        ("supplier_addr", "TEXT"), ("supplier_biz_type", "TEXT"), ("supplier_biz_item", "TEXT"),
        ("supplier_email", "TEXT"), ("receiver_biz_num", "TEXT"), ("receiver_sub_num", "TEXT"),
        ("receiver_name", "TEXT"), ("receiver_ceo", "TEXT"), ("receiver_addr", "TEXT"),
        ("receiver_biz_type", "TEXT"), ("receiver_biz_item", "TEXT"), ("receiver_email", "TEXT"),
        ("is_tax_exempt", "INTEGER DEFAULT 0"), ("item_date", "TEXT"), ("item_name", "TEXT"),
        ("item_spec", "TEXT"), ("item_qty", "INTEGER"), ("item_price", "INTEGER"), ("remarks", "TEXT"),
        ("claim_or_receipt", "TEXT DEFAULT '청구'")
    ]
    for col_name, col_type in columns_to_add:
        try:
            cursor.execute(f"ALTER TABLE invoices ADD COLUMN {col_name} {col_type}")
        except sqlite3.OperationalError:
            pass

    conn.commit()
    conn.close()

init_db()

def parse_int_val(val):
    if pd.isna(val) or val is None:
        return 0
    clean = str(val).replace(',', '').split('.')[0].strip()
    return int(clean) if clean.lstrip('-').isdigit() else 0

def clean_str(val):
    if pd.isna(val) or val is None:
        return ""
    res = str(val).strip()
    return "" if res in ('nan', 'None') else res

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
            raise HTTPException(status_code=400, detail="올바른 국세청 세금계산서 엑셀 양식이 아닙니다.")

        df = pd.read_excel(io.BytesIO(contents), skiprows=header_idx)
        
        inserted_cnt = 0
        duplicate_cnt = 0

        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()

        tax_exempt_flag = 1 if is_tax_exempt else 0

        for _, row in df.iterrows():
            r = {str(k).strip(): v for k, v in row.items()}
            
            date_val = clean_str(r.get('작성일자', r.get('발급일자')))
            if not date_val:
                continue

            app_num = clean_str(r.get('승인번호'))
            send_date = clean_str(r.get('전송일자', r.get('발급일자')))
            
            # 공급자 정보
            supplier_biz_num = clean_str(r.get('공급자사업자등록번호', r.get('공급자 사업자등록번호')))
            supplier_sub_num = clean_str(r.get('종사업장번호', r.get('공급자종사업장번호')))
            supplier_name = clean_str(r.get('상호', r.get('공급자상호', r.get('공급자 상호'))))
            supplier_ceo = clean_str(r.get('대표자명', r.get('대표자 명', r.get('공급자대표자명'))))
            supplier_addr = clean_str(r.get('주소', r.get('공급자주소')))
            supplier_biz_type = clean_str(r.get('공급자업태', r.get('업태')))
            supplier_biz_item = clean_str(r.get('공급자종목', r.get('종목')))
            supplier_email = clean_str(r.get('공급자 이메일', r.get('공급자이메일')))

            # 공급받는자 정보
            receiver_biz_num = clean_str(r.get('공급받는자사업자등록번호', r.get('공급받는자 사업자등록번호')))
            receiver_sub_num = clean_str(r.get('종사업장번호.1', r.get('공급받는자종사업장번호')))
            receiver_name = clean_str(r.get('상호.1', r.get('공급받는자상호', r.get('공급받는자 상호'))))
            receiver_ceo = clean_str(r.get('대표자명.1', r.get('공급받는자대표자명')))
            receiver_addr = clean_str(r.get('주소.1', r.get('공급받는자주소')))
            receiver_biz_type = clean_str(r.get('공급받는자업태'))
            receiver_biz_item = clean_str(r.get('공급받는자종목'))
            receiver_email = clean_str(r.get('공급받는자 이메일1', r.get('공급받는자 이메일2')))

            # 금액 및 품목 정보
            supply_val = parse_int_val(r.get('공급가액', r.get('공급가')))
            tax_val = 0 if tax_exempt_flag else parse_int_val(r.get('세액', r.get('부가세')))
            total_val = parse_int_val(r.get('합계금액', r.get('총금액')))
            if total_val == 0:
                total_val = supply_val + tax_val

            item_date = clean_str(r.get('품목일자', r.get('일자')))
            item_name = clean_str(r.get('품목명', r.get('품목')))
            item_spec = clean_str(r.get('품목규격', r.get('규격')))
            item_qty = parse_int_val(r.get('품목수량', r.get('수량')))
            item_price = parse_int_val(r.get('품목단가', r.get('단가')))
            remarks = clean_str(r.get('비고', r.get('품목비고')))

            # 영수/청구 구분 추출
            claim_or_receipt = clean_str(r.get('영수/청구 구분', r.get('영수/청구', r.get('영수구분'))))
            if not claim_or_receipt:
                claim_or_receipt = '청구'

            if not app_num:
                app_num = f"AUTO-{date_val}-{supplier_name}-{supply_val}-{uuid.uuid4().hex[:6]}"

            try:
                cursor.execute('''
                    INSERT INTO invoices (
                        approval_num, issue_date, send_date,
                        supplier_biz_num, supplier_sub_num, supplier_name, supplier_ceo, supplier_addr, supplier_biz_type, supplier_biz_item, supplier_email,
                        receiver_biz_num, receiver_sub_num, receiver_name, receiver_ceo, receiver_addr, receiver_biz_type, receiver_biz_item, receiver_email,
                        supply_value, tax_amount, total_amount, is_tax_exempt,
                        item_date, item_name, item_spec, item_qty, item_price, remarks, claim_or_receipt
                    )
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ''', (
                    app_num, date_val, send_date,
                    supplier_biz_num, supplier_sub_num, supplier_name, supplier_ceo, supplier_addr, supplier_biz_type, supplier_biz_item, supplier_email,
                    receiver_biz_num, receiver_sub_num, receiver_name, receiver_ceo, receiver_addr, receiver_biz_type, receiver_biz_item, receiver_email,
                    supply_val, tax_val, total_val, tax_exempt_flag,
                    item_date, item_name, item_spec, item_qty, item_price, remarks, claim_or_receipt
                ))
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
            WHERE supplier_name LIKE ? OR supplier_biz_num LIKE ? OR receiver_name LIKE ? OR receiver_biz_num LIKE ? OR approval_num LIKE ?
            ORDER BY issue_date DESC
        ''', (f"%{search}%", f"%{search}%", f"%{search}%", f"%{search}%", f"%{search}%"))
    else:
        cursor.execute('SELECT * FROM invoices ORDER BY issue_date DESC')
        
    rows = cursor.fetchall()
    conn.close()

    items = [dict(row) for row in rows]
    
    total_supply = sum(item['supply_value'] or 0 for item in items)
    total_tax = sum(item['tax_amount'] or 0 for item in items)
    total_amount = sum(item['total_amount'] or 0 for item in items)

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