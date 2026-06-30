from datetime import datetime
from io import BytesIO
from pathlib import Path

from fastapi import Depends, FastAPI, File, HTTPException, UploadFile, status
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from pydantic import BaseModel
from sqlalchemy.orm import Session

from auth import create_access_token, get_current_user
from database import Base, engine, get_db
from models import ExchangeRate
from services import import_ar_excel

# 建立資料表（本機 SQLite 可直接使用）
Base.metadata.create_all(bind=engine)

app = FastAPI(title="ERP SaaS API")
BASE_DIR = Path(__file__).resolve().parent

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


class LoginRequest(BaseModel):
    email: str
    password: str


class ExchangeRateRequest(BaseModel):
    from_currency: str
    to_currency: str = "HKD"
    rate: float


@app.get("/")
def root():
    return {"status": "ok", "message": "ERP API 運行中"}


@app.get("/ui")
def serve_ui():
    """前端展示頁"""
    index_path = BASE_DIR / "index.html"
    if not index_path.exists():
        raise HTTPException(status_code=404, detail="找不到 index.html")
    return FileResponse(index_path)


@app.post("/api/login")
def login(req: LoginRequest, db: Session = Depends(get_db)):
    """測試登入：驗證成功後回傳 JWT（含 tenant_id / user_id / role）"""
    # TODO: 改為查詢 User 表並驗證密碼
    _ = db
    if not req.email or not req.password:
        raise HTTPException(status_code=status.HTTP_400_BAD_REQUEST, detail="請提供 email 與 password")

    tenant_id = 1
    user_id = 1
    role = "admin"
    token = create_access_token(
        {"tenant_id": tenant_id, "user_id": user_id, "role": role, "sub": req.email}
    )
    return {"access_token": token, "token_type": "bearer"}


@app.get("/api/ar/summary")
def ar_summary(
    current_user: dict = Depends(get_current_user),
    db: Session = Depends(get_db),
):
    """依租戶隔離查詢 AR 摘要（測試用）"""
    _ = db
    return {
        "tenant_id": current_user["tenant_id"],
        "total_ar": 0,
        "license_total": 0,
        "product_total": 0,
    }


@app.post("/api/settings/exchange-rates")
def update_exchange_rate(
    payload: ExchangeRateRequest,
    current_user: dict = Depends(get_current_user),
    db: Session = Depends(get_db),
):
    """登入後設定或更新匯率"""
    tenant_id = current_user["tenant_id"]
    record = (
        db.query(ExchangeRate)
        .filter(
            ExchangeRate.tenant_id == tenant_id,
            ExchangeRate.from_currency == payload.from_currency,
            ExchangeRate.to_currency == payload.to_currency,
        )
        .first()
    )

    if record:
        record.rate = payload.rate
        record.updated_at = datetime.utcnow()
    else:
        record = ExchangeRate(
            tenant_id=tenant_id,
            from_currency=payload.from_currency,
            to_currency=payload.to_currency,
            rate=payload.rate,
            updated_at=datetime.utcnow(),
        )
        db.add(record)

    db.commit()
    db.refresh(record)
    return {
        "message": "匯率已更新",
        "from_currency": record.from_currency,
        "to_currency": record.to_currency,
        "rate": float(record.rate),
    }


@app.post("/api/ar/upload")
async def upload_ar_file(
    file: UploadFile = File(...),
    current_user: dict = Depends(get_current_user),
    db: Session = Depends(get_db),
):
    """上傳 AR Excel 並匯入"""
    if not file.filename:
        raise HTTPException(status_code=status.HTTP_400_BAD_REQUEST, detail="未提供檔案")

    content = await file.read()
    import_ar_excel(BytesIO(content), current_user["tenant_id"], db)
    return {"message": "檔案成功上傳", "filename": file.filename}
