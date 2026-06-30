import pandas as pd
from sqlalchemy.orm import Session
from models import ARInvoice, ExchangeRate, Customer

# 導入 Excel 檔案並更新 ARInvoice 表格
def import_ar_excel(file_bytes, tenant_id: int, db: Session):
    # 讀取 Excel 檔案
    df = pd.read_excel(file_bytes, engine="openpyxl")
    for index, row in df.iterrows():
        invoice_currency = row['Invoice Currency']
        invoice_amt = row['Invoice Amt']
        rate = None

        # 查詢 ExchangeRate 表以獲取匯率
        exchange_rate = db.query(ExchangeRate).filter(
            ExchangeRate.tenant_id == tenant_id,
            ExchangeRate.from_currency == invoice_currency,
            ExchangeRate.to_currency == 'HKD'
        ).order_by(ExchangeRate.updated_at.desc()).first()

        if exchange_rate:
            rate = exchange_rate.rate
        elif invoice_currency == 'HKD':
            rate = 1

        # 若無有效匯率，則標記為未設定
        amt_in_hkd = invoice_amt * rate if rate else None
        if amt_in_hkd is None:
            remark = row['Remark'] + ' [未設定匯率提示]'
        else:
            remark = row['Remark']

        # 自動檢查並建立 CustomerSupplier 紀錄
        # ...

        # 將記錄添加到 ARInvoice 表中
        ar_invoice = ARInvoice(
            tenant_id=tenant_id,
            cust_po=row['Customer PO'],
            invoice_number=row['Invoice No'],
            currency=invoice_currency,
            invoice_amount=invoice_amt,
            amount_hkd=amt_in_hkd,
            invoice_date=row['Invoice Date'],
            due_date=row['Due Date'],
            po_type=row['PO Type'],
            remark=remark,
            balance=invoice_amt,
            status='open'
        )
        db.add(ar_invoice)
    db.commit()