# JOBTRACK — วิธีเชื่อม Web App กับ Google Sheets
## ทำแค่ 3 ขั้นตอน ใช้เวลาประมาณ 10 นาที

---

## ขั้นตอนที่ 1 — สร้าง Google Sheet

1. ไปที่ https://sheets.google.com
2. กด "+ สร้างสเปรดชีตใหม่"
3. ตั้งชื่อว่า **JOBTRACK_Database**
4. Copy **Sheet ID** จาก URL
   - URL ตัวอย่าง: https://docs.google.com/spreadsheets/d/**1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgVE2upms**/edit
   - ส่วนที่ต้องเอาคือ: **1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgVE2upms**

---

## ขั้นตอนที่ 2 — ตั้งค่า Google Apps Script

1. ใน Google Sheet กด **Extensions → Apps Script**
2. ลบโค้ดเดิมทั้งหมดที่มีอยู่
3. Copy โค้ดจากไฟล์ **jobtrack_apps_script.gs** วางลงไปทั้งหมด
4. แก้ไขบรรทัดที่ 10:
   ```
   const SHEET_ID = 'วางSheet_ID_ที่ Copy มาตรงนี้';
   ```
5. กด **Save** (Ctrl+S)
6. กดปุ่ม **▶ Run** เลือก function **testInsert** → กด **Allow** (ให้สิทธิ์ครั้งแรก)
   - ถ้าสำเร็จจะมี Sheet Tab ชื่อ **Job_Log** และมีแถวข้อมูลทดสอบ 1 แถว

---

## ขั้นตอนที่ 3 — Deploy เป็น Web App และเชื่อม

1. ใน Apps Script กด **Deploy → New deployment**
2. ตั้งค่าดังนี้:
   - Type: **Web app**
   - Description: JOBTRACK v1
   - Execute as: **Me**
   - Who has access: **Anyone** ← สำคัญมาก
3. กด **Deploy** → Copy **Web App URL**
   - ตัวอย่าง: https://script.google.com/macros/s/XXXX/exec
4. เปิดไฟล์ **job_checkin_app.html** ด้วย Text Editor
5. แก้บรรทัดที่ 7:
   ```
   const APPS_SCRIPT_URL = 'วาง Web App URL ตรงนี้';
   ```
6. Save ไฟล์

---

## ทดสอบ

1. เปิดไฟล์ **job_checkin_app.html** ในมือถือหรือ Browser
2. เลือกสถานะ + Process กด **บันทึก**
3. เปิด Google Sheet ดูที่ Tab **Job_Log**
4. ข้อมูลควรปรากฏทันที ✅

---

## โครงสร้าง Google Sheet ที่จะได้

| Timestamp | วันที่ | เวลา | Job ID | ชื่องาน | Zone | รหัสพนักงาน | ชื่อพนักงาน | แผนก | สถานะ | Process | ภาษา |
|-----------|--------|------|--------|---------|------|-------------|-------------|------|-------|---------|------|
| 27/4/2569 09:32 | 27/4/2569 | 09:32 | T-0042 | แท็งค์น้ำมัน #4 | ลาน B · Zone 3 | EMP001 | สมชาย รักงาน | ฝ่ายผลิต | Check In | Fabrication | th |

---

## หมายเหตุ

- ไฟล์ HTML เปิดได้บน Browser โดยตรง ยังไม่ต้อง Host
- ถ้าต้องการให้เข้าผ่าน QR Code → Host ไฟล์บน GitHub Pages (ฟรี) หรือ Google Sites
- LINE Login เพิ่มได้ทีหลัง ตอนนี้ระบุพนักงานด้วย Hardcode ได้ก่อน
