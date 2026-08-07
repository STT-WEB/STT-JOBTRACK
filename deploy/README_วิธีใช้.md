# อัปเดตระบบ JOBTRACK แบบปุ่มเดียว (clasp)

## วิธีใช้ (ประจำวัน)
ดับเบิลคลิกไฟล์ **`อัปเดตระบบ.bat`** (อยู่นอกโฟลเดอร์ deploy) → รอจนขึ้น **DONE!** → ปิดแท็บแอปเก่า เปิดใหม่ กด `Ctrl+Shift+R`

`.bat` ทำให้อัตโนมัติ 4 สเต็ป:
1. ก๊อปโค้ด backend ล่าสุดจาก `src/` เข้า `deploy/`
2. Backup ขึ้น GitHub (ข้ามให้เองถ้ายังไม่ได้ตั้ง git)
3. `clasp push` — อัปโค้ดขึ้น Apps Script
4. `clasp deploy` — ปล่อยเวอร์ชันใหม่ที่ **URL เดิม** (ลิงก์ไม่เปลี่ยน)

## ครั้งแรกครั้งเดียว
เครื่องนี้ใช้ NOVA Purchase Hub อยู่แล้ว → มี Node + clasp + login ครบ ใช้ได้เลย
ถ้าเจอ error `clasp login` หมดอายุ → เปิด Command Prompt พิมพ์ `clasp login` ครั้งเดียว

## ค่าที่ตั้งไว้
- Script ID: `12ry2hbSjgteJambDLRoKT1oVxsPyHCwjN4sl_YGnsYeDIbSkVd_wcgA9`
- Deployment ID: `AKfycbyqyUP-2oE2PM_AzihyIq0_dMuZol4_lmTeRGFGeZ_ZMfc2gJTcje5rActFXW_OggfE`

## หมายเหตุ (2 ส่วนของ JOBTRACK)
- **Backend (Apps Script)** = ตัวนี้อัปด้วย `.bat` ✅
- **Frontend (หน้าเว็บ job_checkin_app.html บน GitHub Pages)** = ยังต้องอัปแยก
  ถ้าอยากให้ `.bat` อัปหน้าเว็บด้วยในปุ่มเดียว → บอก Candy ว่าโฟลเดอร์ GitHub Pages ของ JOBTRACK อยู่ที่ไหน จะต่อ git ให้
