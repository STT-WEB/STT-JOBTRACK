# JOBTRACK — Technical Overview (ฐานสำหรับ Develop ต่อ)

> ระบบเก็บ Job Cost ของ HR: พนักงานสแกน QR ของแต่ละงาน → Check In/Out → ระบบบันทึกชั่วโมงทำงาน (ปกติ/OT) ลง Google Sheet แยกตามรอบเงินเดือน
> เวอร์ชันปัจจุบัน: **Apps Script v3.2**

---

## 1. สถาปัตยกรรม (Architecture)

```
[QR Code ต่อ 1 งาน]
        │  (?job=&name=&type=)
        ▼
[Frontend: job_checkin_app.html]  ── host บน GitHub Pages
        │   stt-web.github.io/STT-JOBTRACK/
        │
        ├── LINE Login (ยืนยันตัวตนพนักงาน)  → callback.html
        ├── ลงทะเบียนพนักงานใหม่             → register.html / callback_register.html
        ├── อัปโหลดรูปงาน                    → imgbb API
        │
        ▼  (JSONP / POST → action-based dispatch)
[Backend: jobtrack_apps_script.gs]  ── Google Apps Script Web App
        │
        ▼
[Google Sheet: JOBTRACK_Database]
   ID = 1MYWORYN3sOjov3Gxv3UqCV1jRSxgxwGi1tRomFUGSr0
```

**หลักการสื่อสาร:** Frontend ส่ง object `{action: '...', ...}` ไปที่ Apps Script Web App
`doGet`/`doPost` → `dispatch()` แยกตาม `action` แล้วเรียกฟังก์ชันที่เกี่ยวข้อง

---

## 2. ไฟล์ในโปรเจกต์

| ไฟล์ | หน้าที่ |
|------|---------|
| `src/jobtrack_apps_script.gs` | Backend ทั้งหมด (496 บรรทัด) — logic, คำนวณชั่วโมง, จัดการ Sheet |
| `src/job_checkin_app.html` | แอปหลัก (1383 บรรทัด) — หน้า Check In/Out, multi-language, ถ่ายรูป |
| `src/register.html` | หน้าลงทะเบียนพนักงานใหม่ผูก LINE |
| `src/callback.html` | รับ callback จาก LINE Login (login flow) |
| `src/callback_register.html` | รับ callback จาก LINE Login (register flow) |
| `src/test_upload.html` | หน้าทดสอบอัปโหลดรูป |
| `docs/` | SOP, Setup Guide, Executive deck (อ้างอิง) |

---

## 3. Backend — Actions ที่รองรับ (dispatch)

| Action | ฟังก์ชัน | หน้าที่ |
|--------|----------|---------|
| `PING` | — | เช็คว่า Web App ทำงาน |
| `GET_JOB_INFO` | `getJobInfo()` | ดึงชื่องาน/ประเภทจาก `Job_List` ด้วย Job Code |
| `CHECK_STATUS` | `checkOpenJob()` | เช็คว่าพนักงานมีงานค้าง (Check In แต่ยังไม่ Out) |
| `CHECK_IN` | `doCheckIn()` | บันทึกเข้างาน (บล็อกถ้ามีงานค้าง) |
| `CHECK_OUT` | `doCheckOut()` | บันทึกออกงาน + คำนวณชั่วโมง |
| `UPDATE_PROCESS` | `doUpdateProcess()` | แก้ Process ของงานที่ค้างอยู่ |
| `SAVE_PHOTO_URLS` | `doSavePhotoUrls()` | บันทึก URL รูป (จาก imgbb) ลงคอลัมน์หมายเหตุ |
| `REGISTER_EMPLOYEE` | `registerEmployeeByCode()` | ผูก LINE ID กับรหัสพนักงาน |
| `LINE_CALLBACK` | `handleLineCallback()` | แลก code → token → profile → หาพนักงาน |

**ฟังก์ชัน admin (รันมือใน Apps Script):**
`importEmployeeData()` — ดึงพนักงานจาก `DATA EMPLOYEE` → `Employee_List`
`installTrigger()` — ติดตั้ง onEdit trigger สร้าง QR อัตโนมัติใน `Job_List`

---

## 4. Logic การคำนวณชั่วโมง (สำคัญที่สุด)

ค่าคงที่: `WORK_START_HOUR=8`, `OT_START_HOUR=17`, `WORK_HOURS=8`

- **แบ่งปกติ/OT ที่ 17:00** — ก่อน 17:00 = ปกติ, หลัง 17:00 = OT, คร่อม 17:00 = แบ่งสองส่วน
- **วันหยุด/นักขัตฤกษ์** → นับเป็น OT ทั้งหมด (x1.5 วันทำงาน / x3 วันหยุด)
- **ปัดเศษ:** ถ้าชั่วโมงรวมต่างจากชั่วโมงเต็มไม่เกิน 10 นาที → ปัดเป็นชั่วโมงเต็ม
- **เป้าหมาย 8 ชม./วัน:** `updateDailyHourAlert()` รวมชั่วโมงปกติทั้งวัน แล้วใส่สี + note: ✅ ครบ / ⚠️ ขาด / ℹ️ เกิน
- ประเภทวันได้จาก tab `ประเภทวันทำงาน` (ถ้าไม่เจอ default: อาทิตย์=หยุด, อื่น=ปกติ)

---

## 5. โครงสร้าง Google Sheet (Tabs)

| Tab | หน้าที่ |
|-----|---------|
| `ประเภทวันทำงาน` | ปฏิทินกำหนดวันทำงาน/หยุด/นักขัตฤกษ์ (ต้อง maintain ล่วงหน้า) |
| `DATA EMPLOYEE` | ข้อมูลพนักงานต้นทาง (HR กรอก) |
| `Employee_List` | พนักงานที่ผูก LINE แล้ว — คอลัมน์: รหัส, ชื่อ, แผนก, **LINE ID**, ตำแหน่ง, ประเภท, Direct/Indirect, LINE Name, วันลงทะเบียน, Department |
| `Job_List` | ทะเบียนงาน + QR Link (gen อัตโนมัติเมื่อแก้แถว) — ประเภทงาน, Job Code, ชื่องาน/ลูกค้า, สถานะ, QR |
| `Job_Log_2569_MM` | **log หลัก** สร้างใหม่ทุกเดือนตามรอบเงินเดือน (ตัดวันที่ 26) — 21 คอลัมน์ |
| (สรุปรายวัน) | วันที่, รหัส, ชื่อ, แผนก, ชม.ปกติรวม, ชม.OT รวม |
| (สรุป Job/แผนก) | วันที่, แผนก, ชื่อ, Job ID, ชื่องาน, ชม.ปกติ, OT x1.5, ผลรวม |

**คอลัมน์ Job_Log (index ใน COL):** Timestamp(0) วันที่(1) เวลาเข้า(2) เวลาออก(3) ชม.รวม(4) JobID(5) ชื่องาน(6) Zone(7) รหัสพนักงาน(8) ชื่อ(9) แผนก(10) สถานะ(11) Process(12) ภาษา(13) หมายเหตุ(14) ประเภทวัน(15) ประเภทชั่วโมง(16) ชม.ปกติ HH:MM(17) ชม.OT HH:MM(18) ชม.ปกติ ทศนิยม(19) ชม.OT ทศนิยม(20)

---

## 6. Frontend — จุดสำคัญ

- **Config (บรรทัดบนสุด):** `APPS_SCRIPT_URL`, `IMGBB_API_KEY`, `LINE_CHANNEL_ID`, `LINE_REDIRECT`
- **รับ Job จาก URL:** `?job=&name=&type=&zone=` ถ้าไม่มี name → โหลดจาก Sheet ผ่าน `GET_JOB_INFO`
- **Multi-language:** th / la / mm / kh / en (ไทย/ลาว/พม่า/เขมร/อังกฤษ) — รองรับแรงงานต่างชาติ
- **Process แยกตามแผนก:** `getProcsByDept()` → `PROCS_PRODUCTION` (16 รายการ) หรือ `PROCS_QC` (5 รายการ)
- **เก็บ session:** `localStorage`/`sessionStorage` key `jobtrack_user`
- Debug mode: เพิ่ม `?debug=1` ใน URL

---

## 7. Config / Secrets (⚠️ ปัจจุบันฝังใน Code)

| ค่า | ที่อยู่ |
|-----|---------|
| Sheet ID | `jobtrack_apps_script.gs` บรรทัด 5 |
| LINE Channel ID / Secret | `.gs` บรรทัด 10-11 |
| Apps Script Web App URL | `job_checkin_app.html` บรรทัด 5 |
| imgbb API Key | `job_checkin_app.html` บรรทัด 6 |
| LINE Channel ID (frontend) | `job_checkin_app.html` บรรทัด 7 |

---

## 8. ประเด็นที่ควรพิจารณาตอน Develop ต่อ (Candy สังเกตเห็น)

1. **Secrets ฝังใน code** — LINE Channel Secret + API keys อยู่ใน HTML/GS ที่ public บน GitHub Pages
2. **ตารางสรุปรายวันมีค่า 0 หลายแถว** — อาจเป็นสูตร/import ที่ยังไม่ครบ ควรเช็ค logic การ sync
3. **ยังไม่มี Dashboard** — ข้อมูลอยู่ใน Sheet ดิบ ยังไม่มีหน้าสรุปผู้บริหาร/รายงานอัตโนมัติ
4. **การคำนวณ OT** ใช้กฎตายตัว 17:00 — ถ้านโยบายจริงซับซ้อนกว่านี้ (พักเที่ยง, กะกลางคืน) ต้องปรับ
5. **ไม่มีระบบ approve** — Check Out แล้วเข้า Sheet เลย ยังไม่มีขั้นหัวหน้าอนุมัติ

---

*อัปเดต: 30 มิ.ย. 2026 — เอกสารนี้เป็นฐานอ้างอิงสำหรับการพัฒนาต่อ*
