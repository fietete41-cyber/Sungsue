# ระบบสำรวจ/สั่งจองของที่ระลึก (เสื้อ / แก้ว / แหวน)

โครงสร้าง: **หน้าเว็บ (index.html) บน GitHub Pages** ↔ **Google Apps Script (Code.gs) เป็น API** ↔ **Google Sheets** เป็นฐานข้อมูล

## ติดตั้ง Backend (Apps Script)
1. สร้าง Google Sheets ใหม่ → Extensions > Apps Script
2. วางเนื้อหา `Code.gs` ทับไฟล์เดิม (ไม่ต้องสร้างไฟล์ HTML ใน Apps Script)
3. เลือกฟังก์ชัน `setup` แล้ว Run (อนุญาตสิทธิ์) → ดูรหัสแอดมินที่ Execution log / ผลลัพธ์ และจดไว้
4. Deploy > New deployment > Web app — Execute as: **Me**, Who has access: **Anyone** → คัดลอก URL ที่ลงท้าย `/exec`
5. แก้โค้ดภายหลัง: Deploy > Manage deployments > แก้ไข > New version

## ติดตั้ง Frontend (GitHub Pages)
1. เปิด `index.html` แก้ `API_URL` เป็น URL จากข้อ 4
2. push ขึ้น repo → Settings > Pages > Deploy from branch (main / root)

## ตั้งค่า (Apps Script > Project Settings > Script properties)
| Key | ความหมาย |
|---|---|
| `ADMIN_PASS` | รหัสแอดมิน (สุ่มให้ตอน setup) |
| `PAY_INFO` | ข้อความบัญชีรับโอน |
| `SHOP_TITLE` | ชื่อหน้าร้าน |
| `MODE` | survey / pay / order / closed (สลับได้ในหน้าแอดมิน) |

## ขั้นตอนใช้งาน
สำรวจ (ยังไม่โอน) → แอดมินดูสรุปยอด → กด "เปิดชำระเงิน" → ลูกค้าเข้า "ตรวจสอบ" ด้วยเบอร์โทรเพื่อโอนและแนบสลิป → แอดมินเปลี่ยนสถานะ

เพิ่มสินค้า/แก้ตัวเลือกไซซ์ ทำในชีต Products โดยตรง

## รูปสินค้า
รูปอยู่ในโฟลเดอร์ `img/` (ภาพครอปสำหรับการ์ด + โปสเตอร์เต็มที่กดดูขยายได้) อ้างอิงจากคอลัมน์ "รูป (URL)" และ "รูปใหญ่ (URL)" ในชีต Products เปลี่ยนรูปโดยแทนไฟล์ชื่อเดิม หรือแก้ path ในชีต

## ราคา
ราคา = 0 จะแสดงเป็น "ราคาแจ้งภายหลัง" (เหมาะกับช่วงสำรวจ) กรอกราคาจริงได้ที่ชีต Products หรือหน้าแอดมิน

## รีเซ็ตสินค้า
รันฟังก์ชัน `resetProducts` ใน Apps Script เพื่อสร้างชีต Products ใหม่เป็นค่าเริ่มต้น (เสื้อโปโล / เสื้อยืด / แก้ว / แหวน) โดยไม่แตะออเดอร์
