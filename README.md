# QR Code Generator Web Application

เว็บแอปพลิเคชันสำหรับสร้าง QR Code จากข้อมูลในไฟล์ Excel

## คุณสมบัติ

- อัปโหลดไฟล์ Excel (.xlsx, .xls) หรือ CSV
- สร้าง QR Code สำหรับแต่ละแถวของข้อมูล
- ดาวน์โหลดไฟล์ QR Code ทั้งหมดในรูปแบบ ZIP
- รองรับข้อมูลหลายคอลัมน์ (Name, Phone, Email, Address เป็นต้น)

## การใช้งาน

### 1. เริ่มเซิร์ฟเวอร์

```bash
npm install
node server.js
```

### 2. เปิดเว็บไซต์

เปิดเบราว์เซอร์ที่ http://localhost:3000

### 3. อัปโหลดไฟล์

- คลิกปุ่ม "เลือกไฟล์" เพื่อเลือกไฟล์ Excel หรือ CSV
- คลิก "สร้าง QR Code" เพื่อเริ่มการสร้าง
- รอจนกว่าระบบจะสร้าง QR Code เสร็จสิ้น
- คลิก "ดาวน์โหลด ZIP" เพื่อดาวน์โหลดไฟล์ทั้งหมด

## รูปแบบไฟล์ที่รองรับ

- Excel (.xlsx, .xls)
- CSV (.csv)

## ตัวอย่างข้อมูล

ระบบจะรวมข้อมูลทุกคอลัมน์ในแต่ละแถวเป็น QR Code เดียว
ตัวอย่าง: "Name: John Doe, Phone: 081-234-5678, Email: john@email.com, Address: 123 Main St Bangkok"

## ไฟล์ตัวอย่าง

ดูไฟล์ `sample-data.csv` เพื่อดูตัวอย่างการจัดรูปแบบข้อมูล

## Dependencies

- express: เว็บเซิร์ฟเวอร์
- multer: การอัปโหลดไฟล์
- xlsx: อ่านไฟล์ Excel
- qrcode: สร้าง QR Code
- archiver: สร้างไฟล์ ZIP
