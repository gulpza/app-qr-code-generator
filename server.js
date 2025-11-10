const express = require('express');
const multer = require('multer');
const xlsx = require('xlsx');
const QRCode = require('qrcode');
const archiver = require('archiver');
const path = require('path');
const fs = require('fs');
const { createCanvas, loadImage } = require('canvas');

const app = express();
const PORT = process.env.PORT || 3000;

// กำหนด middleware
app.use(express.static('public'));
app.use(express.json());

// สร้างโฟลเดอร์สำหรับเก็บไฟล์ชั่วคราว
const tempDir = path.join(__dirname, 'temp');
if (!fs.existsSync(tempDir)) {
    fs.mkdirSync(tempDir);
}

// กำหนดการอัปโหลดไฟล์
const storage = multer.diskStorage({
    destination: (req, file, cb) => {
        cb(null, tempDir);
    },
    filename: (req, file, cb) => {
        cb(null, Date.now() + '-' + file.originalname);
    }
});

const upload = multer({ 
    storage: storage,
    fileFilter: (req, file, cb) => {
        const allowedTypes = /xlsx|xls/;
        const extname = allowedTypes.test(path.extname(file.originalname).toLowerCase());
        const mimetype = allowedTypes.test(file.mimetype) || 
                        file.mimetype === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' ||
                        file.mimetype === 'application/vnd.ms-excel';
        
        if (mimetype && extname) {
            return cb(null, true);
        } else {
            cb('Error: กรุณาอัปโหลดไฟล์ Excel เท่านั้น (.xlsx หรือ .xls)');
        }
    }
});

// Route หลัก
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// Route สำหรับดาวน์โหลด Excel Template
app.get('/download-template', (req, res) => {
    try {
        // สร้าง workbook และ worksheet
        const wb = xlsx.utils.book_new();
        
        // สร้าง worksheet ด้วย header เท่านั้น
        const ws = xlsx.utils.aoa_to_sheet([
            ['Code', 'Name']  // Header row เท่านั้น
        ]);

        // ปรับความกว้างคอลัมน์
        ws['!cols'] = [
            { wch: 15 }, // Code
            { wch: 25 }  // Name
        ];

        // เพิ่ม worksheet เข้า workbook
        xlsx.utils.book_append_sheet(wb, ws, 'Template');

        // สร้างไฟล์ Excel ชั่วคราว
        const templatePath = path.join(tempDir, `template-${Date.now()}.xlsx`);
        xlsx.writeFile(wb, templatePath);

        // ส่งไฟล์ให้ดาวน์โหลด
        res.download(templatePath, 'QR-Code-Template.xlsx', (err) => {
            if (err) {
                console.error('Error downloading template:', err);
            }
            // ลบไฟล์ชั่วคราว
            setTimeout(() => {
                try {
                    fs.unlinkSync(templatePath);
                } catch (cleanupErr) {
                    console.error('Error cleaning up template:', cleanupErr);
                }
            }, 5000);
        });
    } catch (error) {
        console.error('Error creating template:', error);
        res.status(500).json({ error: 'เกิดข้อผิดพลาดในการสร้าง template' });
    }
});

// Route สำหรับสร้าง QR Code จากไฟล์ Excel
app.post('/generate-qr', upload.single('excelFile'), async (req, res) => {
    try {
        if (!req.file) {
            return res.status(400).json({ error: 'กรุณาเลือกไฟล์ Excel' });
        }

        // อ่านไฟล์ Excel
        const workbook = xlsx.readFile(req.file.path);
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const data = xlsx.utils.sheet_to_json(worksheet);

        if (data.length === 0) {
            return res.status(400).json({ error: 'ไฟล์ Excel ว่างเปล่า' });
        }

        // สร้างโฟลเดอร์สำหรับ QR codes
        const qrDir = path.join(tempDir, `qr-${Date.now()}`);
        fs.mkdirSync(qrDir);

        // สร้าง QR Code สำหรับแต่ละแถว
        const promises = data.map(async (row, index) => {
          
            // สร้างข้อมูลสำหรับ QR Code
            let qrData = row.Code ? row.Code.toString() : `Row ${index + 1}`;
            
            // ถ้ามีข้อมูลใน Name ให้เพิ่มเข้าไปใน QR Code
            if (row.Name && row.Name.toString().trim() !== '') {
                qrData += '|' + row.Name.toString();
            }
            
            // สร้างชื่อไฟล์ (ใช้เฉพาะ Code)
            const fileNameBase = row.Code ? row.Code.toString() : `Row_${index + 1}`;
            const fileName = `${fileNameBase}.png`.replace(/[^a-zA-Z0-9._-]/g, '_');
            const filePath = path.join(qrDir, fileName);

            // สร้าง QR Code เป็น buffer
            const qrBuffer = await QRCode.toBuffer(qrData, {
                width: 300,
                margin: 2,
                color: {
                    dark: '#000000',
                    light: '#FFFFFF'
                }
            });

            // สร้าง canvas สำหรับรวม QR code กับข้อความ
            const qrImage = await loadImage(qrBuffer);
            const canvasWidth = 350;
            const canvasHeight = 350; // เพิ่มความสูงสำหรับข้อความ
            const canvas = createCanvas(canvasWidth, canvasHeight);
            const ctx = canvas.getContext('2d');

            // ตั้งค่าพื้นหลังสีขาว
            ctx.fillStyle = '#FFFFFF';
            ctx.fillRect(0, 0, canvasWidth, canvasHeight);

            // วาด QR Code ตรงกลาง
            const qrX = (canvasWidth - qrImage.width) / 2;
            const qrY = 10;
            ctx.drawImage(qrImage, qrX, qrY);

            // เพิ่มข้อความใต้ QR Code (แสดงเฉพาะ Code)
            const displayText = row.Code ? row.Code.toString() : `Row ${index + 1}`;
            ctx.fillStyle = '#000000';
            ctx.font = 'bold 18px Arial';
            ctx.textAlign = 'center';
            
            // คำนวณตำแหน่งข้อความ
            const textY = qrY + qrImage.height + 10;
            ctx.fillText(displayText, canvasWidth / 2, textY);

            // บันทึกเป็นไฟล์ PNG
            const buffer = canvas.toBuffer('image/png');
            fs.writeFileSync(filePath, buffer);

            return fileName;
        });

        await Promise.all(promises);

        // สร้างไฟล์ ZIP
        const zipPath = path.join(tempDir, `qr-codes-${Date.now()}.zip`);
        const output = fs.createWriteStream(zipPath);
        const archive = archiver('zip', {
            zlib: { level: 9 }
        });

        archive.pipe(output);
        archive.directory(qrDir, false);
        await archive.finalize();

        // ส่งไฟล์ ZIP ให้ดาวน์โหลด
        output.on('close', () => {
            res.download(zipPath, 'qr-codes.zip', (err) => {
                if (err) {
                    console.error('Error downloading file:', err);
                    res.status(500).json({ error: 'เกิดข้อผิดพลาดในการดาวน์โหลด' });
                }

                // ลบไฟล์ชั่วคราว
                setTimeout(() => {
                    try {
                        fs.unlinkSync(req.file.path); // ลบไฟล์ Excel
                        fs.unlinkSync(zipPath); // ลบไฟล์ ZIP
                        fs.rmSync(qrDir, { recursive: true, force: true }); // ลบโฟลเดอร์ QR codes
                    } catch (cleanupErr) {
                        console.error('Error cleaning up files:', cleanupErr);
                    }
                }, 5000);
            });
        });

    } catch (error) {
        console.error('Error processing file:', error);
        res.status(500).json({ error: 'เกิดข้อผิดพลาดในการประมวลผลไฟล์' });
    }
});

// เริ่มเซิร์ฟเวอร์
app.listen(PORT, () => {
    console.log(`Server is running on http://localhost:${PORT}`);
});