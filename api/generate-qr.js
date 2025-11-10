const multer = require('multer');
const xlsx = require('xlsx');
const QRCode = require('qrcode');
const archiver = require('archiver');
const path = require('path');
const fs = require('fs');
const { createCanvas, loadImage } = require('canvas');

// ใช้ /tmp directory สำหรับ Vercel
const tempDir = '/tmp';

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
    limits: {
        fileSize: 10 * 1024 * 1024 // 10MB limit
    },
    fileFilter: (req, file, cb) => {
        const allowedTypes = /xlsx|xls/;
        const extname = allowedTypes.test(path.extname(file.originalname).toLowerCase());
        const mimetype = allowedTypes.test(file.mimetype) || 
                        file.mimetype === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' ||
                        file.mimetype === 'application/vnd.ms-excel';
        
        if (mimetype && extname) {
            return cb(null, true);
        } else {
            cb(new Error('กรุณาอัปโหลดไฟล์ Excel เท่านั้น (.xlsx หรือ .xls)'));
        }
    }
});

// Helper function สำหรับจัดการ multer แบบ promise
const multerMiddleware = (req, res) => {
    return new Promise((resolve, reject) => {
        upload.single('excelFile')(req, res, (err) => {
            if (err) {
                reject(err);
            } else {
                resolve();
            }
        });
    });
};

module.exports = async (req, res) => {
    // ตั้งค่า CORS
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

    if (req.method === 'OPTIONS') {
        return res.status(200).end();
    }

    if (req.method !== 'POST') {
        return res.status(405).json({ error: 'Method not allowed' });
    }

    let uploadedFilePath = null;
    let qrDir = null;
    let zipPath = null;

    try {
        // อัปโหลดไฟล์
        await multerMiddleware(req, res);

        if (!req.file) {
            return res.status(400).json({ error: 'กรุณาเลือกไฟล์ Excel' });
        }

        uploadedFilePath = req.file.path;

        // อ่านไฟล์ Excel
        const workbook = xlsx.readFile(uploadedFilePath);
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const data = xlsx.utils.sheet_to_json(worksheet);

        if (data.length === 0) {
            return res.status(400).json({ error: 'ไฟล์ Excel ว่างเปล่า' });
        }

        // จำกัดจำนวนแถวเพื่อป้องกัน timeout
        if (data.length > 1000) {
            return res.status(400).json({ 
                error: `ไฟล์มีข้อมูลมากเกินไป (${data.length} แถว) กรุณาใช้ไฟล์ที่มีข้อมูลไม่เกิน 1000 แถว` 
            });
        }

        // สร้างโฟลเดอร์สำหรับ QR codes
        qrDir = path.join(tempDir, `qr-${Date.now()}`);
        if (!fs.existsSync(qrDir)) {
            fs.mkdirSync(qrDir, { recursive: true });
        }

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

            // สร้าง canvas แบบง่าย - แสดงเฉพาะ QR Code ไม่มีข้อความ
            // (เพราะข้อมูลอยู่ใน QR Code อยู่แล้ว)
            const qrImage = await loadImage(qrBuffer);
            const padding = 20;
            const canvasWidth = qrImage.width + (padding * 2);
            const canvasHeight = qrImage.height + (padding * 2);
            const canvas = createCanvas(canvasWidth, canvasHeight);
            const ctx = canvas.getContext('2d');

            // ตั้งค่าพื้นหลังสีขาว
            ctx.fillStyle = '#FFFFFF';
            ctx.fillRect(0, 0, canvasWidth, canvasHeight);

            // วาด QR Code ตรงกลาง
            ctx.drawImage(qrImage, padding, padding);

            // บันทึกเป็นไฟล์ PNG
            const buffer = canvas.toBuffer('image/png');
            fs.writeFileSync(filePath, buffer);

            return fileName;
        });

        await Promise.all(promises);

        // สร้างไฟล์ ZIP
        zipPath = path.join(tempDir, `qr-codes-${Date.now()}.zip`);
        const output = fs.createWriteStream(zipPath);
        const archive = archiver('zip', {
            zlib: { level: 9 }
        });

        archive.pipe(output);
        archive.directory(qrDir, false);
        await archive.finalize();

        // รอให้ ZIP เสร็จสมบูรณ์
        await new Promise((resolve, reject) => {
            output.on('close', resolve);
            output.on('error', reject);
        });

        // อ่านไฟล์ ZIP และส่งกลับ
        const zipBuffer = fs.readFileSync(zipPath);
        
        res.setHeader('Content-Type', 'application/zip');
        res.setHeader('Content-Disposition', 'attachment; filename=qr-codes.zip');
        res.setHeader('Content-Length', zipBuffer.length);
        res.send(zipBuffer);

    } catch (error) {
        console.error('Error processing file:', error);
        
        // ส่ง error message ที่เหมาะสม
        const errorMessage = error.message || 'เกิดข้อผิดพลาดในการประมวลผลไฟล์';
        res.status(500).json({ error: errorMessage });
    } finally {
        // ลบไฟล์ชั่วคราว
        setTimeout(() => {
            try {
                if (uploadedFilePath && fs.existsSync(uploadedFilePath)) {
                    fs.unlinkSync(uploadedFilePath);
                }
                if (zipPath && fs.existsSync(zipPath)) {
                    fs.unlinkSync(zipPath);
                }
                if (qrDir && fs.existsSync(qrDir)) {
                    fs.rmSync(qrDir, { recursive: true, force: true });
                }
            } catch (cleanupErr) {
                console.error('Error cleaning up files:', cleanupErr);
            }
        }, 1000);
    }
};
