const xlsx = require('xlsx');
const path = require('path');
const fs = require('fs');

// ใช้ /tmp directory สำหรับ Vercel
const tempDir = '/tmp';

module.exports = async (req, res) => {
    // ตั้งค่า CORS
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'GET, OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

    if (req.method === 'OPTIONS') {
        return res.status(200).end();
    }

    if (req.method !== 'GET') {
        return res.status(405).json({ error: 'Method not allowed' });
    }

    let templatePath = null;

    try {
        // สร้าง workbook และ worksheet
        const wb = xlsx.utils.book_new();
        
        // สร้าง worksheet ด้วย header และ sample data
        const ws = xlsx.utils.aoa_to_sheet([
            ['Code', 'Name'],  // Header row
            ['001', 'Sample Item 1'],  // Sample row 1
            ['002', 'Sample Item 2']   // Sample row 2
        ]);

        // ปรับความกว้างคอลัมน์
        ws['!cols'] = [
            { wch: 15 }, // Code
            { wch: 25 }  // Name
        ];

        // เพิ่ม worksheet เข้า workbook
        xlsx.utils.book_append_sheet(wb, ws, 'Template');

        // สร้างไฟล์ Excel ชั่วคราว
        templatePath = path.join(tempDir, `template-${Date.now()}.xlsx`);
        xlsx.writeFile(wb, templatePath);

        // อ่านไฟล์และส่งกลับ
        const fileBuffer = fs.readFileSync(templatePath);

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename=QR-Code-Template.xlsx');
        res.setHeader('Content-Length', fileBuffer.length);
        res.send(fileBuffer);

    } catch (error) {
        console.error('Error creating template:', error);
        res.status(500).json({ error: 'เกิดข้อผิดพลาดในการสร้าง template' });
    } finally {
        // ลบไฟล์ชั่วคราว
        setTimeout(() => {
            try {
                if (templatePath && fs.existsSync(templatePath)) {
                    fs.unlinkSync(templatePath);
                }
            } catch (cleanupErr) {
                console.error('Error cleaning up template:', cleanupErr);
            }
        }, 1000);
    }
};
