document.addEventListener('DOMContentLoaded', function() {
    const uploadForm = document.getElementById('uploadForm');
    const progressContainer = document.getElementById('progressContainer');
    const progressBar = document.getElementById('progressBar');
    const successMessage = document.getElementById('successMessage');
    const errorMessage = document.getElementById('errorMessage');
    const errorText = document.getElementById('errorText');
    const submitButton = uploadForm.querySelector('button[type="submit"]');
    const fileInput = document.getElementById('excelFile');

    // เพิ่ม event listener สำหรับการเปลี่ยนแปลงไฟล์
    fileInput.addEventListener('change', function() {
        const file = this.files[0];
        if (file) {
            // ตรวจสอบประเภทไฟล์
            const allowedTypes = ['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 
                                'application/vnd.ms-excel'];
            const fileExtension = file.name.split('.').pop().toLowerCase();
            
            if (!allowedTypes.includes(file.type) && !['xlsx', 'xls'].includes(fileExtension)) {
                showError('กรุณาเลือกไฟล์ Excel (.xlsx หรือ .xls) เท่านั้น');
                this.value = '';
                return;
            }
            
            // ตรวจสอบขนาดไฟล์ (จำกัดที่ 10MB)
            if (file.size > 10 * 1024 * 1024) {
                showError('ไฟล์มีขนาดใหญ่เกินไป กรุณาเลือกไฟล์ที่มีขนาดน้อยกว่า 10MB');
                this.value = '';
                return;
            }
            
            hideMessages();
        }
    });

    uploadForm.addEventListener('submit', async function(e) {
        e.preventDefault();
        
        const file = fileInput.files[0];
        if (!file) {
            showError('กรุณาเลือกไฟล์ Excel');
            return;
        }

        // เริ่มกระบวนการอัปโหลด
        startProgress();
        
        const formData = new FormData();
        formData.append('excelFile', file);

        try {
            // อัปเดต progress bar
            updateProgress(30, 'กำลังอัปโหลดไฟล์...');
            
            const response = await fetch('/generate-qr', {
                method: 'POST',
                body: formData
            });

            updateProgress(60, 'กำลังอ่านข้อมูล Excel...');

            if (response.ok) {
                updateProgress(90, 'กำลังสร้าง QR Code...');
                
                // ตรวจสอบ content type ก่อน
                const contentType = response.headers.get('content-type');
                
                if (contentType && contentType.includes('application/zip')) {
                    // รับไฟล์ ZIP
                    const blob = await response.blob();
                    
                    updateProgress(100, 'เสร็จสิ้น!');
                    
                    // สร้างลิงค์ดาวน์โหลด
                    const url = window.URL.createObjectURL(blob);
                    const a = document.createElement('a');
                    a.href = url;
                    a.download = `qr-codes-${new Date().getTime()}.zip`;
                    document.body.appendChild(a);
                    a.click();
                    document.body.removeChild(a);
                    window.URL.revokeObjectURL(url);

                    // แสดงข้อความสำเร็จ
                    setTimeout(() => {
                        stopProgress();
                        showSuccess();
                        
                        // รีเซ็ตฟอร์มหลังจาก 3 วินาที
                        setTimeout(() => {
                            resetForm();
                        }, 3000);
                    }, 1000);
                } else {
                    // ไม่ใช่ ZIP file - อาจเป็น error
                    const text = await response.text();
                    console.error('Unexpected response:', text);
                    throw new Error('Server ไม่ได้ส่งไฟล์ ZIP กลับมา กรุณาลองใหม่อีกครั้ง');
                }

            } else {
                // จัดการ error response
                const contentType = response.headers.get('content-type');
                
                if (contentType && contentType.includes('application/json')) {
                    const errorData = await response.json();
                    throw new Error(errorData.error || 'เกิดข้อผิดพลาดในการประมวลผลไฟล์');
                } else {
                    // Server ส่ง HTML หรือ text กลับมา
                    const text = await response.text();
                    console.error('Server response:', text);
                    
                    if (response.status === 504 || response.status === 502) {
                        throw new Error('การประมวลผลใช้เวลานานเกินไป กรุณาลดจำนวนแถวข้อมูลในไฟล์ (แนะนำไม่เกิน 100 แถว)');
                    } else if (response.status === 413) {
                        throw new Error('ไฟล์มีขนาดใหญ่เกินไป กรุณาเลือกไฟล์ที่เล็กกว่า 10MB');
                    } else {
                        throw new Error(`เกิดข้อผิดพลาด (${response.status}): กรุณาตรวจสอบรูปแบบไฟล์และลองใหม่อีกครั้ง`);
                    }
                }
            }

        } catch (error) {
            console.error('Error:', error);
            stopProgress();
            showError(error.message || 'เกิดข้อผิดพลาดในการเชื่อมต่อกับเซิร์ฟเวอร์');
        }
    });

    function startProgress() {
        hideMessages();
        progressContainer.style.display = 'block';
        progressBar.style.width = '0%';
        progressBar.textContent = '';
        
        submitButton.disabled = true;
        submitButton.innerHTML = '<i class="fas fa-spinner fa-spin me-2"></i>กำลังประมวลผล...';
    }

    function updateProgress(percent, message) {
        progressBar.style.width = percent + '%';
        progressBar.textContent = Math.round(percent) + '%';
        
        const container = progressContainer.querySelector('strong');
        if (container && message) {
            container.textContent = message;
        }
    }

    function stopProgress() {
        progressContainer.style.display = 'none';
        submitButton.disabled = false;
        submitButton.innerHTML = '<i class="fas fa-magic me-2"></i>สร้าง QR Code';
    }

    function showSuccess() {
        successMessage.style.display = 'block';
        errorMessage.style.display = 'none';
    }

    function showError(message) {
        errorText.textContent = message;
        errorMessage.style.display = 'block';
        successMessage.style.display = 'none';
    }

    function hideMessages() {
        successMessage.style.display = 'none';
        errorMessage.style.display = 'none';
    }

    function resetForm() {
        uploadForm.reset();
        hideMessages();
        fileInput.value = '';
    }

    // เพิ่มการจัดการ drag and drop
    const fileInputGroup = fileInput.parentElement;
    
    fileInputGroup.addEventListener('dragover', function(e) {
        e.preventDefault();
        this.classList.add('border-primary');
        this.style.backgroundColor = '#f8f9fa';
    });

    fileInputGroup.addEventListener('dragleave', function(e) {
        e.preventDefault();
        this.classList.remove('border-primary');
        this.style.backgroundColor = '';
    });

    fileInputGroup.addEventListener('drop', function(e) {
        e.preventDefault();
        this.classList.remove('border-primary');
        this.style.backgroundColor = '';
        
        const files = e.dataTransfer.files;
        if (files.length > 0) {
            fileInput.files = files;
            // เรียก event change manually
            const event = new Event('change', { bubbles: true });
            fileInput.dispatchEvent(event);
        }
    });
});