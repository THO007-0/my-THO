let names = [];            // รายชื่อทั้งหมดจาก Excel หรือ Google Sheets
let currentIndex = 0;      // ดัชนีรายชื่อที่กำลังแสดง
let currentTemplate = 'hakiri.jpg'; // เทมเพลตพื้นหลังเริ่มต้น

// โหลดข้อมูลจาก Excel ที่อัปโหลด
document.getElementById('excelFile').addEventListener('change', (e) => {
  const file = e.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = function (event) {
    const data = new Uint8Array(event.target.result);
    const workbook = XLSX.read(data, { type: 'array' });
    const firstSheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[firstSheetName];
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

    names = jsonData
      .map(row => row[0])
      .filter(name => typeof name === 'string' && name.trim().length > 0);

    if (names.length > 0) {
      currentIndex = 0;
      updateDisplay();
      updatePageNum();
    } else {
      alert('ไม่พบชื่อในไฟล์ Excel');
    }
  };
  reader.readAsArrayBuffer(file);
});

// โหลดรายชื่อจาก Google Sheets CSV
async function loadFromSheet() {
  const urlInput = document.getElementById('sheetUrl').value.trim();
  if (!urlInput) {
    alert('กรุณาใส่ URL Google Sheets ที่แชร์แบบ public');
    return;
  }

  try {
    const match = urlInput.match(/\/d\/([a-zA-Z0-9-_]+)/);
    if (!match) throw new Error('URL ไม่ถูกต้อง');

    const sheetId = match[1];
    const csvUrl = `https://docs.google.com/spreadsheets/d/${sheetId}/export?format=csv&id=${sheetId}&gid=0`;

    const response = await fetch(csvUrl);
    if (!response.ok) throw new Error('ไม่สามารถโหลดข้อมูลจาก Google Sheets');

    const csvText = await response.text();
    const rows = csvText.split('\n').map(row => row.split(','));
    names = rows
      .map(row => row[0])
      .filter(name => typeof name === 'string' && name.trim().length > 0);

    if (names.length > 0) {
      currentIndex = 0;
      updateDisplay();
      updatePageNum();
    } else {
      alert('ไม่พบชื่อใน Google Sheets');
    }
  } catch (error) {
    alert('เกิดข้อผิดพลาด: ' + error.message);
  }
}

// อัพเดตข้อความบนใบประกาศตามข้อมูลและอินพุต
function updateDisplay() {
  document.getElementById('nameText').textContent = names[currentIndex] || 'Your Name';
  document.getElementById('titleText').textContent = document.getElementById('title').value;
  document.getElementById('presentedText').textContent = document.getElementById('presented').value;
  document.getElementById('reasonText').textContent = document.getElementById('reason').value;
}

// อัพเดตตัวเลขแสดงหน้าปัจจุบัน / ทั้งหมด
function updatePageNum() {
  const pageNumEl = document.getElementById('pageNum');
  pageNumEl.textContent = `${names.length > 0 ? currentIndex + 1 : 0} / ${names.length || 1}`;
}

// แสดงชื่อถัดไป
function nextName() {
  if (names.length === 0) return;
  currentIndex = (currentIndex + 1) % names.length;
  updateDisplay();
  updatePageNum();
}

// แสดงชื่อก่อนหน้า
function prevName() {
  if (names.length === 0) return;
  currentIndex = (currentIndex - 1 + names.length) % names.length;
  updateDisplay();
  updatePageNum();
}

// เลือกเทมเพลตและบันทึกลง localStorage
function selectTemplate(imgEl, templateSrc) {
  currentTemplate = templateSrc;
  document.getElementById('templateImg').src = templateSrc;
  document.querySelectorAll('.template-thumbnail').forEach(img => img.classList.remove('selected'));
  imgEl.classList.add('selected');
  localStorage.setItem('selectedTemplate', templateSrc);
}

// อัปโหลดเทมเพลตใหม่จากไฟล์ภาพ
function uploadTemplate(inputEl) {
  const file = inputEl.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = function (e) {
    const dataUrl = e.target.result;
    currentTemplate = dataUrl;
    document.getElementById('templateImg').src = dataUrl;
    localStorage.setItem('selectedTemplate', dataUrl);

    const container = document.querySelector('.template-thumbnails');
    const newImg = document.createElement('img');
    newImg.src = dataUrl;
    newImg.classList.add('template-thumbnail');
    newImg.onclick = () => selectTemplate(newImg, dataUrl);

    container.appendChild(newImg);
    selectTemplate(newImg, dataUrl); // เลือกและไฮไลต์ทันที
  };
  reader.readAsDataURL(file);
}

// โหลดเทมเพลตที่เคยเลือกไว้จาก localStorage เมื่อเปิดหน้าเว็บ
window.addEventListener('DOMContentLoaded', () => {
  const savedTemplate = localStorage.getItem('selectedTemplate');
  if (savedTemplate) {
    currentTemplate = savedTemplate;
    document.getElementById('templateImg').src = savedTemplate;

    if (savedTemplate.startsWith('data:')) {
      const container = document.querySelector('.template-thumbnails');
      const newImg = document.createElement('img');
      newImg.src = savedTemplate;
      newImg.classList.add('template-thumbnail');
      newImg.onclick = () => selectTemplate(newImg, savedTemplate);
      container.appendChild(newImg);
      selectTemplate(newImg, savedTemplate);
    } else {
      const thumbnails = document.querySelectorAll('.template-thumbnail');
      thumbnails.forEach(img => {
        if (img.src.endsWith(savedTemplate)) {
          selectTemplate(img, savedTemplate);
        }
      });
    }
  }
});

// ฟังก์ชันดาวน์โหลดใบประกาศ (PDF หรือ JPG)
async function downloadCertificate() {
  const fileType = document.getElementById('fileType').value;
  const certificate = document.getElementById('certificatePreview');

  // แปลง preview เป็น canvas
  const canvas = await html2canvas(certificate, { scale: 2 });

  if (fileType === 'pdf') {
    const imgData = canvas.toDataURL('image/png');
    const pdf = new jspdf.jsPDF({
      orientation: 'landscape',
      unit: 'px',
      format: [certificate.offsetWidth, certificate.offsetHeight],
    });
    pdf.addImage(imgData, 'PNG', 0, 0, certificate.offsetWidth, certificate.offsetHeight);
    pdf.save(`certificate_${sanitizeFileName(names[currentIndex] || 'name')}.pdf`);
  } else if (fileType === 'jpg') {
    const link = document.createElement('a');
    link.download = `certificate_${sanitizeFileName(names[currentIndex] || 'name')}.jpg`;
    link.href = canvas.toDataURL('image/jpeg', 1.0);
    link.click();
  }
}

// ฟังก์ชันดาวน์โหลดใบประกาศทั้งหมดทีละชื่อ (หน่วงเวลา 500ms เพื่อไม่ให้โหลดเกินไป)
async function downloadAll() {
  if (names.length === 0) {
    alert('ไม่มีรายชื่อให้ดาวน์โหลด');
    return;
  }

  const fileType = document.getElementById('fileType').value;

  for (let i = 0; i < names.length; i++) {
    currentIndex = i;
    updateDisplay();
    updatePageNum();
    await new Promise(r => setTimeout(r, 500));  // หน่วง 0.5 วินาที
    await downloadCertificate();
  }

  currentIndex = 0;
  updateDisplay();
  updatePageNum();
}

// ฟังก์ชันช่วย sanitize ชื่อไฟล์ (ตัดอักขระที่ไม่ปลอดภัยออก)
function sanitizeFileName(name) {
  return name.replace(/[\/\\?%*:|"<>]/g, '-');
}

async function downloadCertificate() {
  const fileType = document.getElementById('fileType').value;
  const certificate = document.getElementById('certificatePreview');
  const CERT_WIDTH = 2667;  // 27.78 นิ้ว @ 96dpi
const CERT_HEIGHT = 1777; // 18.51 นิ้ว @ 96dpi


  // แปลง preview เป็น canvas ความละเอียดสูง
  const canvas = await html2canvas(certificate, { scale: 2, useCORS: true });

  if (fileType === 'pdf') {
    const imgData = canvas.toDataURL('image/png');

    const pdf = new jspdf.jsPDF({
      orientation: 'landscape',
      unit: 'px',
      format: [CERT_WIDTH, CERT_HEIGHT]
    });

    pdf.addImage(imgData, 'PNG', 0, 0, CERT_WIDTH, CERT_HEIGHT);
    pdf.save(`certificate_${sanitizeFileName(names[currentIndex] || 'name')}.pdf`);

  } else if (fileType === 'jpg') {
    const link = document.createElement('a');
    link.download = `certificate_${sanitizeFileName(names[currentIndex] || 'name')}.jpg`;
    link.href = canvas.toDataURL('image/jpeg', 1.0);
    link.click();
  }
}


// อัพเดตข้อความบนใบประกาศเมื่อแก้ไขช่อง input
['title', 'presented', 'reason'].forEach(id => {
  document.getElementById(id).addEventListener('input', () => {
    updateDisplay();
  });
});

// เรียกใช้งานครั้งแรก
updateDisplay();
updatePageNum();
