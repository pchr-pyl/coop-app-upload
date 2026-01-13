/**
 * การตั้งค่า ID ของคุณ
 */
const SPREADSHEET_ID = "1otK4BtrGmIbkbqYA2RyZ3iRnaV24KbsNSsG-z1S06Uc";
const FOLDER_ID = "1N9CQoPSCHD1q3DT4--04OwDl5uC4jHPP";

const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
const fileSheet = ss.getSheetByName("Files") || ss.insertSheet("Files");

/**
 * doGet - API สำหรับดึงข้อมูล Mapping
 * รองรับการเรียกจากเว็บภายนอก (Vercel)
 */
function doGet(e) {
  // ตรวจสอบ e และ e.parameter ให้ปลอดภัย
  const action = (e && e.parameter && e.parameter.action) ? e.parameter.action : 'mapping';
  let result;
  
  if (action === 'mapping') {
    result = getMappingData();
  } else {
    result = { error: 'Unknown action' };
  }
  
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * doPost - API สำหรับรับไฟล์จาก Frontend
 * รองรับการเรียกจากเว็บภายนอก (Vercel)
 */
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const result = uploadFile(data);
    
    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ 
        status: 'error', 
        message: 'เกิดข้อผิดพลาด: ' + error.toString() 
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * ฟังก์ชันสร้างโครงสร้างโฟลเดอร์ตามสำนักวิชาและหลักสูตร
 * รันฟังก์ชันนี้จาก Apps Script Editor เพื่อเตรียมโฟลเดอร์ใน Drive
 */
function setupDriveFolders() {
  const mainFolder = DriveApp.getFolderById(FOLDER_ID);
  const data = getMappingData();
  
  Object.keys(data).forEach(facultyName => {
    let facultyFolder;
    const facultyFolders = mainFolder.getFoldersByName(facultyName);
    
    if (facultyFolders.hasNext()) {
      facultyFolder = facultyFolders.next();
    } else {
      facultyFolder = mainFolder.createFolder(facultyName);
    }
    
    data[facultyName].forEach(courseName => {
      const courseFolders = facultyFolder.getFoldersByName(courseName);
      if (!courseFolders.hasNext()) {
        facultyFolder.createFolder(courseName);
      }
    });
  });
  return "สร้างโครงสร้างโฟลเดอร์สำเร็จ";
}

/**
 * ดึงข้อมูล Mapping คณะและหลักสูตร
 */
function getMappingData() {
  return {
    "สำนักวิชาการจัดการ": [
      "บริหารธุรกิจ (การตลาดดิจิทัลและการสร้างแบรนด์)",
      "บริหารธุรกิจ (การจัดการโลจิสติกส์)",
      "การจัดการการท่องเที่ยวและการบริการยุคดิจิทัล",
      "ศิลปะการประกอบอาหารอย่างมืออาชีพ",
      "อุตสาหกรรมการบริการ"
    ],
    "สำนักวิชาการบัญชีและการเงิน": [
      "บริหารธุรกิจ (การจัดการธุรกิจการเงินดิจิทัล)",
      "บัญชี",
      "เศรษฐศาสตร์"
    ],
    "สำนักวิชาเทคโนโลยีการเกษตรและอุตสาหกรรมอาหาร": [
      "เกษตรศาสตร์และนวัตกรรม (พืชศาสตร์)",
      "เกษตรศาสตร์และนวัตกรรม (สัตวศาสตร์)",
      "วิทยาศาสตร์อาหารและนวัตกรรม"
    ],
    "สำนักวิชานิติศาสตร์": ["นิติศาสตร์"],
    "สำนักวิชาแพทยศาสตร์": ["วิทยาศาสตร์การกีฬาและการออกกําลังกาย"],
    "สำนักวิชารัฐศาสตร์และรัฐประศาสนศาสตร์": [
      "รัฐประศาสนศาสตร์",
      "รัฐศาสตร์-การเมืองการปกครอง",
      "รัฐศาสตร์-ความสัมพันธ์ระหว่างประเทศ",
      "รัฐศาสตร์-อาเซียนศึกษา"
    ],
    "สำนักวิชาวิทยาศาสตร์": [
      "วิทยาศาสตร์-คณิตศาสตร์และสถิติ",
      "วิทยาศาสตร์-เคมี",
      "วิทยาศาสตร์-ชีววิทยา",
      "วิทยาศาสตร์ทางทะเล",
      "วิทยาศาสตร์-ฟิสิกส์"
    ],
    "สำนักวิชาวิศวกรรมศาสตร์และเทคโนโลยี": [
      "วิศวกรรมปิโตรเคมีและพอลิเมอร์",
      "วิศวกรรมคอมพิวเตอร์และปัญญาประดิษฐ์",
      "วิศวกรรมคอมพิวเตอร์และระบบอัจฉริยะ",
      "วิศวกรรมเคมีและเคมีเภสัชกรรม",
      "วิศวกรรมเครื่องกลและหุ่นยนต์",
      "วิศวกรรมไฟฟ้า",
      "วิศวกรรมโยธา"
    ],
    "สำนักวิชาศิลปศาสตร์": ["ภาษาจีน", "ภาษาไทย", "ภาษาอังกฤษ"],
    "สำนักวิชาสถาปัตยกรรมศาสตร์และการออกแบบ": ["การออกแบบภายใน", "สถาปัตยกรรม"],
    "สำนักวิชาสาธารณสุขศาสตร์": ["อนามัยสิ่งแวดล้อม", "อาชีวอนามัยและความปลอดภัย"],
    "สำนักวิชาสารสนเทศศาสตร์": [
      "ดิจิทัลคอนเทนต์และสื่อ",
      "เทคโนโลยีมัลติมีเดีย แอนิเมชัน และเกม",
      "เทคโนโลยีสารสนเทศและนวัตกรรมดิจิทัล",
      "นวัตกรรมสารสนเทศศาสตร์ทางการแพทย์",
      "นิเทศศาสตร์ดิจิทัล"
    ]
  };
}

/**
 * อัปโหลดไฟล์ไปยัง Sub-folder ของคณะและหลักสูตร
 */
function uploadFile(data) {
  try {
    const { studentId, studentName, mimeType, fileData, courseName, facultyName } = data;
    const newFileName = `ใบสมัครงาน_${studentName.replace(/\s+/g, '_')}.pdf`;

    const mainFolder = DriveApp.getFolderById(FOLDER_ID);
    
    // ค้นหาหรือสร้างโฟลเดอร์คณะ
    let facultyFolder;
    const facultyFolders = mainFolder.getFoldersByName(facultyName);
    facultyFolder = facultyFolders.hasNext() ? facultyFolders.next() : mainFolder.createFolder(facultyName);

    // ค้นหาหรือสร้างโฟลเดอร์หลักสูตร ภายใต้คณะ
    let courseFolder;
    const courseFolders = facultyFolder.getFoldersByName(courseName);
    courseFolder = courseFolders.hasNext() ? courseFolders.next() : facultyFolder.createFolder(courseName);

    const decodedData = Utilities.base64Decode(fileData.split(',')[1]);
    const blob = Utilities.newBlob(decodedData, mimeType, newFileName);

    // ลบไฟล์เก่า (ถ้ามี)
    const fileDataSheet = fileSheet.getDataRange().getValues();
    const fileHeader = fileDataSheet[0];
    const studentIdCol = fileHeader.indexOf("StudentID");
    const fileIdCol = fileHeader.indexOf("FileID");
    
    let existingRow = -1;
    if (studentIdCol !== -1 && fileIdCol !== -1) {
      for (let i = 1; i < fileDataSheet.length; i++) {
        if (fileDataSheet[i][studentIdCol] == studentId) {
          existingRow = i + 1;
          const oldFileId = fileDataSheet[i][fileIdCol];
          if (oldFileId) { 
            try { DriveApp.getFileById(oldFileId).setTrashed(true); } catch(e) {} 
          }
          break;
        }
      }
    }

    const newFile = courseFolder.createFile(blob);
    const timestamp = new Date().toLocaleString('th-TH', { timeZone: 'Asia/Bangkok' });
    const fileInfo = [Utilities.getUuid(), studentId, studentName, facultyName, courseName, newFile.getId(), newFileName, newFile.getUrl(), timestamp];

    // ตรวจสอบและสร้าง Header ถ้ายังไม่มี
    if (fileSheet.getLastRow() === 0) {
      fileSheet.appendRow(["UUID", "StudentID", "StudentName", "Faculty", "Course", "FileID", "FileName", "FileURL", "Timestamp"]);
    }

    if (existingRow !== -1) {
      fileSheet.getRange(existingRow, 1, 1, 9).setValues([fileInfo]);
    } else {
      fileSheet.appendRow(fileInfo);
    }

    return { status: 'success', message: 'ส่งเอกสารสำเร็จและบันทึกในโฟลเดอร์หลักสูตรแล้ว' };
  } catch (error) {
    return { status: 'error', message: 'เกิดข้อผิดพลาด: ' + error.toString() };
  }
}
