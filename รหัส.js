const MONTHLY_REPORT_SHEET_TEMPLATE_ID = '10GEYj0Hl1Cmeom-x4Dq9vd-Fu4rL7TDM9n7B8B84oyw'; // <--- แก้ไขตรงนี.  
const ATTENDANCE_REPORT_SHEET_NAME = 'Attendance Report Template'; // ชื่อชีตสำหรับรายงานเวลาเรียน
const SCORE_REPORT_SHEET_NAME = 'Score Report Template';       // ชื่อชีตสำหรับรายงานคะแนน
const CHARACTERISTICS_REPORT_SHEET_NAME = 'Characteristics Report Template'; // ชื่อชีตสำหรับรายงานคุณลักษณะฯ
const ACTIVITIES_REPORT_SHEET_NAME = 'Activities Report Template'; // ชื่อชีตสำหรับกิจกรรมพัฒนาผู้เรียน

const VALUES_COMPETENCIES_ITEMS_SHEET_NAME = 'Values Competencies Items Template'; // ชื่อชีตสำหรับรายการค่านิยมและสมรรถนะ
const VALUES_COMPETENCIES_SCORES_SHEET_NAME = 'Values Competencies Scores'; // ชื่อชีตสำหรับคะแนนค่านิยมและสมรรถนะ
const P5_COVER_SHEET_NAME = 'P.5 Cover Template'; // ชื่อชีตสำหรับปก ปพ.5
const VALUES_COMPETENCIES_REPORT_SHEET_NAME = 'Values Competencies Report Template'; // ชื่อชีตสำหรับรายงานค่านิยมและสมรรถนะ


const THEME_MAP = {
  '#FF69B4': { light: '#fbcfe8' }, // Pink
  '#2196F3': { light: '#bbdefb' }, // Blue
  '#3F51B5': { light: '#c5cae9' }, // Navy
  '#4CAF50': { light: '#c8e6c9' }, // Green
  '#FFC107': { light: '#fff9c4' }, // Yellow
  '#9C27B0': { light: '#e1bee7' }, // Purple
  '#FF9800': { light: '#ffe0b2' }, // Orange
  '#F44336': { light: '#ffcdd2' }  // Red
};

function getImageAsBase64(url) {
    try {
        // ดึงข้อมูลรูปภาพจาก URL
        var blob = UrlFetchApp.fetch(url).getBlob();
        var contentType = blob.getContentType();
        // แปลงข้อมูลเป็น Base64
        var base64Data = Utilities.base64Encode(blob.getBytes());
        // คืนค่าในรูปแบบที่ <img> tag สามารถใช้งานได้
        return "data:" + contentType + ";base64," + base64Data;
    } catch (e) {
        Logger.log('Could not fetch or encode image from URL: ' + url + '. Error: ' + e.toString());
        return ''; // คืนค่าว่างหากเกิดข้อผิดพลาด
    }
}

function doGet(e) {
  Logger.log('doGet called with params: ' + JSON.stringify(e));
  try {
    var htmlContent = loadLoginPage();
    return HtmlService
      .createHtmlOutput(htmlContent)
      .setTitle('ล็อกอิน - ระบบเช็คชื่อนักเรียนออนไลน์')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (error) {
    Logger.log('Error in doGet: ' + error.stack);
    throw new Error('ไม่สามารถโหลดหน้า login ได้: ' + error.message);
  }
}

function initializeSheet(sheetName) {
    Logger.log('initializeSheet called: ' + sheetName);
    try {
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var sheet = ss.getSheetByName(sheetName);
        if (!sheet) {
            sheet = ss.insertSheet(sheetName);
            Logger.log('Created new sheet: ' + sheetName);
        }
        var headers = [];
        if (sheetName === 'subjects') {
            headers = ['id', 'code', 'name', 'credits'];
        } else if (sheetName === 'teachers') {
            headers = ['id', 'username', 'password', 'name', 'subjectIds', 'classLevels', 'resetRequested', 'subjectClassPairs'];
        } else if (sheetName === 'students') {
            headers = ['id', 'code', 'name', 'class', 'classroom'];
        } else if (sheetName === 'attendance') {
            headers = ['id', 'studentId', 'subjectId', 'date', 'status', 'studentName', 'class', 'classroom', 'teacherName', 'subjectName', 'remark'];
        } else if (sheetName === 'settings') {
            headers = ['key', 'value'];
        } else if (sheetName === 'score_components') {
            headers = ['id', 'teacherId', 'subjectId', 'classLevel', 'classroom', 'componentName', 'maxScore'];
        } else if (sheetName === 'scores') {
        // --- START: ส่วนที่แก้ไข ---
        // เพิ่ม subjectId เพื่อระบุว่าหมายเหตุ (ร, มผ, มส) เป็นของวิชาใด
        headers = ['id', 'studentId', 'componentId', 'score', 'dateModified', 'remark', 'subjectId'];
        // --- END: ส่วนที่แก้ไข ---
        } else if (sheetName === 'characteristics_items') {
            headers = ['id', 'itemName', 'itemGroup', 'order'];
        } else if (sheetName === 'characteristics_scores') {
            headers = ['id', 'studentId', 'itemId', 'score', 'teacherId', 'dateModified'];
        } else if (sheetName === 'activity_components') {
            headers = ['id', 'teacherId', 'classLevel', 'classroom', 'componentName'];
        } else if (sheetName === 'activity_scores') {
            headers = ['id', 'studentId', 'componentId', 'result', 'teacherId', 'dateModified'];
        } else if (sheetName === 'values_competencies_items') {
            headers = ['id', 'itemName', 'itemGroup', 'order'];
        } else if (sheetName === 'values_competencies_scores') {
            headers = ['id', 'studentId', 'itemId', 'score', 'teacherId', 'dateModified'];
        }

        if (headers.length > 0) {
            var range = sheet.getRange(1, 1, 1, headers.length);
            if (sheet.getLastRow() === 0 || !range.getValues()[0].every((val, i) => val === headers[i])) {
                range.setValues([headers]);
                Logger.log('Set headers for sheet: ' + sheetName);

                if (sheetName === 'students') {
                    sheet.getRange("E2:E").setNumberFormat('@');
                    Logger.log('Set "classroom" column (E) to Plain Text format.');
                } else if (sheetName === 'settings') {
                    sheet.appendRow(['system_name', 'ระบบเช็คชื่อนักเรียนออนไลน์ โรงเรียนห้องสื่อครูคอม']);
                    sheet.appendRow(['logo_url', 'https://img2.pic.in.th/jsz.png']);
                    sheet.appendRow(['header_text', 'ระบบเช็คชื่อนักเรียนออนไลน์ โรงเรียนห้องสื่อครูคอม']);
                    sheet.appendRow(['footer_text', '© 2025 ระบบเช็คชื่อนักเรียนออนไลน์ โรงเรียนห้องสื่อครูคอม | กฤษฎา คำมา']);
                    sheet.appendRow(['theme_color', '#FF69B4']);
                    sheet.appendRow(['school_name', 'โรงเรียนห้องสื่อครูคอม']);
                    sheet.appendRow(['director_name', 'ชื่อผู้อำนวยการ']);
                    sheet.appendRow(['school_area', 'สำนักงานเขตพื้นที่การศึกษาประถมศึกษาขอนแก่น เขต 5']);
                } else if (sheetName === 'characteristics_items') {
                    var defaultItems = [
                        ['1. รักชาติ ศาสน์ กษัตริย์', 'คุณลักษณะอันพึงประสงค์', 1],
                        ['2. ซื่อสัตย์สุจริต', 'คุณลักษณะอันพึงประสงค์', 2],
                        ['3. มีวินัย', 'คุณลักษณะอันพึงประสงค์', 3],
                        ['4. ใฝ่เรียนรู้', 'คุณลักษณะอันพึงประสงค์', 4],
                        ['5. อยู่อย่างพอเพียง', 'คุณลักษณะอันพึงประสงค์', 5],
                        ['6. มุ่งมั่นในการทำงาน', 'คุณลักษณะอันพึงประสงค์', 6],
                        ['7. รักความเป็นไทย', 'คุณลักษณะอันพึงประสงค์', 7],
                        ['8. มีจิตสาธารณะ', 'คุณลักษณะอันพึงประสงค์', 8],
                        ['9. การอ่าน', 'การอ่าน คิดวิเคราะห์ และเขียน', 9],
                        ['10. การคิดวิเคราะห์', 'การอ่าน คิดวิเคราะห์ และเขียน', 10],
                        ['11. การเขียน', 'การอ่าน คิดวิเคราะห์ และเขียน', 11]
                    ];
                    var itemsWithIds = defaultItems.map(function(item) {
                        return [Utilities.getUuid()].concat(item);
                    });
                    if (itemsWithIds.length > 0) {
                        sheet.getRange(2, 1, itemsWithIds.length, 4).setValues(itemsWithIds);
                    }
                } else if (sheetName === 'values_competencies_items') {
                    var defaultItems = [
                        ['1. มีความรักชาติ ศาสนา พระมหากษัตริย์', 'ค่านิยมหลักของคนไทย 12 ประการ', 1],
                        ['2. ซื่อสัตย์ เสียสละ อดทน มีอุดมการณ์ในสิ่งที่ดีงามเพื่อส่วนรวม', 'ค่านิยมหลักของคนไทย 12 ประการ', 2],
                        ['3. กตัญญูต่อพ่อแม่ ผู้ปกครอง ครูบาอาจารย์', 'ค่านิยมหลักของคนไทย 12 ประการ', 3],
                        ['4. ใฝ่หาความรู้ หมั่นศึกษาเล่าเรียนทั้งทางตรง และทางอ้อม', 'ค่านิยมหลักของคนไทย 12 ประการ', 4],
                        ['5. รักษาวัฒนธรรมประเพณีไทยอันงดงาม', 'ค่านิยมหลักของคนไทย 12 ประการ', 5],
                        ['6. มีศีลธรรม รักษาความสัตย์ หวังดีต่อผู้อื่น เผื่อแผ่และแบ่งปัน', 'ค่านิยมหลักของคนไทย 12 ประการ', 6],
                        ['7. เข้าใจเรียนรู้การเป็นประชาธิปไตย อันมีพระมหากษัตริย์ทรงเป็นประมุขที่ถูกต้อง', 'ค่านิยมหลักของคนไทย 12 ประการ', 7],
                        ['8. มีระเบียบ วินัย เคารพกฎหมาย ผู้น้อยรู้จักการเคารพผู้ใหญ่', 'ค่านิยมหลักของคนไทย 12 ประการ', 8],
                        ['9. มีสติรู้ตัว รู้คิด รู้ทำ รู้ปฏิบัติ ตามพระราชดำรัสของพระบาทสมเด็จพระเจ้าอยู่หัว', 'ค่านิยมหลักของคนไทย 12 ประการ', 9],
                        ['10. รู้จักดำรงตนอยู่โดยใช้หลักปรัชญาเศรษฐกิจพอเพียง', 'ค่านิยมหลักของคนไทย 12 ประการ', 10],
                        ['11. มีความเข้มแข็งทั้งร่างกาย และจิตใจ ไม่ยอมแพ้ต่ออำนาจฝ่ายต่ำ', 'ค่านิยมหลักของคนไทย 12 ประการ', 11],
                        ['12. คำนึงถึงผลประโยชน์ของส่วนรวม และของชาติมากกว่าผลประโยชน์ของตนเอง', 'ค่านิยมหลักของคนไทย 12 ประการ', 12],
                        ['1. ความสามารถในการสื่อสาร', 'สมรรถนะสำคัญของผู้เรียน', 13],
                        ['2. ความสามารถในการคิด', 'สมรรถนะสำคัญของผู้เรียน', 14],
                        ['3. ความสามารถในการแก้ปัญหา', 'สมรรถนะสำคัญของผู้เรียน', 15],
                        ['4. ความสามารถในการใช้ทักษะชีวิต', 'สมรรถนะสำคัญของผู้เรียน', 16],
                        ['5. ความสามารถในการใช้เทคโนโลยี', 'สมรรถนะสำคัญของผู้เรียน', 17]
                    ];
                    var itemsWithIds = defaultItems.map(function(item) {
                        return [Utilities.getUuid()].concat(item);
                    });
                    if (itemsWithIds.length > 0) {
                        sheet.getRange(2, 1, itemsWithIds.length, 4).setValues(itemsWithIds);
                    }
                }
            }
        }
        return sheet;
    } catch (error) {
        Logger.log('Error in initializeSheet for ' + sheetName + ': ' + error.stack);
        throw new Error('ไม่สามารถสร้างหรือเข้าถึง Sheet ได้: ' + sheetName);
    }
}


function getInitialData() {
  Logger.log('getInitialData: Bundling all necessary data for the client.');
  try {
    var user = {};
    // ในสถานการณ์จริง ควรมีการตรวจสอบ Session หรือ Token เพื่อระบุผู้ใช้
    // แต่ในโครงสร้างนี้ เราจะส่งข้อมูลทั้งหมดที่จำเป็นไปก่อน
    
    var dataBundle = {
      settings: getSettings(),
      classLevelSettings: getClassLevelSettings(),
      subjects: getData('subjects'),
      students: getData('students'),
      teachers: getData('teachers'),
      attendance: getData('attendance')
    };
    Logger.log('getInitialData: Successfully bundled all data.');
    return dataBundle;
  } catch (error) {
    Logger.log('Error in getInitialData: ' + error.stack);
    throw new Error('ไม่สามารถรวบรวมข้อมูลเริ่มต้นของระบบได้: ' + error.message);
  }
}

function getInitialDataForUser(user) {
  Logger.log('getInitialDataForUser called for role: ' + user.role);
  try {
      var dataBundle = {
          settings: getSettings(),
          classLevelSettings: getClassLevelSettings(),
          characteristicsItems: getData('characteristics_items'),
          valuesCompetenciesItems: getData('values_competencies_items'),
        activityComponents: getData('activity_components')
      };

      if (user.role === 'admin') {
        // สำหรับ Admin ดึงข้อมูลทั้งหมดเหมือนเดิม ไม่มีการเปลี่ยนแปลง
          dataBundle.subjects = getData('subjects');
          dataBundle.students = getData('students');
          dataBundle.teachers = getData('teachers');
          dataBundle.attendance = getData('attendance');
          dataBundle.scoreComponents = getData('score_components');
          dataBundle.scores = getData('scores');
          dataBundle.characteristicsScores = getData('characteristics_scores');
          dataBundle.activityScores = getData('activity_scores');
          dataBundle.valuesCompetenciesScores = getData('values_competencies_scores');

      } else if (user.role === 'teacher') {
          var pairs = [];
          try {
              pairs = JSON.parse(user.subjectClassPairs || '[]');
          } catch (e) {
              Logger.log('Invalid subjectClassPairs for teacher: ' + user.id);
          }

          // กำหนดขอบเขตของครู (วิชา, ระดับชั้น, ห้องเรียน)
          var validSubjectIds = [...new Set(pairs.map(p => p.subjectId))];
          var validClasses = [...new Set(pairs.flatMap(p => p.classLevels || []))];
          var validClassrooms = [...new Set(pairs.flatMap(p => p.classrooms || []))];

        // ดึงข้อมูลวิชาและนักเรียนตามขอบเขต (เหมือนเดิม)
          dataBundle.subjects = getData('subjects').filter(s => validSubjectIds.includes(s.id));
          dataBundle.students = getData('students').filter(student => {
              return validClasses.includes(student.class) && validClassrooms.includes(student.classroom || '');
          });
          var studentIdsInScope = dataBundle.students.map(s => s.id);
        dataBundle.teachers = []; // ครูไม่ต้องเห็นข้อมูลครูคนอื่น

        // -----[ START: ส่วนที่แก้ไข ]-----

        // 1. แก้ไขการดึงข้อมูล "การเข้าเรียน"
        // จากเดิม: กรองตาม teacherName ทำให้เห็นเฉพาะที่ตัวเองบันทึก
        // ใหม่: ไม่กรองตาม teacherName ทำให้เห็นข้อมูลของทุกคนในวิชา/ห้องเรียนที่ตนสอน
          dataBundle.attendance = getData('attendance').filter(a =>
              validClasses.includes(a.class) &&
              validClassrooms.includes(a.classroom || '') &&
              validSubjectIds.includes(String(a.subjectId))
            // ** บรรทัดที่ถูกลบออก: && a.teacherName === user.name **
          );

        // 2. แก้ไขการดึงข้อมูล "องค์ประกอบคะแนน"
        // จากเดิม: กรองตาม teacherId ทำให้เห็นเฉพาะโครงสร้างคะแนนที่ตัวเองสร้าง
        // ใหม่: กรองตามขอบเขตวิชา/ห้องเรียน ทำให้เห็นโครงสร้างคะแนนของทุกคนที่สอนวิชาเดียวกัน
          dataBundle.scoreComponents = getData('score_components').filter(c => 
            validSubjectIds.includes(c.subjectId) &&
            validClasses.includes(c.classLevel) &&
            validClassrooms.includes(c.classroom)
        );
        
        // 3. การดึง "คะแนน" (scores) ถูกต้องอยู่แล้ว เพราะกรองตามนักเรียนในขอบเขต (studentIdsInScope)
          dataBundle.scores = getData('scores').filter(s => studentIdsInScope.includes(s.studentId));
        
        // 4. แก้ไขการดึงข้อมูล "คะแนนคุณลักษณะฯ"
        // จากเดิม: กรองตาม teacherId ทำให้เห็นเฉพาะที่ตัวเองประเมิน
        // ใหม่: กรองตามนักเรียนในห้องที่ตนสอน ทำให้เห็นผลประเมินของทุกคน
          dataBundle.characteristicsScores = getData('characteristics_scores').filter(s => studentIdsInScope.includes(s.studentId));
        
        // 5. แก้ไขการดึงข้อมูล "องค์ประกอบกิจกรรมฯ"
        // จากเดิม: กรองตาม teacherId
        // ใหม่: กรองตามห้องเรียน ทำให้เห็นกิจกรรมของห้องนั้นๆ ทั้งหมด
        dataBundle.activityComponents = getData('activity_components').filter(c => 
            validClasses.includes(c.classLevel) &&
            validClassrooms.includes(c.classroom)
        );

        // 6. แก้ไขการดึงข้อมูล "ผลกิจกรรมฯ"
        // จากเดิม: กรองตาม teacherId
        // ใหม่: กรองตามนักเรียนในห้องที่ตนสอน ทำให้เห็นผลประเมินของทุกคน
          dataBundle.activityScores = getData('activity_scores').filter(s => studentIdsInScope.includes(s.studentId));

        // 7. แก้ไขการดึงข้อมูล "คะแนนค่านิยมและสมรรถนะ"
        // จากเดิม: กรองตาม teacherId
        // ใหม่: กรองตามนักเรียนในห้องที่ตนสอน ทำให้เห็นผลประเมินของทุกคน
          dataBundle.valuesCompetenciesScores = getData('values_competencies_scores').filter(s => studentIdsInScope.includes(s.studentId));

        // -----[ END: ส่วนที่แก้ไข ]-----

      } else if (user.role === 'student') {
        // สำหรับนักเรียน ดึงข้อมูลเหมือนเดิม ไม่มีการเปลี่ยนแปลง
          var allSubjects = getData('subjects');
          var allScoreComponents = getData('score_components');
          
          dataBundle.attendance = getData('attendance').filter(a => a.studentId === user.id);
          var studentScores = getData('scores').filter(s => s.studentId === user.id);
          dataBundle.scores = studentScores;
          dataBundle.characteristicsScores = getData('characteristics_scores').filter(s => s.studentId === user.id);
          dataBundle.activityScores = getData('activity_scores').filter(s => s.studentId === user.id);
          dataBundle.valuesCompetenciesScores = getData('values_competencies_scores').filter(s => s.studentId === user.id);

          var relevantComponentIds = [...new Set(studentScores.map(s => s.componentId))];
          dataBundle.scoreComponents = allScoreComponents.filter(c => relevantComponentIds.includes(c.id));
          
          var relevantSubjectIds = new Set();
          dataBundle.scoreComponents.forEach(comp => relevantSubjectIds.add(comp.subjectId));
          dataBundle.attendance.forEach(att => relevantSubjectIds.add(att.subjectId));

          dataBundle.subjects = allSubjects.filter(s => relevantSubjectIds.has(s.id));
          dataBundle.teachers = getData('teachers');
      }

      Logger.log('getInitialDataForUser: Successfully bundled data for ' + user.role);
      return dataBundle;
  } catch (error) {
      Logger.log('Error in getInitialDataForUser: ' + error.stack);
      throw new Error('ไม่สามารถรวบรวมข้อมูลได้: ' + error.message);
  }
}

function initializeClassLevelsSheet() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetName = 'class_levels_config';
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
        sheet = ss.insertSheet(sheetName);
        var headers = ['groupName', 'levelName'];
        sheet.getRange(1, 1, 1, 2).setValues([headers]);
        var levels = [
            ['ระดับปฐมวัย', 'เตรียมอนุบาล'], ['ระดับปฐมวัย', 'อนุบาล 1'], ['ระดับปฐมวัย', 'อนุบาล 2'], ['ระดับปฐมวัย', 'อนุบาล 3'],
            ['ระดับประถมศึกษาตอนต้น', 'ประถมศึกษาปีที่ 1'], ['ระดับประถมศึกษาตอนต้น', 'ประถมศึกษาปีที่ 2'], ['ระดับประถมศึกษาตอนต้น', 'ประถมศึกษาปีที่ 3'],
            ['ระดับประถมศึกษาตอนปลาย', 'ประถมศึกษาปีที่ 4'], ['ระดับประถมศึกษาตอนปลาย', 'ประถมศึกษาปีที่ 5'], ['ระดับประถมศึกษาตอนปลาย', 'ประถมศึกษาปีที่ 6'],
            ['ระดับมัธยมศึกษาตอนต้น', 'มัธยมศึกษาปีที่ 1'], ['ระดับมัธยมศึกษาตอนต้น', 'มัธยมศึกษาปีที่ 2'], ['ระดับมัธยมศึกษาตอนต้น', 'มัธยมศึกษาปีที่ 3'],
            ['ระดับมัธยมศึกษาตอนปลาย', 'มัธยมศึกษาปีที่ 4'], ['ระดับมัธยมศึกษาตอนปลาย', 'มัธยมศึกษาปีที่ 5'], ['ระดับมัธยมศึกษาตอนปลาย', 'มัธยมศึกษาปีที่ 6'],
            ['ระดับอาชีวศึกษา', 'ปวช. ปี 1'], ['ระดับอาชีวศึกษา', 'ปวช. ปี 2'], ['ระดับอาชีวศึกษา', 'ปวช. ปี 3'], ['ระดับอาชีวศึกษา', 'ปวส. ปี 1'], ['ระดับอาชีวศึกษา', 'ปวส. ปี 2'],
            ['ระดับอุดมศึกษา', 'ปี 1'], ['ระดับอุดมศึกษา', 'ปี 2'], ['ระดับอุดมศึกษา', 'ปี 3'], ['ระดับอุดมศึกษา', 'ปี 4'], ['ระดับอุดมศึกษา', 'ปี 5'], ['ระดับอุดมศึกษา', 'ปี 6']
        ];
        sheet.getRange(2, 1, levels.length, 2).setValues(levels);
    }
    return sheet;
}

function getClassLevelSettings() {
    try {
        var configSheet = initializeClassLevelsSheet();
        var allLevelsData = configSheet.getDataRange().getValues();
        var settings = getSettings();

        var allLevelsGrouped = {};
        for (var i = 1; i < allLevelsData.length; i++) {
            var group = allLevelsData[i][0];
            var level = allLevelsData[i][1];
            if (!allLevelsGrouped[group]) {
                allLevelsGrouped[group] = [];
            }
            allLevelsGrouped[group].push(level);
        }

        var enabledLevels = [];
        // ตรวจสอบว่ามีค่า setting นี้อยู่หรือไม่ และเป็น JSON ที่ถูกต้องหรือไม่
        if (settings.enabled_class_levels) {
            try {
                var parsedLevels = JSON.parse(settings.enabled_class_levels);
                if (Array.isArray(parsedLevels)) {
                     enabledLevels = parsedLevels;
                }
            } catch (e) {
                Logger.log('Could not parse enabled_class_levels, defaulting to empty. Value was: ' + settings.enabled_class_levels);
                enabledLevels = []; // หาก parse ไม่ได้ ให้ใช้ค่าว่าง
            }
        }

        return {
            allSettings: settings,
            allLevels: allLevelsGrouped,
            enabledLevels: enabledLevels
        };
    } catch (e) {
        Logger.log('Error in getClassLevelSettings: ' + e.stack);
        throw new Error('ไม่สามารถดึงข้อมูลการตั้งค่าระดับชั้นได้');
    }
}

function login(username, password) {
  Logger.log('login called for teacher/admin: ' + username);
  try {
    username = username ? username.trim() : '';
    password = password ? password.trim() : '';

    if (username === 'admin' && password === 'admin1234') {
      Logger.log('Admin login successful');
      return { id: 'admin', name: 'Administrator', role: 'admin' };
    }

    var sheet = initializeSheet('teachers');
    var data = sheet.getDataRange().getValues();
    Logger.log('Teachers sheet data length: ' + data.length);

    if (data.length <= 1) {
      Logger.log('Login failed: No teacher data in sheet');
      throw new Error('ไม่มีข้อมูลครูในระบบ กรุณาติดต่อผู้ดูแลระบบ');
    }

    for (var i = 1; i < data.length; i++) {
      if (data[i].length < 8) {
        Logger.log('Invalid row format at index ' + i + ': ' + JSON.stringify(data[i]));
        continue;
      }
      var sheetUsername = data[i][1] ? data[i][1].toString().trim() : '';
      var sheetPassword = data[i][2] ? data[i][2].toString().trim() : '';
      if (sheetUsername === username && sheetPassword === password) {
        var subjectClassPairs = data[i][7] ? data[i][7].toString().trim() : '[]';
        var parsedPairs = [];
        try {
          parsedPairs = JSON.parse(subjectClassPairs);
        } catch (e) {
          Logger.log('Invalid subjectClassPairs for user ' + username + ': ' + subjectClassPairs);
        }
        var classrooms = parsedPairs.flatMap(p => p.classrooms || []);
        Logger.log('Teacher login successful: ' + username);
        return {
          id: data[i][0],
          name: data[i][3] || 'Unknown',
          role: 'teacher',
          subjectIds: data[i][4] ? data[i][4].toString().split(',') : [],
          classLevels: data[i][5] ? data[i][5].toString().split(',') : [],
          subjectClassPairs: subjectClassPairs,
          classrooms: classrooms
        };
      }
    }
    Logger.log('Login failed: Invalid credentials for username: ' + username);
    throw new Error('ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง');
  } catch (error) {
    Logger.log('Error in login: ' + error.message);
    throw new Error(error.message || 'การล็อกอินล้มเหลว กรุณาลองใหม่');
  }
}


function loadIndexPage() {
  Logger.log('loadIndexPage called');
  try {
    var settings = getSettings();
    var htmlContent = HtmlService.createHtmlOutputFromFile('index').getContent();

    var placeholders = {
      '{{SYSTEM_NAME}}': settings.system_name || 'ระบบเช็คชื่อนักเรียนออนไลน์ โรงเรียนห้องสื่อครูคอม',
      '{{LOGO_URL}}': settings.logo_url || 'https://img2.pic.in.th/jsz.png',
      '{{HEADER_TEXT}}': settings.header_text || 'ระบบเช็คชื่อนักเรียนออนไลน์ โรงเรียนห้องสื่อครูคอม',
      '{{FOOTER_TEXT}}': settings.footer_text || '© 2025 ระบบเช็คชื่อนักเรียนออนไลน์ โรงเรียนห้องสื่อครูคอม | กฤษฎา คำมา',
      '{{THEME_COLOR}}': settings.theme_color || '#FF69B4' // เพิ่ม placeholder สำหรับสี
    };

    Object.keys(placeholders).forEach(function(key) {
      htmlContent = htmlContent.replace(new RegExp(key, 'g'), placeholders[key]);
    });

    if (!htmlContent || htmlContent.trim() === '') {
      throw new Error('เนื้อหา HTML ว่างเปล่า');
    }

    Logger.log('Index page loaded successfully');
    return htmlContent;
  } catch (error) {
    Logger.log('Error in loadIndexPage: ' + error.stack);
    throw new Error('ไม่สามารถโหลดหน้า index ได้: ' + error.message);
  }
}

function loadLoginPage() {
    Logger.log('loadLoginPage called');
    try {
        var settings = getSettings();
        var html = HtmlService.createHtmlOutputFromFile('login').getContent();

        var primaryColor = settings.theme_color || '#FF69B4';
        var lightColor = THEME_MAP[primaryColor] ? THEME_MAP[primaryColor].light : '#fbcfe8';

        html = html.replace(/{{SYSTEM_NAME}}/g, settings.system_name || 'ระบบเช็คชื่อนักเรียนออนไลน์ โรงเรียนห้องสื่อครูคอม');
        html = html.replace(/{{LOGO_URL}}/g, settings.logo_url || 'https://img2.pic.in.th/jsz.png');
        html = html.replace(/{{HEADER_TEXT}}/g, settings.header_text || 'ระบบเช็คชื่อนักเรียนออนไลน์ โรงเรียนห้องสื่อครูคอม');
        html = html.replace(/{{FOOTER_TEXT}}/g, settings.footer_text || '© 2025 ระบบเช็คชื่อนักเรียนออนไลน์ โรงเรียนห้องสื่อครูคอม | กฤษฎา คำมา');
        html = html.replace(/{{THEME_COLOR}}/g, primaryColor);
        html = html.replace(/{{THEME_LIGHT_COLOR}}/g, lightColor);

        return html;
    } catch (error) {
        Logger.log('Error in loadLoginPage: ' + error.message);
        throw new Error('ไม่สามารถโหลดหน้า login ได้');
    }
}


function loginAndLoad(username, password) {
    Logger.log('loginAndLoad called: ' + username);
    try {
        var user = login(username, password);
        var htmlContent = loadIndexPage();
        var initialData = getInitialDataForUser(user);
        return {
            user: user,
            htmlContent: htmlContent,
            initialData: initialData
        };
    } catch (error) {
        Logger.log('Error in loginAndLoad: ' + error.stack);
        throw new Error(error.message || 'การล็อกอินล้มเหลว กรุณาลองใหม่');
    }
}

function studentLoginAndLoad(studentCode) {
    Logger.log('studentLoginAndLoad called for student code: ' + studentCode);
    try {
        if (!studentCode) {
            throw new Error('กรุณากรอกรหัสนักเรียน');
        }

        var allStudents = getData('students');
        var student = allStudents.find(function(s) {
            return s.code === studentCode.trim();
        });

        if (!student) {
            throw new Error('ไม่พบรหัสนักเรียนนี้ในระบบ');
        }

        // START: สร้าง object ของ user ให้สมบูรณ์
        var user = {
            id: student.id,
            name: student.name,
            role: 'student',
            code: student.code,
            class: student.class,
            classroom: student.classroom
        };
        // END: สร้าง object ของ user ให้สมบูรณ์

        // START: เปลี่ยนมาเรียกใช้ getInitialDataForUser เพื่อให้ได้ข้อมูลครบถ้วน
        var initialData = getInitialDataForUser(user);
        // END: เปลี่ยนมาเรียกใช้ getInitialDataForUser

        var htmlContent = loadIndexPage();

        return {
            user: user,
            htmlContent: htmlContent,
            initialData: initialData
        };

    } catch (error) {
        Logger.log('Error in studentLoginAndLoad: ' + error.stack);
        throw new Error(error.message || 'การล็อกอินล้มเหลว กรุณาลองใหม่');
    }
}

function requestPasswordReset(username) {
  Logger.log('requestPasswordReset called: ' + username);
  try {
    var sheet = initializeSheet('teachers');
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][1] === username) {
        sheet.getRange(i + 1, 7).setValue(true);
        Logger.log('Password reset requested for: ' + username);
        return;
      }
    }
    Logger.log('Password reset failed: Username not found');
    throw new Error('ไม่พบชื่อผู้ใช้');
  } catch (error) {
    Logger.log('Error in requestPasswordReset: ' + error.message);
    throw new Error('ไม่พบชื่อผู้ใช้ การขอรีเซ็ตรหัสผ่านล้มเหลว');
  }
}

function resetTeacherPassword(teacherId, newPassword) {
  Logger.log('resetTeacherPassword called: ' + teacherId);
  try {
    var sheet = initializeSheet('teachers');
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === teacherId) {
        sheet.getRange(i + 1, 3).setValue(newPassword);
        sheet.getRange(i + 1, 7).setValue(false);
        Logger.log('Password reset for teacher: ' + teacherId);
        return;
      }
    }
    Logger.log('Password reset failed: Teacher not found');
    throw new Error('ไม่พบครู');
  } catch (error) {
    Logger.log('Error in resetTeacherPassword: ' + error.message);
    throw new Error('การรีเซ็ตรหัสผ่านล้มเหลว');
  }
}

function getData(sheetName) {
  Logger.log('getData called: ' + sheetName);
  try {
    var sheet = initializeSheet(sheetName);
    var dataRange = sheet.getDataRange();
    var values = dataRange.getValues();
    var result = [];
    
    if (values.length <= 1) {
      if (sheetName === 'characteristics_items') {
        Logger.log('No data found in characteristics_items sheet, returning default items.');
        return [
            {id: 'char-item-1', itemName: '1. รักชาติ ศาสน์ กษัตริย์', itemGroup: 'คุณลักษณะอันพึงประสงค์', order: 1},
            {id: 'char-item-2', itemName: '2. ซื่อสัตย์สุจริต', itemGroup: 'คุณลักษณะอันพึงประสงค์', order: 2},
            {id: 'char-item-3', itemName: '3. มีวินัย', itemGroup: 'คุณลักษณะอันพึงประสงค์', order: 3},
            {id: 'char-item-4', itemName: '4. ใฝ่เรียนรู้', itemGroup: 'คุณลักษณะอันพึงประสงค์', order: 4},
            {id: 'char-item-5', itemName: '5. อยู่อย่างพอเพียง', itemGroup: 'คุณลักษณะอันพึงประสงค์', order: 5},
            {id: 'char-item-6', itemName: '6. มุ่งมั่นในการทำงาน', itemGroup: 'คุณลักษณะอันพึงประสงค์', order: 6},
            {id: 'char-item-7', itemName: '7. รักความเป็นไทย', itemGroup: 'คุณลักษณะอันพึงประสงค์', order: 7},
            {id: 'char-item-8', itemName: '8. มีจิตสาธารณะ', itemGroup: 'คุณลักษณะอันพึงประสงค์', order: 8},
            {id: 'char-item-9', itemName: '1. การอ่าน', itemGroup: 'การอ่าน คิดวิเคราะห์ และเขียน', order: 9},
            {id: 'char-item-10', itemName: '2. การคิดวิเคราะห์', itemGroup: 'การอ่าน คิดวิเคราะห์ และเขียน', order: 10},
            {id: 'char-item-11', itemName: '3. การเขียน', itemGroup: 'การอ่าน คิดวิเคราะห์ และเขียน', order: 11}
        ];
      }
      Logger.log('No data found in sheet: ' + sheetName);
      return [];
    }

    var headers = values[0];
    for (var i = 1; i < values.length; i++) {
      if (!values[i][0]) continue;

      var row = {};
      for (var j = 0; j < headers.length; j++) {
        var headerName = headers[j];
        var cellValue = values[i][j];

        // --- START: เพิ่ม Logic ดักจับ Date ในช่อง Classroom ---
        if (sheetName === 'students' && headerName === 'classroom' && cellValue instanceof Date && !isNaN(cellValue)) {
            // แปลง Date กลับไปเป็นรูปแบบ "เดือน/วัน" ซึ่งตรงกับที่ผู้ใช้พิมพ์
            var month = cellValue.getMonth() + 1;
            var day = cellValue.getDate();
            row[headerName] = `${month}/${day}`;
            Logger.log(`Corrected a Date object in 'classroom' column back to string: '${row[headerName]}'`);
        } else if ((headerName === 'date' || headerName === 'dateModified') && cellValue instanceof Date) {
            row[headerName] = Utilities.formatDate(cellValue, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        } else {
            row[headerName] = cellValue !== null ? cellValue.toString() : '';
        }
        // --- END: เพิ่ม Logic ดักจับ ---
      }
      result.push(row);
    }

    Logger.log('Data retrieved: ' + sheetName + ', count: ' + result.length);
    return result;
  } catch (error) {
    Logger.log('Error in getData: ' + error.message + " Stack: " + error.stack);
    throw new Error('ไม่สามารถดึงข้อมูล ' + sheetName + ' ได้: ' + error.message);
  }
}

function addData(sheetName, data) {
    Logger.log('addData called: ' + sheetName);
    try {
        var sheet = initializeSheet(sheetName);
        if (sheetName === 'teachers') {
            if (!/^[a-zA-Z0-9]{3,}$/.test(data.username)) {
                throw new Error('ชื่อผู้ใช้ต้องประกอบด้วยภาษาอังกฤษและตัวเลขอย่างน้อย 3 ตัวอักษร');
            }
            if (data.password && data.password.length < 6) {
                throw new Error('รหัสผ่านต้องมีอย่างน้อย 6 ตัวอักษร');
            }
            var teachers = sheet.getDataRange().getValues();
            for (var i = 1; i < teachers.length; i++) {
                var existingUsername = teachers[i][1] ? teachers[i][1].toString().trim() : '';
                if (existingUsername === data.username.trim()) {
                    throw new Error('ชื่อผู้ใช้นี้มีอยู่แล้วในระบบ');
                }
            }

            var pairs = JSON.parse(data.subjectClassPairs || '[]');
            if (pairs.length > 0) {
                var allStudents = getData('students');
                var existingStudentCombos = new Set();
                allStudents.forEach(function(s) {
                    if (s.class && s.classroom) {
                        existingStudentCombos.add(s.class + '/' + s.classroom);
                    }
                });

                for (var i = 0; i < pairs.length; i++) {
                    var pair = pairs[i];
                    var classLevels = pair.classLevels || [];
                    var classrooms = pair.classrooms || [];
                    for (var j = 0; j < classLevels.length; j++) {
                        for (var k = 0; k < classrooms.length; k++) {
                            var comboToCheck = classLevels[j] + '/' + classrooms[k];
                            if (!existingStudentCombos.has(comboToCheck)) {
                                throw new Error('บันทึกล้มเหลว: ไม่สามารถมอบหมายให้ดูแลห้องเรียน "' + classrooms[k] + '" ของระดับชั้น "' + classLevels[j] + '" ได้ เนื่องจากยังไม่มีข้อมูลนักเรียนในห้องดังกล่าว');
                            }
                        }
                    }
                }
            }

        } else if (sheetName === 'students') {
            var students = getData('students');
            var codeExists = students.some(s => s.code === data.code);
            if (codeExists) {
                throw new Error('รหัสนักเรียนนี้มีอยู่แล้วในระบบ');
            }
        }
        var id = Utilities.getUuid();
        var row = [id];
        if (sheetName === 'subjects') {
            row.push(data.code || '', data.name || '', data.credits || '');
        } else if (sheetName === 'teachers') {
            row.push(
                data.username || '',
                data.password || '',
                data.name || '',
                data.subjectIds || '',
                data.classLevels || '',
                false,
                data.subjectClassPairs || '[]'
            );
        } else if (sheetName === 'students') {
            row.push(
                data.code || '',
                data.name || '',
                data.class || '',
                data.classroom || ''
            );
        }
        sheet.appendRow(row);
        Logger.log('Data added: ' + sheetName + ', id: ' + id);
        return {
            success: true,
            id: id
        };
    } catch (error) {
        Logger.log('Error in addData: ' + error.message);
        throw new Error(error.message || 'ไม่สามารถเพิ่มข้อมูลได้');
    }
}

function updateData(sheetName, id, data) {
    Logger.log('updateData called: ' + sheetName + ', id: ' + id);
    try {
        var sheet = initializeSheet(sheetName);
        var dataRange = sheet.getDataRange();
        var values = dataRange.getValues();
        if (sheetName === 'teachers' && data.username) {
            if (!/^[a-zA-Z0-9]{3,}$/.test(data.username)) {
                throw new Error('ชื่อผู้ใช้ต้องประกอบด้วยภาษาอังกฤษและตัวเลขอย่างน้อย 3 ตัวอักษร');
            }
            if (data.password && data.password.length < 6) {
                throw new Error('รหัสผ่านต้องมีอย่างน้อย 6 ตัวอักษร');
            }
            for (var r = 1; r < values.length; r++) {
                var rowId = values[r][0];
                var existingUsername = values[r][1] ? values[r][1].toString().trim() : '';
                if (rowId !== id && existingUsername === data.username.trim()) {
                    throw new Error('ชื่อผู้ใช้นี้มีอยู่แล้วในระบบ');
                }
            }

            var pairs = JSON.parse(data.subjectClassPairs || '[]');
            if (pairs.length > 0) {
                var allStudents = getData('students');
                var existingStudentCombos = new Set();
                allStudents.forEach(function(s) {
                    if (s.class && s.classroom) {
                        existingStudentCombos.add(s.class + '/' + s.classroom);
                    }
                });

                for (var i = 0; i < pairs.length; i++) {
                    var pair = pairs[i];
                    var classLevels = pair.classLevels || [];
                    var classrooms = pair.classrooms || [];
                    for (var j = 0; j < classLevels.length; j++) {
                        for (var k = 0; k < classrooms.length; k++) {
                            var comboToCheck = classLevels[j] + '/' + classrooms[k];
                            if (!existingStudentCombos.has(comboToCheck)) {
                                throw new Error('บันทึกล้มเหลว: ไม่สามารถมอบหมายให้ดูแลห้องเรียน "' + classrooms[k] + '" ของระดับชั้น "' + classLevels[j] + '" ได้ เนื่องจากยังไม่มีข้อมูลนักเรียนในห้องดังกล่าว');
                            }
                        }
                    }
                }
            }

        } else if (sheetName === 'students' && data.code) {
            var students = getData('students');
            var codeExists = students.some(s => s.code === data.code && s.id !== id);
            if (codeExists) {
                throw new Error('รหัสนักเรียนนี้มีอยู่ในระบบแล้ว');
            }
        }
        for (var i = 1; i < values.length; i++) {
            if (values[i][0] === id) {
                if (sheetName === 'subjects') {
                    values[i][1] = data.code || '';
                    values[i][2] = data.name || '';
                    values[i][3] = data.credits || '';
                } else if (sheetName === 'teachers') {
                    values[i][1] = data.username || '';
                    if (data.password) values[i][2] = data.password;
                    values[i][3] = data.name || '';
                    values[i][4] = data.subjectIds || '';
                    values[i][5] = data.classLevels || '';
                    values[i][6] = data.resetRequested || false;
                    values[i][7] = data.subjectClassPairs || '[]';
                } else if (sheetName === 'students') {
                    values[i][1] = data.code || '';
                    values[i][2] = data.name || '';
                    if (data.class) values[i][3] = data.class;
                    if (data.classroom) values[i][4] = data.classroom;
                } else if (sheetName === 'attendance') {
                    values[i][1] = data.studentId || '';
                    values[i][2] = data.subjectId || '';
                    values[i][3] = data.date || '';
                    values[i][4] = data.status || '';
                    values[i][5] = data.studentName || '';
                    values[i][6] = data.class || '';
                    values[i][7] = data.classroom || '';
                    values[i][8] = data.teacherName || '';
                    values[i][9] = data.subjectName || '';
                    values[i][10] = data.remark || '';
                }
                dataRange.setValues(values);
                Logger.log('Data updated: ' + sheetName + ', id: ' + id);
                return {
                    success: true
                };
            }
        }
        throw new Error('ไม่พบข้อมูลที่ต้องการอัปเดต');
    } catch (error) {
        Logger.log('Error in updateData: ' + error.message);
        throw new Error(error.message || 'การอัปเดตข้อมูลล้มเหลว');
    }
}

function deleteData(sheetName, id) {
  Logger.log('deleteData called: ' + sheetName + ', id: ' + id);
  try {
    var sheet = initializeSheet(sheetName);
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === id) {
        sheet.deleteRow(i + 1);
        
        if (sheetName === 'subjects') {
          var attendanceSheet = initializeSheet('attendance');
          var attendanceData = attendanceSheet.getDataRange().getValues();
          var rowsToDelete = [];
          
          for (var k = 1; k < attendanceData.length; k++) {
            if (attendanceData[k][2] === id) {
              rowsToDelete.push(k + 1);
            }
          }
          
          rowsToDelete.sort((a, b) => b - a).forEach(rowIndex => {
            attendanceSheet.deleteRow(rowIndex);
            Logger.log('Deleted attendance row: ' + rowIndex + ' due to subject deletion');
          });
        } else if (sheetName === 'teachers') {
          var attendanceSheet = initializeSheet('attendance');
          var attendanceData = attendanceSheet.getDataRange().getValues();
          for (var k = attendanceData.length - 1; k >= 1; k--) {
            if (attendanceData[k][8] === data[i][3]) {
              attendanceSheet.deleteRow(k + 1);
              Logger.log('Deleted attendance row: ' + (k + 1) + ' due to teacher deletion');
            }
          }
        } else if (sheetName === 'students') {
          var attendanceSheet = initializeSheet('attendance');
          var attendanceData = attendanceSheet.getDataRange().getValues();
          for (var k = attendanceData.length - 1; k >= 1; k--) {
            if (attendanceData[k][1] === id) {
              attendanceSheet.deleteRow(k + 1);
              Logger.log('Deleted attendance row: ' + (k + 1) + ' due to student deletion');
            }
          }
        }
        
        Logger.log('Data deleted: ' + sheetName + ', id: ' + id);
        return;
      }
    }
    Logger.log('Delete failed: ID not found');
    throw new Error('ไม่พบข้อมูลที่ต้องการลบ');
  } catch (error) {
    Logger.log('Error in deleteData: ' + error.message);
    throw new Error('การลบข้อมูลล้มเหลว: ' + error.message);
  }
}

function saveAllAttendance(records) {
  Logger.log('saveAllAttendance called: %s records', records.length);
  try {
    var sheet = initializeSheet('attendance');
    var data = sheet.getDataRange().getValues();
    var count = 0;
    var savedRecords = []; //--- ส่วนสำคัญ

    var dataMap = {};
    for (var i = 1; i < data.length; i++) {
      if (data[i][0]) {
        dataMap[data[i][0]] = i + 1;
      }
    }

    var newRows = [];
    
    records.forEach(function(r) {
      var isNew = !r.id || !dataMap[r.id];
      var recordId = isNew ? Utilities.getUuid() : r.id;

      var rowData = [
        recordId, r.studentId || '', r.subjectId || '', r.date || '', r.status || '',
        r.studentName || '', r.class || '', r.classroom || '',
        r.teacherName || '', r.subjectName || '', r.remark || ''
      ];
      
      if (isNew) {
        newRows.push(rowData);
      } else {
        var rowIndex = dataMap[r.id];
        var range = sheet.getRange(rowIndex, 1, 1, 11);
        range.setValues([rowData]);
      }
      
      //--- สร้าง object ที่สมบูรณ์เพื่อส่งกลับไปอัปเดต State ที่ Client
      savedRecords.push({
        id: rowData[0], studentId: rowData[1], subjectId: rowData[2], date: rowData[3], status: rowData[4],
        studentName: rowData[5], class: rowData[6], classroom: rowData[7], teacherName: rowData[8],
        subjectName: rowData[9], remark: rowData[10]
      });
      //--- สิ้นสุดส่วนที่สร้าง object

      count++;
    });

    if (newRows.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, 11).setValues(newRows);
    }

    SpreadsheetApp.flush();
    Logger.log('Saved %s records.', count);
    //--- แก้ไข return เพื่อส่งข้อมูลที่บันทึกแล้วกลับไป
    return { success: true, count: count, savedRecords: savedRecords };
  } catch (e) {
    Logger.log('Error in saveAllAttendance: ' + e.message + " stack: " + e.stack);
    throw new Error('การบันทึกเวลาเรียนล้มเหลว: ' + e.message);
  }
}

function getMonthlyScoreReportData(params) {
    Logger.log('getMonthlyScoreReportData (Fixed) called with: ' + JSON.stringify(params));
    try {
        const { subjectId, classLevel, classroom, teacherName } = params;

        const studentsInClass = getData('students').filter(s => s.class === classLevel && s.classroom === classroom);
        const studentIds = studentsInClass.map(s => s.id);
        const subjectInfo = getData('subjects').find(s => s.id === subjectId);

        const componentsForClass = getData('score_components').filter(c =>
            c.subjectId === subjectId &&
            c.classLevel === classLevel &&
            c.classroom === classroom
        );
        const allScores = getData('scores');
        const scoresForClass = allScores.filter(s =>
            studentIds.includes(s.studentId)
        );

        const reportData = studentsInClass.map(student => {
            let totalScore = 0;
            let totalMaxScore = 0;
            const scoresByComponent = {};
            
            // --- START: ส่วนที่แก้ไข ---
            // กรอง remark โดยใช้ subjectId เพิ่มเติม เพื่อให้ได้ข้อมูลที่ถูกต้องของวิชานั้นๆ
            const remarkRecord = allScores.find(s =>
                s.studentId === student.id &&
                s.componentId === 'remark' &&
                s.subjectId === subjectId
            );
            const studentRemark = remarkRecord ? remarkRecord.remark : '-';
            // --- END: ส่วนที่แก้ไข ---

            componentsForClass.forEach(component => {
                const scoreRecord = scoresForClass.find(s =>
                    s.studentId === student.id && s.componentId === component.id
                );
                const score = scoreRecord ? parseFloat(scoreRecord.score || 0) : 0;
                scoresByComponent[component.id] = score;
                totalScore += score;
                totalMaxScore += parseFloat(component.maxScore);
            });
            
            const finalGrade = calculateGrade(totalScore, totalMaxScore, studentRemark);

            return {
                studentId: student.id,
                studentCode: student.code,
                studentName: student.name,
                scores: scoresByComponent,
                totalScore: totalScore,
                totalMaxScore: totalMaxScore,
                finalGrade: finalGrade
            };
        });

        return {
            success: true,
            reportData: reportData,
            components: componentsForClass,
            reportDetails: {
                subjectInfo: subjectInfo,
                classLevel: classLevel,
                classroom: classroom,
                teacherName: teacherName
            }
        };

    } catch (e) {
        Logger.log('Error in getMonthlyScoreReportData: ' + e.stack);
        throw new Error('ไม่สามารถสร้างข้อมูลรายงานคะแนนได้: ' + e.message);
    }
}


function getMonthlyReportData(params) {
  Logger.log('getMonthlyReportData called with params: ' + JSON.stringify(params));
  try {
    const { year, month, subjectId, classLevel, classroom, teacherName } = params;

    const allAttendance = getData('attendance');
    const allStudents = getData('students').filter(s => s.class === classLevel && s.classroom === classroom);
    const subject = getData('subjects').find(s => s.id === subjectId);

    const startDate = new Date(year, month - 1, 1);
    const endDate = new Date(year, month, 0);
    const daysInMonth = endDate.getDate();

    let workingDays = 0;
    for (let day = 1; day <= daysInMonth; day++) {
        const currentDate = new Date(year, month - 1, day);
        const dayOfWeek = currentDate.getDay();
        if (dayOfWeek !== 0 && dayOfWeek !== 6) {
            workingDays++;
        }
    }

    const reportData = allStudents.map(student => {
      const studentAttendance = {};
      const summary = { present: 0, absent: 0, leave: 0, late: 0, none: 0 };

      for (let day = 1; day <= daysInMonth; day++) {
        const currentDate = new Date(year, month - 1, day);
        const dateStr = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        
        // -----[ START: ส่วนที่แก้ไข ]-----
        // แก้ไขการค้นหา record โดยการลบเงื่อนไข && a.teacherName === teacherName ออก
        // เพื่อให้ดึงข้อมูลการเข้าเรียนทั้งหมดของวิชานี้ โดยไม่จำกัดครูผู้สอน
        const record = allAttendance.find(a => 
          a.studentId === student.id &&
          a.subjectId === subjectId &&
          normalizeDate(a.date) === dateStr
          // ** บรรทัดที่ถูกลบออก: && a.teacherName === teacherName **
        );
        // -----[ END: ส่วนที่แก้ไข ]-----

        const status = record ? record.status : 'none';
        studentAttendance[day] = status;
        if (summary.hasOwnProperty(status)) {
          summary[status]++;
        }
      }

      return {
        studentId: student.id,
        studentCode: student.code,
        studentName: student.name,
        attendance: studentAttendance,
        summary: summary
      };
    });
    
    return {
      success: true,
      reportData: reportData,
      daysInMonth: daysInMonth,
      workingDays: workingDays,
      reportDetails: {
          year: year,
          month: month,
          classLevel: classLevel,
          classroom: classroom,
          teacherName: teacherName,
          subjectInfo: {
            code: subject ? subject.code : 'N/A',
            name: subject ? subject.name : 'ไม่พบข้อมูลวิชา'
          }
      }
    };

  } catch (error) {
    Logger.log('Error in getMonthlyReportData: ' + error.stack);
    throw new Error('ไม่สามารถสร้างข้อมูลรายงานประจำเดือนได้: ' + error.message);
  }
}

function createMonthlyReportPdfFromSheet(params) {
    Logger.log('createMonthlyReportPdfFromSheet called (v13 - Final with comments).');
    
    let newSpreadsheetFile = null;

    try {
        const { year, month, subjectId, classLevel, classroom, teacherName } = params;
        const data = getMonthlyReportData(params); 
        const settings = getSettings();
        
        if (!data.success || data.reportData.length === 0) {
            throw new Error('ไม่พบข้อมูลสำหรับสร้างรายงาน');
        }

        const templateFile = DriveApp.getFileById(MONTHLY_REPORT_SHEET_TEMPLATE_ID);
        const newFileName = `รายงาน PDF ประจำเดือน_${data.reportDetails.subjectInfo.name}_${classLevel}_${classroom}_${month}-${year}`;
        newSpreadsheetFile = templateFile.makeCopy(newFileName);
        
        const newSS = SpreadsheetApp.openById(newSpreadsheetFile.getId());
        const sheet = newSS.getSheetByName(ATTENDANCE_REPORT_SHEET_NAME);
        if (!sheet) {
            throw new Error('ไม่พบชีตเทมเพลตสำหรับรายงานเวลาเรียนที่ชื่อ: ' + ATTENDANCE_REPORT_SHEET_NAME);
        }

        const monthNames = ["มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"];
        const monthYearText = `ประจำเดือน ${monthNames[month-1]} พ.ศ. ${parseInt(year) + 543}`;
        
        const subjectInfo = `รายวิชา: ${data.reportDetails.subjectInfo.code} - ${data.reportDetails.subjectInfo.name}`;
        const classInfo = `ระดับชั้น: ${classLevel} ห้อง: ${classroom}`;
        const teacherInfo = `ครูผู้สอน: ${teacherName}`;
        const detailsText = `${subjectInfo} | ${classInfo} | ${teacherInfo}`;
        const workingDaysText = `จำนวนวันทำการทั้งหมด: ${data.workingDays} วัน`;

        const numCols = data.daysInMonth + 7;

        // เคลียร์และรีเซ็ตความสูงแถว 1-5
        sheet.getRange('1:5').clearContent();
        for (var i = 1; i <= 5; i++) {
          sheet.setRowHeight(i, 21); // รีเซ็ตความสูงกลับเป็นค่าเริ่มต้น
        }
        
        // --- ส่วนของการตั้งค่าหัวกระดาษ ---
        sheet.getRange(1, 1, 1, numCols).merge().setValue(settings.header_text || 'รายงานการเข้าเรียน').setHorizontalAlignment('center').setFontWeight('bold').setFontSize(20);
        sheet.getRange(2, 1, 1, numCols).merge().setValue(monthYearText).setHorizontalAlignment('center').setFontWeight('bold').setFontSize(20);
        sheet.getRange(3, 1, 1, numCols).merge().setValue(detailsText).setHorizontalAlignment('center').setFontSize(20);
        sheet.getRange(4, 1, 1, numCols).merge().setValue(workingDaysText).setHorizontalAlignment('center').setFontSize(16).setFontWeight('bold');
        
        sheet.setRowHeight(1, 1);
        sheet.setRowHeight(2, 1);
        sheet.setRowHeight(3, 1);
        sheet.setRowHeight(4, 1);
        
        // --- ส่วนของการสร้างตาราง ---
        const students = data.reportData;
        const daysInMonth = data.daysInMonth;
        const workingDays = data.workingDays;
        const tableHeader = ['ลำดับ', 'รหัสนักเรียน', 'ชื่อ-นามสกุล'];
        for (let i = 1; i <= daysInMonth; i++) { tableHeader.push(String(i)); }
        tableHeader.push('มา', 'ขาด', 'ลา', 'สาย');

        let totalSummary = { present: 0, absent: 0, leave: 0, late: 0 };
        const tableData = students.map((student, index) => {
            const row = [index + 1, student.studentCode, student.studentName];
            for (let day = 1; day <= daysInMonth; day++) {
                const status = student.attendance[day];
                row.push({ present: 'ม', absent: 'ข', leave: 'ล', late: 'ส', none: '-' }[status] || '-');
            }
            row.push(student.summary.present, student.summary.absent, student.summary.leave, student.summary.late);
            totalSummary.present += student.summary.present;
            totalSummary.absent += student.summary.absent;
            totalSummary.leave += student.summary.leave;
            totalSummary.late += student.summary.late;
            return row;
        });
        
        const totalRow = ['', '', 'รวมทั้งหมด'];
        for(let i=0; i < daysInMonth; i++) { totalRow.push(''); }
        totalRow.push(totalSummary.present, totalSummary.absent, totalSummary.leave, totalSummary.late);

        const totalStudents = students.length;
        const totalPossibleDays = totalStudents * workingDays;
        const percentPresent = totalPossibleDays > 0 ? ((totalSummary.present / totalPossibleDays) * 100).toFixed(2) + '%' : "0.00%";
        const percentAbsent = totalPossibleDays > 0 ? ((totalSummary.absent / totalPossibleDays) * 100).toFixed(2) + '%' : "0.00%";
        const percentLeave = totalPossibleDays > 0 ? ((totalSummary.leave / totalPossibleDays) * 100).toFixed(2) + '%' : "0.00%";
        const percentLate = totalPossibleDays > 0 ? ((totalSummary.late / totalPossibleDays) * 100).toFixed(2) + '%' : "0.00%";
        
        const percentRow = ['', '', 'ร้อยละ (%)'];
        for(let i=0; i < daysInMonth; i++) { percentRow.push(''); }
        percentRow.push(percentPresent, percentAbsent, percentLeave, percentLate);

        const fullTable = [tableHeader, ...tableData, totalRow, percentRow];

        const startRow = 6;
        const tableRange = sheet.getRange(startRow, 1, fullTable.length, numCols);

        tableRange.setValues(fullTable);
        
        tableRange.setFontFamily("Sarabun").setFontSize(16).setVerticalAlignment("middle").setHorizontalAlignment("center");
        sheet.getRange(startRow + 1, 3, tableData.length, 1).setHorizontalAlignment("left"); 
        
        const headerRange = sheet.getRange(startRow, 1, 1, numCols);
        headerRange.setBackground('#F3F4F6').setFontWeight('bold');

        const summaryColStart = 4 + daysInMonth;
        sheet.getRange(startRow, summaryColStart, fullTable.length, 1).setBackground('#BBF7D0'); 
        sheet.getRange(startRow, summaryColStart + 1, fullTable.length, 1).setBackground('#FECACA'); 
        sheet.getRange(startRow, summaryColStart + 2, fullTable.length, 1).setBackground('#FEF3C7'); 
        sheet.getRange(startRow, summaryColStart + 3, fullTable.length, 1).setBackground('#FBCFE8'); 
        sheet.getRange(startRow, summaryColStart, 1, 4).setFontWeight('bold');

        tableRange.setBorder(true, true, true, true, true, true);
        
        sheet.setColumnWidth(1, 40); 
        sheet.autoResizeColumn(2);
        sheet.setColumnWidth(3, 180);
        for(let i=1; i<=daysInMonth; i++){ sheet.setColumnWidth(3 + i, 25); }
        for(let i=1; i<=4; i++){ sheet.setColumnWidth(3 + daysInMonth + i, 50); }
        
        sheet.setFrozenRows(startRow);

        // --- ส่วนของการเพิ่มส่วนลงชื่อท้ายกระดาษ ---
        const signatureStartRow = startRow + fullTable.length + 3; // เว้น 3 แถวจากตาราง
        const signatureLine = 'ลงชื่อ.......................................................';
        const nameLine = '(.......................................................)';
        const signatureFontSize = 18;

        const signatureBlockWidth = 10;
        
        // ครูประจำวิชา (ซ้าย)
        sheet.getRange(signatureStartRow, 2, 1, signatureBlockWidth).merge().setValue(signatureLine).setHorizontalAlignment('center').setFontSize(signatureFontSize);
        sheet.getRange(signatureStartRow + 1, 2, 1, signatureBlockWidth).merge().setValue(nameLine).setHorizontalAlignment('center').setFontSize(signatureFontSize);
        sheet.getRange(signatureStartRow + 2, 2, 1, signatureBlockWidth).merge().setValue('ครูประจำวิชา').setHorizontalAlignment('center').setFontSize(signatureFontSize);

        // หัวหน้าวิชาการ (กลาง)
        const centerCol = Math.floor(numCols / 2) - Math.floor(signatureBlockWidth / 2);
        sheet.getRange(signatureStartRow, centerCol, 1, signatureBlockWidth).merge().setValue(signatureLine).setHorizontalAlignment('center').setFontSize(signatureFontSize);
        sheet.getRange(signatureStartRow + 1, centerCol, 1, signatureBlockWidth).merge().setValue(nameLine).setHorizontalAlignment('center').setFontSize(signatureFontSize);
        sheet.getRange(signatureStartRow + 2, centerCol, 1, signatureBlockWidth).merge().setValue('หัวหน้าวิชาการ').setHorizontalAlignment('center').setFontSize(signatureFontSize);

        // ผู้บริหาร (ขวา)
        const rightCol = numCols - signatureBlockWidth;
        sheet.getRange(signatureStartRow, rightCol, 1, signatureBlockWidth).merge().setValue(signatureLine).setHorizontalAlignment('center').setFontSize(signatureFontSize);
        sheet.getRange(signatureStartRow + 1, rightCol, 1, signatureBlockWidth).merge().setValue(nameLine).setHorizontalAlignment('center').setFontSize(signatureFontSize);
        sheet.getRange(signatureStartRow + 2, rightCol, 1, signatureBlockWidth).merge().setValue('ผู้อำนวยการสถานศึกษา').setHorizontalAlignment('center').setFontSize(signatureFontSize);

        SpreadsheetApp.flush();
        const fileId = newSpreadsheetFile.getId();
        
        const pdfUrl = `https://docs.google.com/spreadsheets/d/${fileId}/export?` +
            'format=pdf' +
            '&size=a4' +
            '&portrait=false' + 
            '&fitw=true' +
            '&sheetnames=false&printtitle=false' +
            '&gridlines=false' + 
            '&top_margin=0.45' +
            '&bottom_margin=0.25' +
            '&left_margin=0.45' +
            '&right_margin=0.45' +
            '&gid=' + sheet.getSheetId(); 

        Logger.log('PDF Report from Sheet created. URL: ' + pdfUrl);
        return { success: true, url: pdfUrl };

    } catch (error) {
        Logger.log('Error in createMonthlyReportPdfFromSheet: ' + error.stack);
        if (newSpreadsheetFile) {
            DriveApp.getFileById(newSpreadsheetFile.getId()).setTrashed(true);
        }
        throw new Error('ไม่สามารถสร้างไฟล์ PDF ได้: ' + error.message);
    }
}

function importCSV(sheetName, csvContent) {
    Logger.log('importCSV called: ' + sheetName);
    try {
        var sheet = initializeSheet(sheetName);
        var csvData = Utilities.parseCsv(csvContent);
        var count = 0;

        if (sheetName === 'subjects') {
            var existingCodes = getData('subjects').map(s => s.code);
            for (var i = 1; i < csvData.length; i++) {
                try {
                    var row = csvData[i];
                    if (row.length >= 2 && row[0] && row[1]) {
                        if (existingCodes.includes(row[0])) {
                            Logger.log('Skipping duplicate subject code: ' + row[0]);
                            continue;
                        }
                        sheet.appendRow([Utilities.getUuid(), row[0], row[1], row[2] || '']);
                        existingCodes.push(row[0]);
                        count++;
                    }
                } catch (e) {
                    Logger.log('Error importing row ' + i + ': ' + e.message);
                }
            }
        } else if (sheetName === 'students') {
            var classSettings = getClassLevelSettings();
            var enabledLevels = classSettings.enabledLevels;
            var existingCodes = getData('students').map(s => s.code);

            for (var i = 1; i < csvData.length; i++) {
                var row = csvData[i];
                if (row.length < 4 || !row[0] || !row[1] || !row[2] || !row[3]) {
                    continue; 
                }

                var studentCodeFromCsv = row[0].trim();
                var classLevelFromCsv = row[2].trim();

                if (existingCodes.includes(studentCodeFromCsv)) {
                    Logger.log('Import failed: Duplicate student code found - ' + studentCodeFromCsv);
                    throw new Error('นำเข้าล้มเหลว: รหัสนักเรียน "' + studentCodeFromCsv + '" มีอยู่แล้วในระบบ หรือมีรหัสซ้ำกันในไฟล์นำเข้า (แถวที่ ' + (i + 1) + ')');
                }

                if (!enabledLevels.includes(classLevelFromCsv)) {
                    Logger.log('Import failed: Class level not found - ' + classLevelFromCsv);
                    throw new Error('นำเข้าล้มเหลว: ไม่พบระดับชั้น "' + classLevelFromCsv + '" ในระบบที่ตั้งค่าไว้ (แถวที่ ' + (i + 1) + ')');
                }

                try {
                    sheet.appendRow([
                        Utilities.getUuid(),
                        studentCodeFromCsv, 
                        row[1].trim(),   
                        classLevelFromCsv, 
                        row[3].trim()      
                    ]);
                    existingCodes.push(studentCodeFromCsv);
                    count++;
                } catch (e) {
                    Logger.log('Error importing row ' + (i + 1) + ': ' + e.message);
                }
            }
        }

        Logger.log('Imported ' + count + ' records to ' + sheetName);
        return {
            count: count
        };
    } catch (error) {
        Logger.log('Error in importCSV: ' + error.message);
        throw new Error('การนำเข้าข้อมูลล้มเหลว: ' + error.message);
    }
}

function checkStudentCode(params) {
  Logger.log('checkStudentCode called: ' + JSON.stringify(params));
  try {
    var sheet = initializeSheet('students');
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][1] === params.code && (!params.studentId || data[i][0] !== params.studentId)) {
        Logger.log('Student code exists: ' + params.code);
        return { result: false };
      }
    }
    Logger.log('Student code is unique: ' + params.code);
    return { result: true };
  } catch (error) {
    Logger.log('Error in checkStudentCode: ' + error.message);
    throw new Error('การตรวจสอบรหัสนักเรียนล้มเหลว: ' + error.message);
  }
}

function getClassroomReport(params) {
    Logger.log('getClassroomReport called: ' + JSON.stringify(params));
    try {
        var startDate = new Date(params.startDate);
        var endDate = new Date(params.endDate);
        var classLevelFilter = params.classLevel || '';
        var classroomFilter = params.classroom || '';
        var subjectIdFilter = params.subjectId || '';
        var teacherId = params.teacherId || null;

        var attendanceData = getData('attendance');
        var subjects = getData('subjects');
        var teachers = getData('teachers');

        if (teacherId) {
            var teacher = teachers.find(function(t) { return t.id === teacherId; });
            if (teacher && teacher.subjectClassPairs) {
                var pairs = [];
                try {
                    pairs = JSON.parse(teacher.subjectClassPairs);
                } catch (e) {
                    Logger.log('Could not parse subjectClassPairs for teacher ' + teacherId + ': ' + e.message);
                }
                var validClasses = pairs.flatMap(p => p.classLevels || []);
                var validClassrooms = pairs.flatMap(p => p.classrooms || []);
                var validSubjectIds = pairs.map(p => p.subjectId);

                attendanceData = attendanceData.filter(function(a) {
                    return validClasses.includes(a.class) &&
                        validClassrooms.includes(a.classroom || '') &&
                        validSubjectIds.includes(String(a.subjectId));
                });
            }
        }

        var reportMap = {};

        attendanceData.forEach(function(a) {
            var attDateStr = normalizeDate(a.date);
            if (!attDateStr) return;

            var dObj = new Date(attDateStr);
            if (dObj < startDate || dObj > endDate) return;

            // *** CORRECTED LOGIC FOR "ALL" OPTION ***
            if (classLevelFilter && classLevelFilter !== 'all' && a.class !== classLevelFilter) return;

            if (classroomFilter && (a.classroom || '').toString().trim() !== classroomFilter) return;
            if (subjectIdFilter && a.subjectId !== subjectIdFilter) return;
            if (['present', 'absent', 'leave', 'late'].indexOf(a.status) === -1) return;

            var key = attDateStr + '_' + (a.class || '') + '_' + (a.classroom || '') + '_' + a.subjectId;
            if (!reportMap[key]) {
                var subjObj = subjects.find(function(s) { return s.id === a.subjectId; }) || {};
                reportMap[key] = {
                    date: attDateStr,
                    class: a.class || '',
                    classroom: a.classroom || '',
                    subjectId: a.subjectId,
                    subjectCode: subjObj.code || '',
                    subjectName: subjObj.name || '',
                    totalStudents: 0,
                    present: 0,
                    absent: 0,
                    leave: 0,
                    late: 0,
                    studentSet: new Set() // Use a Set to count unique students
                };
            }
            var entry = reportMap[key];

            // Only increment total students if the student hasn't been counted for this entry yet
            if (!entry.studentSet.has(a.studentId)) {
                entry.studentSet.add(a.studentId);
                entry.totalStudents++;
            }

            if (a.status === 'present') entry.present++;
            else if (a.status === 'absent') entry.absent++;
            else if (a.status === 'leave') entry.leave++;
            else if (a.status === 'late') entry.late++;
        });

        // Convert map to array and remove the temporary Set
        var result = Object.values(reportMap).map(function(item) {
            delete item.studentSet;
            return item;
        });

        Logger.log('Classroom report generated with ' + result.length + ' rows.');
        return result;
    } catch (error) {
        Logger.log('Error in getClassroomReport: ' + error.message);
        throw new Error('ไม่สามารถสร้างรายงานห้องเรียนได้: ' + error.message);
    }
}

function getSchoolReport(params) {
    Logger.log('getSchoolReport called: ' + JSON.stringify(params));
    try {
        var startDate = new Date(params.startDate);
        var endDate = new Date(params.endDate);
        var classLevelFilter = params.classLevel || '';
        var classroomFilter = params.classroom || '';
        var subjectIdFilter = params.subjectId || '';

        var attendance = getData('attendance');
        var subjects = getData('subjects');

        var filteredAttendance = attendance.filter(a => {
            var attDate = new Date(a.date);
            return attDate >= startDate && attDate <= endDate;
        });
        
        // *** CORRECTED LOGIC FOR "ALL" OPTION ***
        if (classLevelFilter && classLevelFilter !== 'all') {
            filteredAttendance = filteredAttendance.filter(a => a.class === classLevelFilter);
        }
        if (classroomFilter) {
            filteredAttendance = filteredAttendance.filter(a => (a.classroom || '').toString().trim() === classroomFilter);
        }
        if (subjectIdFilter) {
            filteredAttendance = filteredAttendance.filter(a => a.subjectId === subjectIdFilter);
        }

        var reportMap = {};
        
        // Iterate through the already filtered data
        filteredAttendance.forEach(a => {
            if (['present', 'absent', 'leave', 'late'].indexOf(a.status) === -1) return;

            var key = normalizeDate(a.date) + '_' + a.class + '_' + (a.classroom || '') + '_' + a.subjectId;
            if (!reportMap[key]) {
                var subject = subjects.find(s => s.id === a.subjectId) || { code: 'N/A', name: 'N/A' };
                reportMap[key] = {
                    date: normalizeDate(a.date),
                    subjectCode: subject.code,
                    subjectName: subject.name,
                    classLevel: a.class,
                    classroom: a.classroom || '',
                    studentIds: new Set(),
                    present: 0,
                    absent: 0,
                    leave: 0,
                    late: 0,
                };
            }
            reportMap[key].studentIds.add(a.studentId);
            if (reportMap[key].hasOwnProperty(a.status)) {
                reportMap[key][a.status]++;
            }
        });

        var result = Object.values(reportMap).map(item => {
            item.totalStudents = item.studentIds.size;
            delete item.studentIds;
            return item;
        });

        Logger.log('School report generated: ' + result.length);
        return result;
    } catch (error) {
        Logger.log('Error in getSchoolReport: ' + error.message);
        throw new Error('ไม่สามารถสร้างรายงานภาพรวมโรงเรียนได้: ' + error.message);
    }
}

function getSettings() {
  Logger.log('getSettings called');
  try {
    var sheet = initializeSheet('settings');
    var data = getData('settings');
    var settings = {};
    data.forEach(row => {
      settings[row.key] = row.value;
    });
    Logger.log('Settings retrieved: ' + JSON.stringify(settings));
    return settings;
  } catch (error) {
    Logger.log('Error in getSettings: ' + error.message);
    throw new Error('ไม่สามารถดึงการตั้งค่าได้: ' + error.message);
  }
}

function saveSettings(data) {
    Logger.log('saveSettings called with data: ' + JSON.stringify(data));
    try {
        if (!data || !data.admin_password) {
            return { success: false, message: 'ข้อมูลรหัสผ่านผู้ดูแลระบบไม่ครบถ้วน' };
        }
        if (data.admin_password !== 'admin1234') {
            return { success: false, message: 'รหัสผ่านผู้ดูแลระบบไม่ถูกต้อง' };
        }
        if (data.hasOwnProperty('enabled_class_levels')) {
            var currentSettings = getSettings();
            var currentEnabledLevels = [];
            if (currentSettings.enabled_class_levels) {
                try {
                    currentEnabledLevels = JSON.parse(currentSettings.enabled_class_levels);
                } catch(e) {
                    Logger.log('Could not parse enabled_class_levels, defaulting to empty. Error: ' + e.message);
                }
            }
            var newEnabledLevels = JSON.parse(data.enabled_class_levels);
            var levelsToBeDisabled = currentEnabledLevels.filter(level => !newEnabledLevels.includes(level));
            if (levelsToBeDisabled.length > 0) {
                var allStudents = getData('students');
                var conflictingLevels = levelsToBeDisabled.filter(level => 
                    allStudents.some(student => student.class === level)
                );
                if (conflictingLevels.length > 0) {
                    return { success: false, message: 'ไม่สามารถปิดระดับชั้น: ' + conflictingLevels.join(', ') + ' ได้ เพราะยังมีข้อมูลนักเรียนอยู่' };
                }
            }
        }
        var sheet = initializeSheet('settings');
        var dataRange = sheet.getDataRange();
        var values = dataRange.getValues();
        var settingsInSheet = {};
        for (var i = 1; i < values.length; i++) {
            if (values[i][0]) {
                settingsInSheet[values[i][0]] = i + 1;
            }
        }
        for (var key in data) {
            if (data.hasOwnProperty(key) && key !== 'admin_password') {
                var valueToSave = data[key];
                var rowNum = settingsInSheet[key];
                if (rowNum) {
                    sheet.getRange(rowNum, 2).setValue(valueToSave);
                } else {
                    sheet.appendRow([key, valueToSave]);
                }
            }
        }
        SpreadsheetApp.flush(); // บังคับบันทึกข้อมูลลง Sheet
        Logger.log('Settings saved successfully');
        return { success: true };
    } catch (error) {
        Logger.log('Error in saveSettings: ' + error.stack);
        return { success: false, message: error.message || 'การบันทึกการตั้งค่าล้มเหลว' };
    }
}

function loadScoringPage() {
    Logger.log('loadScoringPage called');
    try {
        var settings = getSettings();
        var html = HtmlService.createHtmlOutputFromFile('scoring').getContent();
        var primaryColor = settings.theme_color || '#FF69B4';
        html = html.replace(/{{THEME_COLOR}}/g, primaryColor);
        return html;
    } catch (error) {
        Logger.log('Error in loadScoringPage: ' + error.message);
        throw new Error('ไม่สามารถโหลดหน้าจัดการคะแนนได้');
    }
}

function saveScoreComponent(componentData) {
    Logger.log('saveScoreComponent called with: ' + JSON.stringify(componentData));
    try {
        var sheet = initializeSheet('score_components');
        var id = componentData.id || Utilities.getUuid();
        var dataToSave = [
            id,
            componentData.teacherId,
            componentData.subjectId,
            componentData.classLevel,
            componentData.classroom,
            componentData.componentName,
            componentData.maxScore
        ];

        if (componentData.id) { // Update existing
            var dataRange = sheet.getDataRange();
            var values = dataRange.getValues();
            for (var i = 1; i < values.length; i++) {
                if (values[i][0] === componentData.id) {
                    sheet.getRange(i + 1, 1, 1, dataToSave.length).setValues([dataToSave]);
                    Logger.log('Updated score component id: ' + id);
                    return { success: true, id: id };
                }
            }
            throw new Error("ไม่พบองค์ประกอบคะแนนที่ต้องการอัปเดต");
        } else { // Add new
            sheet.appendRow(dataToSave);
            Logger.log('Added new score component id: ' + id);
            return { success: true, id: id };
        }
    } catch (e) {
        Logger.log('Error in saveScoreComponent: ' + e.stack);
        throw new Error('ไม่สามารถบันทึกองค์ประกอบคะแนนได้: ' + e.message);
    }
}

function deleteScoreComponent(componentId) {
    Logger.log('deleteScoreComponent called for id: ' + componentId);
    try {
        var sheet = initializeSheet('score_components');
        var data = sheet.getDataRange().getValues();
        for (var i = 1; i < data.length; i++) {
            if (data[i][0] === componentId) {
                sheet.deleteRow(i + 1);
                
                // Also delete related scores
                var scoresSheet = initializeSheet('scores');
                var scoresData = scoresSheet.getDataRange().getValues();
                var rowsToDelete = [];
                for (var j = 1; j < scoresData.length; j++) {
                    if (scoresData[j][2] === componentId) {
                        rowsToDelete.push(j + 1);
                    }
                }
                rowsToDelete.sort((a, b) => b - a).forEach(rowIndex => {
                    scoresSheet.deleteRow(rowIndex);
                });
                
                Logger.log('Deleted score component and ' + rowsToDelete.length + ' related scores.');
                return { success: true };
            }
        }
        throw new Error("ไม่พบองค์ประกอบคะแนนที่ต้องการลบ");
    } catch (e) {
        Logger.log('Error in deleteScoreComponent: ' + e.stack);
        throw new Error('การลบองค์ประกอบคะแนนล้มเหลว: ' + e.message);
    }
}

function saveScores(scoresToSave) {
    Logger.log('saveScores (Fixed v5) called with %s records', scoresToSave.length);
    if (!scoresToSave || scoresToSave.length === 0) {
        return { success: true, count: 0, savedRecords: [] };
    }
    try {
        var sheet = initializeSheet('scores');
        var allScores = sheet.getDataRange().getValues();
        var scoreMap = {}; // key: studentId-componentId-subjectId, value: { rowIndex, id }
        
        // --- START: ส่วนที่แก้ไข ---
        // สร้าง key โดยเพิ่ม subjectId เข้าไปด้วย
        for (var i = 1; i < allScores.length; i++) {
            var key = allScores[i][1] + '-' + allScores[i][2] + '-' + (allScores[i][6] || ''); // studentId-componentId-subjectId
            scoreMap[key] = { rowIndex: i + 1, id: allScores[i][0] };
        }
        // --- END: ส่วนที่แก้ไข ---

        var updates = [];
        var newRows = [];
        var savedRecordsForClient = [];
        var count = 0;

        scoresToSave.forEach(function(record) {
            // --- START: ส่วนที่แก้ไข ---
            // ใช้ key ใหม่ในการค้นหา
            var key = record.studentId + '-' + record.componentId + '-' + (record.subjectId || '');
            // --- END: ส่วนที่แก้ไข ---
            
            var existing = scoreMap[key];
            var recordId = existing ? existing.id : Utilities.getUuid();
            
            var rowData = [
                recordId,
                record.studentId,
                record.componentId,
                record.score || '',
                new Date(),
                record.remark || '-',
                record.subjectId || '' // บันทึก subjectId ลงในคอลัมน์ที่ 7
            ];

            if (existing) {
                updates.push({
                    range: sheet.getRange(existing.rowIndex, 1, 1, rowData.length),
                    values: [rowData]
                });
            } else {
                newRows.push(rowData);
            }
            
            // เพิ่ม subjectId ในข้อมูลที่จะส่งกลับไป Client
            savedRecordsForClient.push({
                id: rowData[0],
                studentId: rowData[1],
                componentId: rowData[2],
                score: rowData[3],
                remark: rowData[5],
                subjectId: rowData[6] // ส่ง subjectId กลับไปด้วย
            });
            count++;
        });

        if (updates.length > 0) {
            updates.forEach(update => update.range.setValues(update.values));
        }
        if (newRows.length > 0) {
            sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, newRows[0].length).setValues(newRows);
        }
        
        SpreadsheetApp.flush();
        Logger.log('Saved scores. Processed: ' + count + ' records.');
        return { success: true, count: count, savedRecords: savedRecordsForClient };
    } catch (e) {
        Logger.log('Error in saveScores: ' + e.stack);
        throw new Error('การบันทึกคะแนนล้มเหลว: ' + e.message);
    }
}


function normalizeDate(dateStr) {
  try {
    const date = new Date(dateStr);
    if (isNaN(date.getTime())) {
      return '';
    }
    return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  } catch (e) {
    return '';
  }
}

function createScoreReportPdfFromSheet(params) {
    Logger.log('createScoreReportPdfFromSheet called with: ' + JSON.stringify(params));
    let newSpreadsheetFile = null;
    try {
        const data = getMonthlyScoreReportData(params);
        if (!data.success || data.reportData.length === 0) {
            throw new Error('ไม่พบข้อมูลสำหรับสร้างรายงานสรุปคะแนน');
        }

        const templateFile = DriveApp.getFileById(MONTHLY_REPORT_SHEET_TEMPLATE_ID);
        const newFileName = `รายงานสรุปคะแนน_${data.reportDetails.subjectInfo.name}_${data.reportDetails.classLevel}_${data.reportDetails.classroom}`;
        newSpreadsheetFile = templateFile.makeCopy(newFileName);
        
        const newSS = SpreadsheetApp.openById(newSpreadsheetFile.getId());
        const sheet = newSS.getSheetByName(SCORE_REPORT_SHEET_NAME);
        if (!sheet) {
            throw new Error('ไม่พบชีตเทมเพลตสำหรับรายงานคะแนนที่ชื่อ: ' + SCORE_REPORT_SHEET_NAME);
        }
        sheet.clear();

        const { reportData, components, reportDetails } = data;
        const settings = getSettings();

        // --- คำนวณความกว้างตารางและตัดสินใจเรื่องการจัดหน้า ---
        const A4_PORTRAIT_TARGET_WIDTH_PX = 750;
        let totalTablePixelWidth = 0;
        totalTablePixelWidth += 40;  // ลำดับ
        totalTablePixelWidth += 100; // รหัสนักเรียน
        totalTablePixelWidth += 200; // ชื่อ-นามสกุล
        totalTablePixelWidth += components.length * 100;
        totalTablePixelWidth += 90;  // คะแนนรวม
        totalTablePixelWidth += 80;  // เกรด
        
        let useFitToWidth = false;
        let leftPaddingWidth = 0;

        if (totalTablePixelWidth > A4_PORTRAIT_TARGET_WIDTH_PX) {
            useFitToWidth = true;
        } else {
            leftPaddingWidth = Math.round((A4_PORTRAIT_TARGET_WIDTH_PX - totalTablePixelWidth) / 2);
        }
        const tableStartCol = leftPaddingWidth > 0 ? 2 : 1;

        const tableHeader = ['ลำดับ', 'รหัสนักเรียน', 'ชื่อ-นามสกุล'];
        let totalMaxScore = 0;
        components.forEach(c => {
            tableHeader.push(`${c.componentName}\n(เต็ม ${c.maxScore})`);
            totalMaxScore += parseFloat(c.maxScore);
        });
        tableHeader.push(`คะแนนรวม\n(เต็ม ${totalMaxScore.toFixed(2)})`, 'เกรด');
        const numCols = tableHeader.length;

        // --- 1. ส่วนหัวกระดาษ ---
        const headerCols = tableStartCol === 1 ? numCols : numCols + 1;
        sheet.getRange(1, 1, 1, headerCols).merge().setValue(settings.header_text || 'รายงานสรุปผลการเรียน').setFontWeight('bold').setFontSize(20).setFontFamily("Sarabun").setHorizontalAlignment('center');
        sheet.getRange(2, 1, 1, headerCols).merge().setValue(`รายวิชา: ${reportDetails.subjectInfo.code} - ${reportDetails.subjectInfo.name}`).setFontSize(18).setFontFamily("Sarabun").setHorizontalAlignment('center');
        sheet.getRange(3, 1, 1, headerCols).merge().setValue(`ระดับชั้น: ${reportDetails.classLevel} ห้อง: ${reportDetails.classroom} | ครูผู้สอน: ${reportDetails.teacherName}`).setFontSize(16).setFontFamily("Sarabun").setHorizontalAlignment('center');
        
        // --- 2. ส่วนตารางข้อมูล ---
        const tableData = reportData.map((student, index) => {
            const row = [index + 1, student.studentCode, student.studentName];
            components.forEach(c => {
                row.push(student.scores[c.id] || 0);
            });
            
            // --- START: ส่วนที่แก้ไข ---
            // เปลี่ยนจากการเรียก calculateGrade() ใหม่ เป็นการใช้ student.finalGrade ที่คำนวณไว้แล้ว
            row.push(student.totalScore.toFixed(2), student.finalGrade);
            // --- END: ส่วนที่แก้ไข ---

            return row;
        });

        const footerRowData = ['', '', 'ค่าเฉลี่ยของห้อง'];
        const componentAvgs = components.map(c => {
            const sum = reportData.reduce((acc, student) => acc + (student.scores[c.id] || 0), 0);
            return reportData.length > 0 ? (sum / reportData.length).toFixed(2) : '0.00';
        });
        const totalAvg = componentAvgs.reduce((acc, avg) => acc + parseFloat(avg), 0);
        footerRowData.push(...componentAvgs, totalAvg.toFixed(2), '-');
        
        const startRow = 5;
        const fullTable = [tableHeader, ...tableData];
        
        sheet.getRange(startRow, tableStartCol, fullTable.length, numCols).setValues(fullTable);
        const footerRange = sheet.getRange(startRow + fullTable.length, tableStartCol, 1, numCols);
        footerRange.setValues([footerRowData]);
        
        const tableRange = sheet.getRange(startRow, tableStartCol, fullTable.length + 1, numCols);

        // --- 3. การจัดรูปแบบตาราง ---
        tableRange.setFontFamily("Sarabun").setFontSize(16).setVerticalAlignment("middle").setHorizontalAlignment("center");
        sheet.getRange(startRow + 1, tableStartCol + 2, tableData.length, 1).setHorizontalAlignment("left"); 
        sheet.getRange(startRow, tableStartCol, 1, numCols).setFontWeight('bold').setBackground('#F3F4F6').setWrap(true);
        
        footerRange.setFontWeight('bold').setBackground('#F3F4F6');
        
        sheet.getRange(footerRange.getRow(), tableStartCol, 1, 3).merge().setHorizontalAlignment('center');

        // --- 4. การปรับความกว้างคอลัมน์ ---
        if(leftPaddingWidth > 0){
            sheet.setColumnWidth(1, leftPaddingWidth);
        }
        sheet.setColumnWidth(tableStartCol, 40);
        sheet.setColumnWidth(tableStartCol + 1, 100);
        sheet.setColumnWidth(tableStartCol + 2, 200);
        for(let i = 0; i < components.length; i++) {
            sheet.setColumnWidth(tableStartCol + 3 + i, 100);
        }
        sheet.setColumnWidth(tableStartCol + 3 + components.length, 90);
        sheet.setColumnWidth(tableStartCol + 4 + components.length, 80);

        tableRange.setBorder(true, true, true, true, true, true);
        sheet.setFrozenRows(startRow);

        // --- ส่วนท้ายกระดาษ (ลายเซ็น) ---
        const signatureStartRow = startRow + tableRange.getNumRows() + 2; 
        
        const signatureBlockNumCols = Math.max(3, Math.floor(numCols / 2) - 2); 

        // --- บล็อกซ้าย: ครูประจำวิชา (จัดชิดขอบซ้ายของตาราง) ---
        const leftSigStartCol = tableStartCol;
        sheet.getRange(signatureStartRow, leftSigStartCol, 1, signatureBlockNumCols)
            .merge()
            .setValue('ลงชื่อ ..................................................\n(.......................................................)\nครูประจำวิชา')
            .setWrap(true)
            .setHorizontalAlignment('center')
            .setVerticalAlignment('bottom')
            .setFontFamily("Sarabun").setFontSize(16);

        // --- บล็อกขวา: หัวหน้ากลุ่มสาระ (จัดชิดขอบขวาของตาราง) ---
        const rightSigStartCol = (tableStartCol + numCols) - signatureBlockNumCols;
        sheet.getRange(signatureStartRow, rightSigStartCol, 1, signatureBlockNumCols)
            .merge()
            .setValue('ลงชื่อ ..................................................\n(.......................................................)\nหัวหน้าวิชาการ')
            .setWrap(true)
            .setHorizontalAlignment('center')
            .setVerticalAlignment('bottom')
            .setFontFamily("Sarabun").setFontSize(16);

        sheet.setRowHeight(signatureStartRow, 75);
        
        SpreadsheetApp.flush();
        const fileId = newSpreadsheetFile.getId();
        
        // --- 6. สร้าง URL สำหรับ Export PDF ---
        const pdfUrl = `https://docs.google.com/spreadsheets/d/${fileId}/export?` +
            'format=pdf' +
            '&size=a4' +
            '&portrait=true' +
            '&fitw=' + useFitToWidth + 
            '&sheetnames=false&printtitle=false' +
            '&gridlines=false' + 
            '&top_margin=0.25' +      
            '&bottom_margin=0.25' +   
            '&left_margin=0.7' +   
            '&right_margin=0.35' +    
            '&gid=' + sheet.getSheetId();

        Logger.log('Score Report PDF created (Smart Centering). URL: ' + pdfUrl);
        return { success: true, url: pdfUrl };

    } catch (error) {
        Logger.log('Error in createScoreReportPdfFromSheet: ' + error.stack);
        if (newSpreadsheetFile) {
            DriveApp.getFileById(newSpreadsheetFile.getId()).setTrashed(true);
        }
        return { success: false, message: 'ไม่สามารถสร้างไฟล์ PDF ได้: ' + error.message };
    }
}

function calculateGrade(score, maxScore, remark) {
    // --- START: ส่วนที่แก้ไข ---
    // เพิ่มการตรวจสอบ remark ก่อน
    if (remark && remark !== '-') {
        return remark;
    }
    // --- END: ส่วนที่แก้ไข ---

    if (typeof score !== 'number' && isNaN(parseFloat(score))) {
        return score;
    }

    if (maxScore === 0) return '-';
    const percentage = (score / maxScore) * 100;
    if (percentage >= 80) return '4';
    if (percentage >= 75) return '3.5';
    if (percentage >= 70) return '3';
    if (percentage >= 65) return '2.5';
    if (percentage >= 60) return '2';
    if (percentage >= 55) return '1.5';
    if (percentage >= 50) return '1';
    return '0';
}

/**
 * โหลดเนื้อหา HTML สำหรับหน้าจัดการคุณลักษณะฯ
 */
function loadCharacteristicsPage() {
  Logger.log('loadCharacteristicsPage called');
  try {
    return HtmlService.createHtmlOutputFromFile('characteristics').getContent();
  } catch (error) {
    Logger.log('Error in loadCharacteristicsPage: ' + error.message);
    throw new Error('ไม่สามารถโหลดหน้าจัดการคุณลักษณะฯ ได้');
  }
}

/**
 * บันทึกคะแนนคุณลักษณะฯ ลงชีต
 */
function saveCharacteristicsScores(records) {
    Logger.log('saveCharacteristicsScores called with %s records', records.length);
    if (!records || records.length === 0) {
        return { success: true, count: 0, savedRecords: [] };
    }
    try {
        var sheet = initializeSheet('characteristics_scores');
        var allScores = sheet.getDataRange().getValues();
        var scoreMap = {}; 
        // สร้าง map เพื่อการค้นหาที่รวดเร็ว: key = studentId-itemId, value = { rowIndex, id }
        for (var i = 1; i < allScores.length; i++) {
            var key = allScores[i][1] + '-' + allScores[i][2]; // studentId-itemId
            scoreMap[key] = { rowIndex: i + 1, id: allScores[i][0] };
        }

        var newRows = [];
        var updatedCount = 0;
        var savedRecordsForClient = [];

        records.forEach(function(record) {
            var key = record.studentId + '-' + record.itemId;
            var existing = scoreMap[key];
            var recordId = existing ? existing.id : Utilities.getUuid();
            var rowData = [recordId, record.studentId, record.itemId, record.score, record.teacherId, new Date()];

            if (existing) {
                // อัปเดตแถวที่มีอยู่แล้ว
                sheet.getRange(existing.rowIndex, 1, 1, rowData.length).setValues([rowData]);
            } else {
                // เพิ่มแถวใหม่
                newRows.push(rowData);
            }
            updatedCount++;
            savedRecordsForClient.push({id: rowData[0], studentId: rowData[1], itemId: rowData[2], score: rowData[3]});
        });

        if (newRows.length > 0) {
            sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, newRows[0].length).setValues(newRows);
        }

        SpreadsheetApp.flush();
        Logger.log('Saved characteristics scores. Total records processed: ' + updatedCount);
        return { success: true, count: updatedCount, savedRecords: savedRecordsForClient };
    } catch (e) {
        Logger.log('Error in saveCharacteristicsScores: ' + e.stack);
        throw new Error('การบันทึกคะแนนคุณลักษณะฯ ล้มเหลว: ' + e.message);
    }
}

function createCharacteristicsReportPdf(params) {
    Logger.log('createCharacteristicsReportPdf final attempt called with: ' + JSON.stringify(params));
    let newSpreadsheetFile = null; 
    try {
        const reportResult = getCharacteristicsReportData(params);
        if (!reportResult.success) {
            throw new Error(reportResult.message || 'ไม่สามารถดึงข้อมูลรายงานได้');
        }
        const { reportData, items, reportDetails } = reportResult;
        const { classLevel, classroom, teacherName } = reportDetails;

        if (reportData.length === 0) {
            throw new Error('ไม่พบข้อมูลสำหรับสร้างรายงาน');
        }
        
        // ใช้วิธีคัดลอกไฟล์เทมเพลตหลักเหมือนเดิม ไฟล์นี้จะถูกเก็บไว้ใน Drive ของคุณ
        const templateFile = DriveApp.getFileById(MONTHLY_REPORT_SHEET_TEMPLATE_ID);
        const newFileName = `รายงานคุณลักษณะฯ_${classLevel}_${classroom}_${new Date().getTime()}`;
        newSpreadsheetFile = templateFile.makeCopy(newFileName);
        
        const newSS = SpreadsheetApp.openById(newSpreadsheetFile.getId());
        const sheet = newSS.getSheetByName(CHARACTERISTICS_REPORT_SHEET_NAME);
        if (!sheet) {
            throw new Error('ไม่พบชีตเทมเพลตสำหรับรายงานคุณลักษณะฯ ที่ชื่อว่า "' + CHARACTERISTICS_REPORT_SHEET_NAME + '" กรุณาสร้างชีตเปล่าในไฟล์เทมเพลตหลักและตั้งชื่อให้ตรงกัน');
        }
        sheet.clear(); 
        
        const characteristicItems = items.filter(i => i.itemGroup === 'คุณลักษณะอันพึงประสงค์');
        const readingItems = items.filter(i => i.itemGroup === 'การอ่าน คิดวิเคราะห์ และเขียน');
        const totalColumns = 3 + items.length + 2; 

        // 1. ส่วนหัวกระดาษ (เหมือนเดิม)
        sheet.getRange(1, 1, 1, totalColumns).merge().setValue('แบบประเมินคุณลักษณะอันพึงประสงค์ และการอ่าน คิดวิเคราะห์ และเขียน').setFontWeight('bold').setFontSize(18).setFontFamily("Sarabun").setHorizontalAlignment('center');
        sheet.getRange(2, 1, 1, totalColumns).merge().setValue(`ระดับชั้น ${classLevel} ห้อง ${classroom}`).setFontSize(18).setFontFamily("Sarabun").setHorizontalAlignment('center');
        sheet.getRange(3, 1, 1, totalColumns).merge().setValue(`ครูผู้ประเมิน: ${teacherName}`).setFontSize(16).setFontFamily("Sarabun").setHorizontalAlignment('center');

        // 2. ส่วนตารางข้อมูล (เหมือนเดิม)
        const startRow = 5;
        // ... (ส่วนของการสร้าง header และข้อมูลนักเรียนทั้งหมดเหมือนเดิม ไม่มีการเปลี่ยนแปลง) ...
        const headerRow1Data = ['', '', ''];
        if (characteristicItems.length > 0) {
            headerRow1Data.push('คุณลักษณะอันพึงประสงค์');
            for (let i = 1; i < characteristicItems.length; i++) headerRow1Data.push('');
        }
        if (readingItems.length > 0) {
            headerRow1Data.push('การอ่าน คิดวิเคราะห์ และเขียน');
            for (let i = 1; i < readingItems.length; i++) headerRow1Data.push('');
        }
        headerRow1Data.push('', '');
        sheet.getRange(startRow, 1, 1, headerRow1Data.length).setValues([headerRow1Data]);
        const headerRow2Data = ['ลำดับ', 'รหัสนักเรียน', 'ชื่อ-นามสกุล'];
        items.forEach(item => headerRow2Data.push(item.itemName.replace(/(\d+\.\s)/, '')));
        headerRow2Data.push('สรุป\n(คุณลักษณะฯ)', 'สรุป\n(การอ่านฯ)');
        sheet.getRange(startRow + 1, 1, 1, headerRow2Data.length).setValues([headerRow2Data]);
        sheet.getRange(startRow, 1, 2, 1).merge();
        sheet.getRange(startRow, 2, 2, 1).merge();
        sheet.getRange(startRow, 3, 2, 1).merge();
        let currentCol = 4;
        if (characteristicItems.length > 0) {
            sheet.getRange(startRow, currentCol, 1, characteristicItems.length).merge();
            currentCol += characteristicItems.length;
        }
        if (readingItems.length > 0) {
            sheet.getRange(startRow, currentCol, 1, readingItems.length).merge();
        }
        sheet.getRange(startRow, totalColumns - 1, 2, 1).merge();
        sheet.getRange(startRow, totalColumns, 2, 1).merge();
        const tableStartRow = startRow + 2;
        const tableData = reportData.map((student, index) => {
            const row = [index + 1, student.studentCode, student.studentName];
            items.forEach(item => row.push(student.scores[item.id] !== undefined ? student.scores[item.id] : ''));
            row.push(student.characteristicResult, student.readingResult);
            return row;
        });
        sheet.getRange(tableStartRow, 1, tableData.length, totalColumns).setValues(tableData);
        
        // 3. จัดรูปแบบตาราง (เหมือนเดิม)
        const fullTableRange = sheet.getRange(startRow, 1, tableData.length + 2, totalColumns);
        fullTableRange.setFontFamily("Sarabun").setFontSize(16).setVerticalAlignment("middle").setBorder(true, true, true, true, true, true);
        const fullHeaderRange = sheet.getRange(startRow, 1, 2, totalColumns);
        fullHeaderRange.setFontWeight('bold').setBackground('#F3F4F6').setHorizontalAlignment('center');
        if (items.length > 0) {
            sheet.getRange(startRow + 1, 4, 1, items.length).setTextRotation(90);
        }
        sheet.setRowHeight(startRow + 1, 150);
        sheet.getRange(tableStartRow, 1, tableData.length, 1).setHorizontalAlignment("center");
        sheet.getRange(tableStartRow, 2, tableData.length, 1).setHorizontalAlignment("center");
        sheet.getRange(tableStartRow, 3, tableData.length, 1).setHorizontalAlignment("left");
        sheet.getRange(tableStartRow, 4, tableData.length, items.length).setHorizontalAlignment("center");
        for(let i = 0; i < tableData.length; i++) {
          sheet.setRowHeight(tableStartRow + i, 30);
        }
        const summaryCols = sheet.getRange(startRow, totalColumns - 1, tableData.length + 2, 2);
        summaryCols.setBackground('#E0E7FF').setFontWeight('bold').setHorizontalAlignment("center");

        // 4. ปรับขนาดคอลัมน์ (เหมือนเดิม)
        sheet.setColumnWidth(1, 50);
        sheet.setColumnWidth(2, 100);
        sheet.setColumnWidth(3, 220);
        if (items.length > 0) {
           for (let i = 0; i < items.length; i++) {
               sheet.setColumnWidth(4 + i, 60); 
           }
        }
        sheet.setColumnWidth(4 + items.length, 90);
        sheet.setColumnWidth(5 + items.length, 90);

        // 5. ส่วนลงชื่อท้ายกระดาษ (เหมือนเดิม)
        const signatureStartRow = tableStartRow + tableData.length + 2;
        const signatureBlockNumCols = 6; 
        const leftSigStartCol = 2;
        sheet.getRange(signatureStartRow, leftSigStartCol, 1, signatureBlockNumCols).merge().setValue('ลงชื่อ ..................................................\n(.......................................................)\nครูประจำวิชา/ครูที่ปรึกษา').setWrap(true).setHorizontalAlignment('center').setVerticalAlignment('bottom').setFontFamily("Sarabun").setFontSize(16);
        sheet.setRowHeight(signatureStartRow, 85);

        const rightSigStartCol = totalColumns - signatureBlockNumCols;
        sheet.getRange(signatureStartRow, rightSigStartCol, 1, signatureBlockNumCols).merge().setValue('ลงชื่อ ..................................................\n(.......................................................)\nหัวหน้าฝ่ายวิชาการ').setWrap(true).setHorizontalAlignment('center').setVerticalAlignment('bottom').setFontFamily("Sarabun").setFontSize(16);
        
        // การตรึงแถวจะไม่มีผลกับการพิมพ์ PDF เมื่อ printtitle=false แต่ยังคงมีประโยชน์เมื่อเปิดดูไฟล์ Sheet โดยตรง
        sheet.setFrozenRows(startRow + 1);

        SpreadsheetApp.flush();
        const fileId = newSS.getId();

        // --- START: แก้ไข URL ของ PDF ---
        const pdfUrl = `https://docs.google.com/spreadsheets/d/${fileId}/export?` +
                       'format=pdf' +
                       '&size=a4' +
                       '&portrait=true' +
                       '&fitw=true' +
                       // ใช้ printtitle=false เพื่อเอาชื่อไฟล์ออก (แต่หัวตารางจะไม่ซ้ำ)
                       '&sheetnames=false&printtitle=false' + 
                       '&gridlines=false' + 
                       '&gid=' + sheet.getSheetId();
        // --- END: แก้ไข URL ของ PDF ---
        
        newSpreadsheetFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        
        return { success: true, url: pdfUrl };
    } catch (error) {
        Logger.log('Error in createCharacteristicsReportPdf: ' + error.stack);
        if (newSpreadsheetFile) {
           // DriveApp.getFileById(newSpreadsheetFile.getId()).setTrashed(true);
        }
        return { success: false, message: 'ไม่สามารถสร้างไฟล์ PDF ได้: ' + error.message };
    }
}

function getCharacteristicsReportData(params) {
    try {
        const { classLevel, classroom, teacherName } = params; 
        if (!classLevel || !classroom) {
            throw new Error('กรุณาระบุระดับชั้นและห้องเรียน');
        }

        const studentsInClass = getData('students')
            .filter(s => s.class === classLevel && s.classroom === classroom)
            .sort((a,b) => (a.code || '').localeCompare(b.code));

        if (studentsInClass.length === 0) {
            return { success: true, reportData: [], items: [], reportDetails: params }; 
        }

        const studentIds = studentsInClass.map(s => s.id);
        const allScores = getData('characteristics_scores').filter(s => studentIds.includes(s.studentId));
        const allItems = getData('characteristics_items').sort((a, b) => a.order - b.order);

        const reportData = studentsInClass.map(student => {
            let scores = {};
            let characteristicSum = 0, characteristicCount = 0;
            let readingSum = 0, readingCount = 0;

            allItems.forEach(item => {
                const scoreRecord = allScores.find(s => s.studentId === student.id && s.itemId === item.id);
                if (scoreRecord && scoreRecord.score !== '') {
                    const score = parseInt(scoreRecord.score);
                    scores[item.id] = score;

                    if (item.itemGroup === 'คุณลักษณะอันพึงประสงค์') {
                        characteristicSum += score;
                        characteristicCount++;
                    } else if (item.itemGroup === 'การอ่าน คิดวิเคราะห์ และเขียน') {
                        readingSum += score;
                        readingCount++;
                    }
                }
            });
            
            // --- START: ส่วนที่แก้ไข ---
            let characteristicResult = '-';
            if (characteristicCount > 0) {
                const percentage = (characteristicSum / (characteristicCount * 3)) * 100;
                characteristicResult = calculateCharacteristicGrade(percentage);
            }

            let readingResult = '-';
            if (readingCount > 0) {
                const percentage = (readingSum / (readingCount * 3)) * 100;
                readingResult = calculateCharacteristicGrade(percentage);
            }
            // --- END: ส่วนที่แก้ไข ---

            return {
                studentId: student.id,
                studentCode: student.code,
                studentName: student.name,
                scores: scores,
                characteristicResult: characteristicResult,
                readingResult: readingResult,
            };
        });

        return { success: true, reportData: reportData, items: allItems, reportDetails: params };

    } catch (e) {
        Logger.log('Error in getCharacteristicsReportData: ' + e.stack);
        return { success: false, message: e.message };
    }
}

function loadActivitiesPage() {
  Logger.log('loadActivitiesPage called');
  try {
    // เนื่องจาก UI เหมือนกับหน้า Scoring เราสามารถใช้ไฟล์เดียวกันและปรับแก้เล็กน้อยที่ Client-side ได้
    // แต่เพื่อความชัดเจน เราจะสร้างไฟล์ใหม่ชื่อ activities.html
    return HtmlService.createHtmlOutputFromFile('activities').getContent();
  } catch (error) {
    Logger.log('Error in loadActivitiesPage: ' + error.message);
    throw new Error('ไม่สามารถโหลดหน้าจัดการกิจกรรมฯ ได้');
  }
}

/**
 * บันทึก (เพิ่ม/แก้ไข) องค์ประกอบกิจกรรม
 */
function saveActivityComponent(componentData) {
  Logger.log('saveActivityComponent called with: ' + JSON.stringify(componentData));
  try {
    var sheet = initializeSheet('activity_components');
    var id = componentData.id || Utilities.getUuid();
    var dataToSave = [
      id,
      componentData.teacherId,
      componentData.classLevel,
      componentData.classroom,
      componentData.componentName
    ];

    if (componentData.id) { // Update
      var dataRange = sheet.getDataRange();
      var values = dataRange.getValues();
      for (var i = 1; i < values.length; i++) {
        if (values[i][0] === componentData.id) {
          sheet.getRange(i + 1, 1, 1, dataToSave.length).setValues([dataToSave]);
          Logger.log('Updated activity component id: ' + id);
          return { success: true, id: id };
        }
      }
      throw new Error("ไม่พบองค์ประกอบกิจกรรมที่ต้องการอัปเดต");
    } else { // Add new
      sheet.appendRow(dataToSave);
      Logger.log('Added new activity component id: ' + id);
      return { success: true, id: id };
    }
  } catch (e) {
    Logger.log('Error in saveActivityComponent: ' + e.stack);
    throw new Error('ไม่สามารถบันทึกองค์ประกอบกิจกรรมได้: ' + e.message);
  }
}

/**
 * ลบองค์ประกอบกิจกรรมและผลการประเมินที่เกี่ยวข้อง
 */
function deleteActivityComponent(componentId) {
  Logger.log('deleteActivityComponent called for id: ' + componentId);
  try {
    var sheet = initializeSheet('activity_components');
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === componentId) {
        sheet.deleteRow(i + 1);
        
        var scoresSheet = initializeSheet('activity_scores');
        var scoresData = scoresSheet.getDataRange().getValues();
        var rowsToDelete = [];
        for (var j = 1; j < scoresData.length; j++) {
          if (scoresData[j][2] === componentId) {
            rowsToDelete.push(j + 1);
          }
        }
        rowsToDelete.sort((a, b) => b - a).forEach(rowIndex => {
          scoresSheet.deleteRow(rowIndex);
        });
        
        Logger.log('Deleted activity component and ' + rowsToDelete.length + ' related results.');
        return { success: true };
      }
    }
    throw new Error("ไม่พบองค์ประกอบกิจกรรมที่ต้องการลบ");
  } catch (e) {
    Logger.log('Error in deleteActivityComponent: ' + e.stack);
    throw new Error('การลบองค์ประกอบกิจกรรมล้มเหลว: ' + e.message);
  }
}

/**
 * บันทึกผลการประเมินกิจกรรม (ผ่าน/ไม่ผ่าน)
 */
function saveActivityScores(records) {
  Logger.log('saveActivityScores called with %s records', records.length);
  if (!records || records.length === 0) {
    return { success: true, count: 0, savedRecords: [] };
  }
  try {
    var sheet = initializeSheet('activity_scores');
    var allScores = sheet.getDataRange().getValues();
    var scoreMap = {};
    for (var i = 1; i < allScores.length; i++) {
      var key = allScores[i][1] + '-' + allScores[i][2]; // studentId-componentId
      scoreMap[key] = { rowIndex: i + 1, id: allScores[i][0] };
    }

    var newRows = [];
    var count = 0;
    var savedRecordsForClient = [];

    records.forEach(function(record) {
      var key = record.studentId + '-' + record.componentId;
      var existing = scoreMap[key];
      var recordId = existing ? existing.id : Utilities.getUuid();
      var rowData = [recordId, record.studentId, record.componentId, record.result, record.teacherId, new Date()];
      
      if (existing) {
        sheet.getRange(existing.rowIndex, 1, 1, rowData.length).setValues([rowData]);
      } else {
        newRows.push(rowData);
      }
      count++;
      savedRecordsForClient.push({id: rowData[0], studentId: rowData[1], componentId: rowData[2], result: rowData[3]});
    });

    if (newRows.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, newRows[0].length).setValues(newRows);
    }

    SpreadsheetApp.flush();
    Logger.log('Saved activity results. Total records processed: ' + count);
    return { success: true, count: count, savedRecords: savedRecordsForClient };
  } catch (e) {
    Logger.log('Error in saveActivityScores: ' + e.stack);
    throw new Error('การบันทึกผลกิจกรรมล้มเหลว: ' + e.message);
  }
}

/**
 * ดึงข้อมูลสำหรับสร้างรายงานสรุปผลกิจกรรม
 */
function getActivitiesReportData(params) {
  Logger.log('getActivitiesReportData called with params: ' + JSON.stringify(params));
  try {
    const { classLevel, classroom, teacherName } = params;
    if (!classLevel || !classroom) {
      throw new Error('กรุณาระบุระดับชั้นและห้องเรียน');
    }

    const studentsInClass = getData('students')
      .filter(s => s.class === classLevel && s.classroom === classroom)
      .sort((a,b) => (a.code || '').localeCompare(b.code));

    if (studentsInClass.length === 0) {
      return { success: true, reportData: [], components: [], reportDetails: params };
    }

    const studentIds = studentsInClass.map(s => s.id);
    const allScores = getData('activity_scores').filter(s => studentIds.includes(s.studentId));
    const allComponents = getData('activity_components')
      .filter(c => c.classLevel === classLevel && c.classroom === classroom)
      .sort((a,b) => (a.componentName || '').localeCompare(b.componentName));
    
    const reportData = studentsInClass.map(student => {
      let results = {};
      let passCount = 0;
      let failCount = 0;

      allComponents.forEach(component => {
        const scoreRecord = allScores.find(s => s.studentId === student.id && s.componentId === component.id);
        if (scoreRecord) {
          results[component.id] = scoreRecord.result;
          if (scoreRecord.result === 'ผ่าน') {
            passCount++;
          } else if (scoreRecord.result === 'ไม่ผ่าน') {
            failCount++;
          }
        }
      });
      
      let finalResult = 'ไม่ระบุ';
      if (passCount + failCount > 0) {
          finalResult = failCount > 0 ? 'ไม่ผ่าน' : 'ผ่าน';
      }

      return {
        studentId: student.id,
        studentCode: student.code,
        studentName: student.name,
        results: results,
        finalResult: finalResult,
      };
    });

    return { success: true, reportData: reportData, components: allComponents, reportDetails: params };
  } catch (e) {
    Logger.log('Error in getActivitiesReportData: ' + e.stack);
    return { success: false, message: e.message };
  }
}


/**
 * สร้างไฟล์ PDF รายงานกิจกรรมพัฒนาผู้เรียน
 */
/**
 * สร้างไฟล์ PDF รายงานกิจกรรมพัฒนาผู้เรียน
 */
function createActivitiesReportPdf(params) {
    Logger.log('createActivitiesReportPdf called with: ' + JSON.stringify(params));
    let newSpreadsheetFile = null;
    try {
        const reportResult = getActivitiesReportData(params);
        if (!reportResult.success) {
            throw new Error(reportResult.message || 'ไม่สามารถดึงข้อมูลรายงานได้');
        }
        const {
            reportData,
            components,
            reportDetails
        } = reportResult;
        const {
            classLevel,
            classroom,
            teacherName
        } = reportDetails;

        if (reportData.length === 0) {
            throw new Error('ไม่พบข้อมูลสำหรับสร้างรายงาน');
        }

        const templateFile = DriveApp.getFileById(MONTHLY_REPORT_SHEET_TEMPLATE_ID);
        const newFileName = `รายงานกิจกรรมฯ_${classLevel}_${classroom}_${new Date().getTime()}`;
        newSpreadsheetFile = templateFile.makeCopy(newFileName);

        const newSS = SpreadsheetApp.openById(newSpreadsheetFile.getId());
        const sheet = newSS.getSheetByName(ACTIVITIES_REPORT_SHEET_NAME);
        if (!sheet) {
            throw new Error('ไม่พบชีตเทมเพลตสำหรับรายงานกิจกรรมฯ ที่ชื่อว่า "' + ACTIVITIES_REPORT_SHEET_NAME + '"');
        }
        sheet.clear();

        const totalColumns = 3 + components.length + 1; // ลำดับ, รหัส, ชื่อ, components, สรุป

        // 1. ส่วนหัวกระดาษ
        sheet.getRange(1, 1, 1, totalColumns).merge().setValue('แบบประเมินกิจกรรมพัฒนาผู้เรียน').setFontWeight('bold').setFontSize(18).setFontFamily("Sarabun").setHorizontalAlignment('center');
        sheet.getRange(2, 1, 1, totalColumns).merge().setValue(`ระดับชั้น ${classLevel} ห้อง ${classroom}`).setFontSize(18).setFontFamily("Sarabun").setHorizontalAlignment('center');
        sheet.getRange(3, 1, 1, totalColumns).merge().setValue(`ครูผู้ประเมิน: ${teacherName}`).setFontSize(16).setFontFamily("Sarabun").setHorizontalAlignment('center');

        // 2. ส่วนตารางข้อมูล
        const startRow = 5;
        const headerRow1Data = ['', '', ''];
        if (components.length > 0) {
            headerRow1Data.push('รายการกิจกรรม');
            for (let i = 1; i < components.length; i++) headerRow1Data.push('');
        }
        headerRow1Data.push('');
        sheet.getRange(startRow, 1, 1, headerRow1Data.length).setValues([headerRow1Data]);
        const headerRow2Data = ['ลำดับ', 'รหัสนักเรียน', 'ชื่อ-นามสกุล'];
        components.forEach(item => headerRow2Data.push(item.componentName));
        headerRow2Data.push('สรุปผล');
        sheet.getRange(startRow + 1, 1, 1, headerRow2Data.length).setValues([headerRow2Data]);

        // Merge cells for headers
        sheet.getRange(startRow, 1, 2, 1).merge();
        sheet.getRange(startRow, 2, 2, 1).merge();
        sheet.getRange(startRow, 3, 2, 1).merge();
        if (components.length > 0) {
            sheet.getRange(startRow, 4, 1, components.length).merge();
        }
        sheet.getRange(startRow, totalColumns, 2, 1).merge();

        const tableStartRow = startRow + 2;
        const tableData = reportData.map((student, index) => {
            const row = [index + 1, student.studentCode, student.studentName];
            components.forEach(item => row.push(student.results[item.id] || '-'));
            row.push(student.finalResult);
            return row;
        });

        // ตรวจสอบจำนวนแถวก่อนเขียนข้อมูล
        if (tableData.length === 0) {
            throw new Error('ไม่มีข้อมูลนักเรียนสำหรับเขียนลงในตาราง');
        }
        Logger.log(`Writing ${tableData.length} rows to table starting at row ${tableStartRow}`);

        // เขียนข้อมูลตาราง (แก้ไขจาก [tableData] เป็น tableData)
        sheet.getRange(tableStartRow, 1, tableData.length, totalColumns).setValues(tableData);

        // 3. จัดรูปแบบตาราง
        const fullTableRange = sheet.getRange(startRow, 1, tableData.length + 2, totalColumns);
        fullTableRange.setFontFamily("Sarabun").setFontSize(15).setVerticalAlignment("middle").setBorder(true, true, true, true, true, true);
        const fullHeaderRange = sheet.getRange(startRow, 1, 2, totalColumns);
        fullHeaderRange.setFontWeight('bold').setBackground('#F3F4F6').setHorizontalAlignment('center');
        if (components.length > 0) {
            sheet.getRange(startRow + 1, 4, 1, components.length).setTextRotation(90);
        }
        sheet.setRowHeight(startRow + 1, 150);
        sheet.getRange(tableStartRow, 1, tableData.length, 1).setHorizontalAlignment("center");
        sheet.getRange(tableStartRow, 2, tableData.length, 1).setHorizontalAlignment("center");
        sheet.getRange(tableStartRow, 3, tableData.length, 1).setHorizontalAlignment("left");
        sheet.getRange(tableStartRow, 4, tableData.length, components.length + 1).setHorizontalAlignment("center");
        for (let i = 0; i < tableData.length; i++) {
            sheet.setRowHeight(tableStartRow + i, 30);
        }
        const summaryCol = sheet.getRange(startRow, totalColumns, tableData.length + 2, 1);
        summaryCol.setBackground('#E0E7FF').setFontWeight('bold');

        // 4. ปรับขนาดคอลัมน์
        sheet.setColumnWidth(1, 50);
        sheet.setColumnWidth(2, 100);
        sheet.setColumnWidth(3, 220);
        if (components.length > 0) {
            for (let i = 0; i < components.length; i++) {
                sheet.setColumnWidth(4 + i, 50);
            }
        }
        sheet.setColumnWidth(4 + components.length, 90);

        // 5. ส่วนลงชื่อท้ายกระดาษ
        const signatureStartRow = tableStartRow + tableData.length + 2;
        const signatureBlockNumCols = Math.min(6, Math.floor(totalColumns / 2) - 1);

        // บล็อกซ้าย: ครูประจำวิชา/ครูที่ปรึกษา
        const leftSigStartCol = 2;
        sheet.getRange(signatureStartRow, leftSigStartCol, 1, signatureBlockNumCols)
            .merge()
            .setValue('ลงชื่อ ................................................\n(.....................................................)\nครูประจำวิชา/ครูที่ปรึกษา')
            .setWrap(true)
            .setHorizontalAlignment('center')
            .setVerticalAlignment('bottom')
            .setFontFamily("Sarabun")
            .setFontSize(12);

        // บล็อกขวา: หัวหน้าฝ่ายวิชาการ
        const rightSigStartCol = totalColumns - signatureBlockNumCols;
        sheet.getRange(signatureStartRow, rightSigStartCol, 1, signatureBlockNumCols)
            .merge()
            .setValue('ลงชื่อ ............................................\n(.................................................)\nหัวหน้าฝ่ายวิชาการ')
            .setWrap(true)
            .setHorizontalAlignment('center')
            .setVerticalAlignment('bottom')
            .setFontFamily("Sarabun")
            .setFontSize(12);

        sheet.setRowHeight(signatureStartRow, 85);

        // ตรึงแถวหัวตาราง
        sheet.setFrozenRows(startRow + 1);

        SpreadsheetApp.flush();
        const fileId = newSS.getId();

        // สร้าง URL สำหรับ export PDF
        const pdfUrl = `https://docs.google.com/spreadsheets/d/${fileId}/export?` +
            'format=pdf' +
            '&size=a4' +
            '&portrait=true' +
            '&fitw=true' +
            '&sheetnames=false&printtitle=false' +
            '&gridlines=false' +
            '&top_margin=0.25' +
            '&bottom_margin=0.25' +
            '&left_margin=0.7' +
            '&right_margin=0.35' +
            '&gid=' + sheet.getSheetId();

        newSpreadsheetFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

        Logger.log('Activities Report PDF created successfully. URL: ' + pdfUrl);
        return {
            success: true,
            url: pdfUrl
        };
    } catch (error) {
        Logger.log('Error in createActivitiesReportPdf: ' + error.stack);
        if (newSpreadsheetFile) {
            DriveApp.getFileById(newSpreadsheetFile.getId()).setTrashed(true);
        }
        return {
            success: false,
            message: 'ไม่สามารถสร้างไฟล์ PDF ได้: ' + error.message
        };
    }
}

function calculateCharacteristicGrade(percentage) {
  if (percentage >= 80) return 'ดีเยี่ยม';
  if (percentage >= 70) return 'ดี';
  if (percentage >= 50) return 'ผ่าน';
  return 'ไม่ผ่าน';
}

/**
 * โหลดเนื้อหา HTML สำหรับหน้าจัดการค่านิยมและสมรรถนะ
 */
function loadValuesCompetenciesPage() {
  Logger.log('loadValuesCompetenciesPage called');
  try {
    // เราจะสร้างไฟล์ values_competencies.html ในขั้นตอนถัดไป
    return HtmlService.createHtmlOutputFromFile('values_competencies').getContent();
  } catch (error) {
    Logger.log('Error in loadValuesCompetenciesPage: ' + error.message);
    throw new Error('ไม่สามารถโหลดหน้าจัดการค่านิยมและสมรรถนะได้');
  }
}

/**
 * บันทึกคะแนนค่านิยมและสมรรถนะลงชีต
 */
function saveValuesCompetenciesScores(records) {
  Logger.log('saveValuesCompetenciesScores called with %s records', records.length);
  if (!records || records.length === 0) {
    return { success: true, count: 0, savedRecords: [] };
  }
  try {
    var sheet = initializeSheet('values_competencies_scores');
    var allScores = sheet.getDataRange().getValues();
    var scoreMap = {};
    for (var i = 1; i < allScores.length; i++) {
      var key = allScores[i][1] + '-' + allScores[i][2]; // studentId-itemId
      scoreMap[key] = { rowIndex: i + 1, id: allScores[i][0] };
    }

    var newRows = [];
    var updatedCount = 0;
    var savedRecordsForClient = [];

    records.forEach(function(record) {
      var key = record.studentId + '-' + record.itemId;
      var existing = scoreMap[key];
      var recordId = existing ? existing.id : Utilities.getUuid();
      var rowData = [recordId, record.studentId, record.itemId, record.score, record.teacherId, new Date()];

      if (existing) {
        sheet.getRange(existing.rowIndex, 1, 1, rowData.length).setValues([rowData]);
      } else {
        newRows.push(rowData);
      }
      updatedCount++;
      savedRecordsForClient.push({id: rowData[0], studentId: rowData[1], itemId: rowData[2], score: rowData[3]});
    });

    if (newRows.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, newRows[0].length).setValues(newRows);
    }

    SpreadsheetApp.flush();
    Logger.log('Saved values/competencies scores. Total records processed: ' + updatedCount);
    return { success: true, count: updatedCount, savedRecords: savedRecordsForClient };
  } catch (e) {
    Logger.log('Error in saveValuesCompetenciesScores: ' + e.stack);
    throw new Error('การบันทึกคะแนนค่านิยมและสมรรถนะล้มเหลว: ' + e.message);
  }
}


/**
 * ดึงข้อมูลสำหรับสร้างรายงานสรุปผลค่านิยมและสมรรถนะ
 */
function getValuesCompetenciesReportData(params) {
    try {
        const { classLevel, classroom, teacherName } = params;
        if (!classLevel || !classroom) {
            throw new Error('กรุณาระบุระดับชั้นและห้องเรียน');
        }

        const studentsInClass = getData('students')
            .filter(s => s.class === classLevel && s.classroom === classroom)
            .sort((a,b) => (a.code || '').localeCompare(b.code));

        if (studentsInClass.length === 0) {
            return { success: true, reportData: [], items: [], reportDetails: params };
        }

        const studentIds = studentsInClass.map(s => s.id);
        const allScores = getData('values_competencies_scores').filter(s => studentIds.includes(s.studentId));
        const allItems = getData('values_competencies_items').sort((a, b) => a.order - b.order);

        const reportData = studentsInClass.map(student => {
            let scores = {};
            let valuesSum = 0, valuesCount = 0;
            let competenciesSum = 0, competenciesCount = 0;

            allItems.forEach(item => {
                const scoreRecord = allScores.find(s => s.studentId === student.id && s.itemId === item.id);
                if (scoreRecord && scoreRecord.score !== '') {
                    const score = parseInt(scoreRecord.score);
                    scores[item.id] = score;

                    if (item.itemGroup === 'ค่านิยมหลักของคนไทย 12 ประการ') {
                        valuesSum += score;
                        valuesCount++;
                    } else if (item.itemGroup === 'สมรรถนะสำคัญของผู้เรียน') {
                        competenciesSum += score;
                        competenciesCount++;
                    }
                }
            });
            
            let valuesResult = '-';
            if (valuesCount > 0) {
                const percentage = (valuesSum / (valuesCount * 3)) * 100;
                valuesResult = calculateCharacteristicGrade(percentage);
            }

            let competenciesResult = '-';
            if (competenciesCount > 0) {
                const percentage = (competenciesSum / (competenciesCount * 3)) * 100;
                competenciesResult = calculateCharacteristicGrade(percentage);
            }

            return {
                studentId: student.id,
                studentCode: student.code,
                studentName: student.name,
                scores: scores,
                valuesResult: valuesResult,
                competenciesResult: competenciesResult,
            };
        });

        return { success: true, reportData: reportData, items: allItems, reportDetails: params };

    } catch (e) {
        Logger.log('Error in getValuesCompetenciesReportData: ' + e.stack);
        return { success: false, message: e.message };
    }
}


/**
 * สร้างไฟล์ PDF รายงานสรุปผลค่านิยมและสมรรถนะ
 */
function createValuesCompetenciesReportPdf(params) {
    Logger.log('createValuesCompetenciesReportPdf called with: ' + JSON.stringify(params));
    let newSpreadsheetFile = null;
    try {
        const reportResult = getValuesCompetenciesReportData(params);
        if (!reportResult.success) {
            throw new Error(reportResult.message || 'ไม่สามารถดึงข้อมูลรายงานได้');
        }
        const { reportData, items, reportDetails } = reportResult;
        const { classLevel, classroom, teacherName } = reportDetails;

        if (reportData.length === 0) {
            throw new Error('ไม่พบข้อมูลสำหรับสร้างรายงาน');
        }
        
        const templateFile = DriveApp.getFileById(MONTHLY_REPORT_SHEET_TEMPLATE_ID);
        const newFileName = `รายงานค่านิยมฯ_${classLevel}_${classroom}_${new Date().getTime()}`;
        newSpreadsheetFile = templateFile.makeCopy(newFileName);
        
        const newSS = SpreadsheetApp.openById(newSpreadsheetFile.getId());
        const sheet = newSS.getSheetByName(VALUES_COMPETENCIES_REPORT_SHEET_NAME);
        if (!sheet) {
            throw new Error('ไม่พบชีตเทมเพลตสำหรับรายงานค่านิยมฯ ที่ชื่อว่า "' + VALUES_COMPETENCIES_REPORT_SHEET_NAME + '"');
        }
        sheet.clear();
        
        const valuesItems = items.filter(i => i.itemGroup === 'ค่านิยมหลักของคนไทย 12 ประการ');
        const competenciesItems = items.filter(i => i.itemGroup === 'สมรรถนะสำคัญของผู้เรียน');
        const totalColumns = 3 + items.length + 2;

        // 1. ส่วนหัวกระดาษ
        sheet.getRange(1, 1, 1, totalColumns).merge().setValue('แบบประเมินค่านิยม 12 ประการ และสมรรถนะสำคัญของผู้เรียน').setFontWeight('bold').setFontSize(18).setFontFamily("Sarabun").setHorizontalAlignment('center');
        sheet.getRange(2, 1, 1, totalColumns).merge().setValue(`ระดับชั้น ${classLevel} ห้อง ${classroom}`).setFontSize(18).setFontFamily("Sarabun").setHorizontalAlignment('center');
        sheet.getRange(3, 1, 1, totalColumns).merge().setValue(`ครูผู้ประเมิน: ${teacherName}`).setFontSize(16).setFontFamily("Sarabun").setHorizontalAlignment('center');

        // 2. ส่วนตารางข้อมูล
        const startRow = 5;
        const headerRow1Data = ['', '', ''];
        if (valuesItems.length > 0) {
            headerRow1Data.push('ค่านิยมหลักของคนไทย 12 ประการ');
            for (let i = 1; i < valuesItems.length; i++) headerRow1Data.push('');
        }
        if (competenciesItems.length > 0) {
            headerRow1Data.push('สมรรถนะสำคัญของผู้เรียน');
            for (let i = 1; i < competenciesItems.length; i++) headerRow1Data.push('');
        }
        headerRow1Data.push('', '');
        sheet.getRange(startRow, 1, 1, headerRow1Data.length).setValues([headerRow1Data]);
        
        const headerRow2Data = ['ลำดับ', 'รหัสนักเรียน', 'ชื่อ-นามสกุล'];
        items.forEach(item => headerRow2Data.push(item.itemName.replace(/(\d+\.\s)/, '')));
        headerRow2Data.push('สรุป\n(ค่านิยมฯ)', 'สรุป\n(สมรรถนะฯ)');
        sheet.getRange(startRow + 1, 1, 1, headerRow2Data.length).setValues([headerRow2Data]);

        // Merge cells
        sheet.getRange(startRow, 1, 2, 1).merge();
        sheet.getRange(startRow, 2, 2, 1).merge();
        sheet.getRange(startRow, 3, 2, 1).merge();
        let currentCol = 4;
        if (valuesItems.length > 0) {
            sheet.getRange(startRow, currentCol, 1, valuesItems.length).merge();
            currentCol += valuesItems.length;
        }
        if (competenciesItems.length > 0) {
            sheet.getRange(startRow, currentCol, 1, competenciesItems.length).merge();
        }
        sheet.getRange(startRow, totalColumns - 1, 2, 1).merge();
        sheet.getRange(startRow, totalColumns, 2, 1).merge();

        const tableStartRow = startRow + 2;
        const tableData = reportData.map((student, index) => {
            const row = [index + 1, student.studentCode, student.studentName];
            items.forEach(item => row.push(student.scores[item.id] !== undefined ? student.scores[item.id] : ''));
            row.push(student.valuesResult, student.competenciesResult);
            return row;
        });
        sheet.getRange(tableStartRow, 1, tableData.length, totalColumns).setValues(tableData);
        
        // 3. จัดรูปแบบ
        const fullTableRange = sheet.getRange(startRow, 1, tableData.length + 2, totalColumns);
        fullTableRange.setFontFamily("Sarabun").setFontSize(16).setVerticalAlignment("middle").setBorder(true, true, true, true, true, true);
        const fullHeaderRange = sheet.getRange(startRow, 1, 2, totalColumns);
        fullHeaderRange.setFontWeight('bold').setBackground('#F3F4F6').setHorizontalAlignment('center');
        if (items.length > 0) {
            sheet.getRange(startRow + 1, 4, 1, items.length).setTextRotation(90);
        }
        sheet.setRowHeight(startRow + 1, 150);
        sheet.getRange(tableStartRow, 3, tableData.length, 1).setHorizontalAlignment("left"); // Align names to left
        sheet.getRange(tableStartRow, 1, tableData.length, totalColumns).setHorizontalAlignment("center");
        for(let i = 0; i < tableData.length; i++) { sheet.setRowHeight(tableStartRow + i, 30); }
        const summaryCols = sheet.getRange(startRow, totalColumns - 1, tableData.length + 2, 2);
        summaryCols.setBackground('#E0E7FF').setFontWeight('bold');

        // 4. ปรับขนาดคอลัมน์
        sheet.setColumnWidth(1, 50);
        sheet.setColumnWidth(2, 100);
        sheet.setColumnWidth(3, 220);
        if (items.length > 0) {
           for (let i = 0; i < items.length; i++) {
               sheet.setColumnWidth(4 + i, 60); 
           }
        }
        sheet.setColumnWidth(4 + items.length, 90);
        sheet.setColumnWidth(5 + items.length, 90);

        // 5. ส่วนลงชื่อท้ายกระดาษ
        const signatureStartRow = tableStartRow + tableData.length + 2;
        const signatureBlockNumCols = 6;
        sheet.getRange(signatureStartRow, 2, 1, signatureBlockNumCols).merge().setValue('ลงชื่อ ..................................................\n(.......................................................)\nครูประจำวิชา/ครูที่ปรึกษา').setWrap(true).setHorizontalAlignment('center').setVerticalAlignment('bottom').setFontFamily("Sarabun").setFontSize(16);
        sheet.setRowHeight(signatureStartRow, 85);

        sheet.getRange(signatureStartRow, totalColumns - signatureBlockNumCols, 1, signatureBlockNumCols).merge().setValue('ลงชื่อ ..................................................\n(.......................................................)\nหัวหน้าฝ่ายวิชาการ').setWrap(true).setHorizontalAlignment('center').setVerticalAlignment('bottom').setFontFamily("Sarabun").setFontSize(16);
        
        sheet.setFrozenRows(startRow + 1);

        SpreadsheetApp.flush();
        const fileId = newSS.getId();

        // --- START: แก้ไขส่วนนี้ ---
        const pdfUrl = `https://docs.google.com/spreadsheets/d/${fileId}/export?` +
                       'format=pdf' +
                       '&size=a4' +
                       '&portrait=true' + // เปลี่ยนจาก false เป็น true เพื่อให้เป็นแนวตั้ง
                       '&fitw=true' +
                       '&sheetnames=false&printtitle=false' +
                       '&gridlines=false' +
                       '&gid=' + sheet.getSheetId();
        // --- END: แก้ไขส่วนนี้ ---
        
        newSpreadsheetFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        
        return { success: true, url: pdfUrl };
    } catch (error) {
        Logger.log('Error in createValuesCompetenciesReportPdf: ' + error.stack);
        if (newSpreadsheetFile) {
            DriveApp.getFileById(newSpreadsheetFile.getId()).setTrashed(true);
        }
        return { success: false, message: 'ไม่สามารถสร้างไฟล์ PDF ได้: ' + error.message };
    }
}

// --- START: เพิ่ม 3 ฟังก์ชันใหม่สำหรับปก ปพ.5 ---

function loadP5CoverPage() {
    Logger.log('loadP5CoverPage called');
    try {
        // เราจะสร้างไฟล์ p5_cover.html ในขั้นตอนถัดไป
        return HtmlService.createHtmlOutputFromFile('p5_cover').getContent();
    } catch (error) {
        Logger.log('Error in loadP5CoverPage: ' + error.message);
        throw new Error('ไม่สามารถโหลดหน้าสำหรับพิมพ์ปก ปพ.5 ได้');
    }
}

function getP5CoverData(dataForServer) {
    Logger.log('getP5CoverData (Fixed) called');
    try {
        const { params, appData } = dataForServer;
        const {
            academicYear, semester, classLevel, classroom, subjectId,
            teacherName, learningHours, learningArea, advisorName,
            headAcademicName, approvalDate
        } = params;
        const {
            settings, subjects, students, characteristicsScores,
            valuesCompetenciesScores, characteristicsItems, valuesCompetenciesItems,
            scoreComponents, scores,
            activityComponents, activityScores
        } = appData;

        const subject = subjects.find(s => s.id === subjectId);

        const gradeSummary = { '4': 0, '3.5': 0, '3': 0, '2.5': 0, '2': 0, '1.5': 0, '1': 0, '0': 0, 'ร': 0, 'มผ': 0, 'มส': 0 };
        const componentsForClass = scoreComponents.filter(c => c.subjectId === subjectId);
        const totalMaxScore = componentsForClass.reduce((sum, c) => sum + parseFloat(c.maxScore), 0);
        students.forEach(student => {
            // --- START: ส่วนที่แก้ไข (จุดสำคัญของปัญหานี้) ---
            // ค้นหา remark โดยกรองจาก subjectId ด้วย เพื่อให้ได้ข้อมูลที่ถูกต้องสำหรับวิชานี้
            const remarkRecord = scores.find(s =>
                s.studentId === student.id &&
                s.componentId === 'remark' &&
                s.subjectId === subjectId
            );
            const studentRemark = remarkRecord ? remarkRecord.remark : '-';
            // --- END: ส่วนที่แก้ไข ---

            if (studentRemark !== '-') {
                if (gradeSummary.hasOwnProperty(studentRemark)) {
                    gradeSummary[studentRemark]++;
                }
            } else {
                const scoresForStudent = scores.filter(s => s.studentId === student.id && componentsForClass.some(c => c.id === s.componentId));
                const totalScore = scoresForStudent.reduce((sum, s) => sum + parseFloat(s.score), 0);
                if (totalMaxScore > 0) {
                    const finalGrade = calculateGrade(totalScore, totalMaxScore, '-');
                    if (gradeSummary.hasOwnProperty(finalGrade)) {
                        gradeSummary[finalGrade]++;
                    }
                }
            }
        });

        const activitySummary = [];
        if (activityComponents && activityComponents.length > 0) {
            activityComponents.slice(0, 4).forEach(component => {
                let passCount = 0;
                let failCount = 0;
                students.forEach(student => {
                    const record = activityScores.find(s => s.studentId === student.id && s.componentId === component.id);
                    if (record) {
                        if (record.result === 'ผ่าน') {
                            passCount++;
                        } else if (record.result === 'ไม่ผ่าน') {
                            failCount++;
                        }
                    }
                });
                activitySummary.push({ name: component.componentName, pass: passCount, fail: failCount });
            });
        }

        const evaluationSummary = {
            characteristics: { 'ดีเยี่ยม': 0, 'ดี': 0, 'ผ่าน': 0, 'ไม่ผ่าน': 0 },
            reading: { 'ดีเยี่ยม': 0, 'ดี': 0, 'ผ่าน': 0, 'ไม่ผ่าน': 0 },
            values: { 'ดีเยี่ยม': 0, 'ดี': 0, 'ผ่าน': 0, 'ไม่ผ่าน': 0 },
            competencies: { 'ดีเยี่ยม': 0, 'ดี': 0, 'ผ่าน': 0, 'ไม่ผ่าน': 0 },
        };
        students.forEach(student => {
            const charItems = characteristicsItems.filter(i => i.itemGroup === 'คุณลักษณะอันพึงประสงค์');
            const readingItems = characteristicsItems.filter(i => i.itemGroup === 'การอ่าน คิดวิเคราะห์ และเขียน');
            let charSum = 0, charCount = 0;
            charItems.forEach(item => { const scoreRecord = characteristicsScores.find(s => s.studentId === student.id && s.itemId === item.id); if (scoreRecord && scoreRecord.score !== '') { charSum += parseInt(scoreRecord.score); charCount++; }});
            let readingSum = 0, readingCount = 0;
            readingItems.forEach(item => { const scoreRecord = characteristicsScores.find(s => s.studentId === student.id && s.itemId === item.id); if (scoreRecord && scoreRecord.score !== '') { readingSum += parseInt(scoreRecord.score); readingCount++; }});
            if (charCount > 0) { const grade = calculateCharacteristicGrade((charSum / (charCount * 3)) * 100); if (evaluationSummary.characteristics.hasOwnProperty(grade)) { evaluationSummary.characteristics[grade]++; }}
            if (readingCount > 0) { const grade = calculateCharacteristicGrade((readingSum / (readingCount * 3)) * 100); if (evaluationSummary.reading.hasOwnProperty(grade)) { evaluationSummary.reading[grade]++; }}
            const valuesItems = valuesCompetenciesItems.filter(i => i.itemGroup === 'ค่านิยมหลักของคนไทย 12 ประการ');
            const competenciesItems = valuesCompetenciesItems.filter(i => i.itemGroup === 'สมรรถนะสำคัญของผู้เรียน');
            let valuesSum = 0, valuesCount = 0;
            valuesItems.forEach(item => { const scoreRecord = valuesCompetenciesScores.find(s => s.studentId === student.id && s.itemId === item.id); if (scoreRecord && scoreRecord.score !== '') { valuesSum += parseInt(scoreRecord.score); valuesCount++; }});
            let competenciesSum = 0, competenciesCount = 0;
            competenciesItems.forEach(item => { const scoreRecord = valuesCompetenciesScores.find(s => s.studentId === student.id && s.itemId === item.id); if (scoreRecord && scoreRecord.score !== '') { competenciesSum += parseInt(scoreRecord.score); competenciesCount++; }});
            if (valuesCount > 0) { const grade = calculateCharacteristicGrade((valuesSum / (valuesCount * 3)) * 100); if (evaluationSummary.values.hasOwnProperty(grade)) { evaluationSummary.values[grade]++; }}
            if (competenciesCount > 0) { const grade = calculateCharacteristicGrade((competenciesSum / (competenciesCount * 3)) * 100); if (evaluationSummary.competencies.hasOwnProperty(grade)) { evaluationSummary.competencies[grade]++; }}
        });

        const coverData = {
            schoolName: settings.school_name || "ชื่อสถานศึกษา",
            schoolArea: settings.school_area || "เขตพื้นที่การศึกษา",
            schoolTypeName: settings.school_type_name || "โรงเรียน",
            academicYear: academicYear,
            semester: semester,
            subjectCode: subject ? subject.code : '-',
            subjectName: subject ? subject.name : '-',
            learningArea: learningArea,
            learningHours: learningHours,
            classLevel: classLevel,
            teacherName: teacherName,
            advisorName: advisorName,
            headAcademicName: headAcademicName,
            directorName: settings.director_name || "ชื่อผู้อำนวยการ",
            studentTotal: students.length,
            evaluationSummary: evaluationSummary,
            gradeSummary: gradeSummary,
            activitySummary: activitySummary,
            approvalDate: approvalDate,
            logoUrl: settings.logo_url // ส่ง URL ไปตรงๆ
        };

        if (settings.logo_url) {
            coverData.logoBase64 = getImageAsBase64(settings.logo_url);
        } else {
            coverData.logoBase64 = '';
        }

        return { success: true, coverData: coverData };

    } catch (e) {
        Logger.log('Error in getP5CoverData: ' + e.stack);
        return { success: false, message: 'เกิดข้อผิดพลาดฝั่งเซิร์ฟเวอร์: ' + e.message };
    }
}

function createP5CoverPdf(params) {
    Logger.log('createP5CoverPdf (Revised Version) called with: ' + JSON.stringify(params));
    let newSpreadsheetFile = null;
    try {
        // START: ส่วนดึงข้อมูล (เหมือนเดิม)
        const studentDataForServer = {
            params: params,
            appData: {
                settings: getSettings(),
                subjects: getData('subjects'),
                students: getData('students').filter(s => s.class === params.classLevel && s.classroom === params.classroom),
                characteristicsScores: getData('characteristics_scores'),
                valuesCompetenciesScores: getData('values_competencies_scores'),
                characteristicsItems: getData('characteristics_items'),
                valuesCompetenciesItems: getData('values_competencies_items'),
                scoreComponents: getData('score_components'),
                scores: getData('scores'),
                activityComponents: getData('activity_components'),
                activityScores: getData('activity_scores')
            }
        };

        const result = getP5CoverData(studentDataForServer);
        if (!result.success) {
            throw new Error(result.message);
        }
        const data = result.coverData;

        const templateFile = DriveApp.getFileById(MONTHLY_REPORT_SHEET_TEMPLATE_ID);
        const newFileName = `ปก ปพ.5_${data.subjectName}_${data.classLevel}`;
        newSpreadsheetFile = templateFile.makeCopy(newFileName);

        const newSS = SpreadsheetApp.openById(newSpreadsheetFile.getId());
        const sheet = newSS.getSheetByName(P5_COVER_SHEET_NAME);
        if (!sheet) {
            throw new Error('ไม่พบชีตเทมเพลตสำหรับปก ปพ.5');
        }
        // END: ส่วนดึงข้อมูล (เหมือนเดิม)

        sheet.clear(); // ล้างชีตเทมเพลตให้สะอาดก่อนเริ่ม

        // =========================================================================
        // START: ส่วนแก้ไขปัญหา
        // =========================================================================

        // --- แก้ไขข้อที่ 3: ขยับเนื้อหาทั้งหมดลงมา ---
        // โดยการแทรกแถวว่าง 2 แถวที่ด้านบนสุดของชีต
        sheet.insertRowsBefore(1, 2);
        sheet.setRowHeight(1, 20); // กำหนดความสูงของขอบบน
        sheet.setRowHeight(2, 20);

        // --- แก้ไขข้อที่ 2: ทำให้โลโก้แสดงผลแน่นอน ---
        // เปลี่ยนจากการใช้สูตร =IMAGE() มาเป็นการแทรกรูปภาพโดยตรง ซึ่งเสถียรกว่ามาก
        if (data.logoUrl) {
            try {
                const blob = UrlFetchApp.fetch(data.logoUrl).getBlob();
                // แทรกรูปภาพลงในเซลล์ C3 และปรับขนาดให้พอดีกับเซลล์
                const image = sheet.insertImage(blob, 'C', 3); 
                image.setHeight(80); // กำหนดความสูงของโลโก้
                // จัดตำแหน่งรูปภาพให้อยู่กึ่งกลางเซลล์ที่ถูกรวมไว้
                const logoRange = sheet.getRange('A3:M3');
                logoRange.merge().setHorizontalAlignment('center').setVerticalAlignment('middle');
                sheet.setRowHeight(3, 85); // กำหนดความสูงของแถวที่ใส่โลโก้
            } catch (e) {
                Logger.log('Could not fetch or insert logo from URL: ' + data.logoUrl + '. Error: ' + e.message);
                // หากดึงรูปไม่ได้ ให้แสดงเป็นข้อความแทน
                sheet.getRange('A3:M3').merge().setValue('[ ไม่สามารถโหลดโลโก้ได้ ]').setHorizontalAlignment('center');
            }
        }
        
        // ข้อมูลหลักของโรงเรียนและรายวิชา (ปรับตำแหน่งแถวลงมา)
        const headerStartRow = 4; // เริ่มจากแถวที่ 4
        sheet.getRange(headerStartRow, 1, 1, 13).merge().setValue(data.schoolName).setFontWeight('bold').setFontSize(22);
        sheet.getRange(headerStartRow + 1, 1, 1, 13).merge().setValue(data.schoolArea).setFontSize(18);
        sheet.getRange(headerStartRow + 2, 1, 1, 13).merge().setValue(`สมุดบันทึกการพัฒนาคุณภาพผู้เรียน (ปพ.๕)`).setFontWeight('bold').setFontSize(20);
        
        // จัดรูปแบบหัวกระดาษ
        sheet.getRange(headerStartRow, 1, 3, 13).setFontFamily("Sarabun").setHorizontalAlignment('center');

        // ข้อมูลรายละเอียดรายวิชา (ปรับตำแหน่งแถวลงมา)
        const detailsStartRow = headerStartRow + 4; // เว้น 1 แถว
        sheet.getRange(`C${detailsStartRow}`).setValue(data.classLevel);
        sheet.getRange(`F${detailsStartRow}`).setValue(data.semester);
        sheet.getRange(`I${detailsStartRow}`).setValue(data.academicYear);
        sheet.getRange(`C${detailsStartRow+1}`).setValue(data.subjectName);
        sheet.getRange(`F${detailsStartRow+1}`).setValue(data.subjectCode);
        sheet.getRange(`I${detailsStartRow+1}`).setValue(data.learningHours);
        sheet.getRange(`C${detailsStartRow+2}`).setValue(data.learningArea);
        sheet.getRange(`E${detailsStartRow+3}`).setValue(data.teacherName);
        sheet.getRange(`I${detailsStartRow+3}`).setValue(data.advisorName);
        
        // ข้อมูลสรุป (ปรับตำแหน่งแถวลงมา)
        const summaryStartRow = detailsStartRow + 5;
        sheet.getRange(`H${summaryStartRow}`).setValue(data.studentTotal);
        
        const gradeSummary = data.gradeSummary;
        const gradeSummaryRow = summaryStartRow + 6;
        sheet.getRange(`F${gradeSummaryRow}`).setValue(gradeSummary['0']);
        sheet.getRange(`G${gradeSummaryRow}`).setValue(gradeSummary['1']);
        sheet.getRange(`H${gradeSummaryRow}`).setValue(gradeSummary['1.5']);
        sheet.getRange(`I${gradeSummaryRow}`).setValue(gradeSummary['2']);
        sheet.getRange(`J${gradeSummaryRow}`).setValue(gradeSummary['2.5']);
        sheet.getRange(`K${gradeSummaryRow}`).setValue(gradeSummary['3']);
        sheet.getRange(`L${gradeSummaryRow}`).setValue(gradeSummary['3.5']);
        sheet.getRange(`M${gradeSummaryRow}`).setValue(gradeSummary['4']);

        const evalSummary = data.evaluationSummary;
        const evalStartRow = summaryStartRow + 8;
        sheet.getRange(`F${evalStartRow}`).setValue(evalSummary.characteristics['ไม่ผ่าน']);
        sheet.getRange(`G${evalStartRow}`).setValue(evalSummary.characteristics['ผ่าน']);
        sheet.getRange(`H${evalStartRow}`).setValue(evalSummary.characteristics['ดี']);
        sheet.getRange(`I${evalStartRow}`).setValue(evalSummary.characteristics['ดีเยี่ยม']);
        sheet.getRange(`F${evalStartRow+1}`).setValue(evalSummary.reading['ไม่ผ่าน']);
        sheet.getRange(`G${evalStartRow+1}`).setValue(evalSummary.reading['ผ่าน']);
        sheet.getRange(`H${evalStartRow+1}`).setValue(evalSummary.reading['ดี']);
        sheet.getRange(`I${evalStartRow+1}`).setValue(evalSummary.reading['ดีเยี่ยม']);
        sheet.getRange(`F${evalStartRow+3}`).setValue(evalSummary.values['ไม่ผ่าน']);
        sheet.getRange(`G${evalStartRow+3}`).setValue(evalSummary.values['ผ่าน']);
        sheet.getRange(`H${evalStartRow+3}`).setValue(evalSummary.values['ดี']);
        sheet.getRange(`I${evalStartRow+3}`).setValue(evalSummary.values['ดีเยี่ยม']);
        sheet.getRange(`F${evalStartRow+5}`).setValue(evalSummary.competencies['ไม่ผ่าน']);
        sheet.getRange(`G${evalStartRow+5}`).setValue(evalSummary.competencies['ผ่าน']);
        sheet.getRange(`H${evalStartRow+5}`).setValue(evalSummary.competencies['ดี']);
        sheet.getRange(`I${evalStartRow+5}`).setValue(evalSummary.competencies['ดีเยี่ยม']);

        // --- แก้ไขข้อที่ 1: ข้อความทับเส้น และจัดวางลายเซ็นใหม่ ---
        // กำหนดแถวเริ่มต้นสำหรับส่วนลงนาม (เว้นระยะจากตารางสรุป)
        const signatureStartRow = evalStartRow + 7;
        const directorTitle = `ผู้อำนวยการ${data.schoolTypeName}`;
        
        // ครูผู้สอน (ซ้าย)
        sheet.getRange(signatureStartRow, 2, 3, 4).merge(); // B...:E...
        sheet.getRange(signatureStartRow, 2)
              .setValue(`ลงชื่อ ................................................\n\n( ${data.teacherName} )\nครูผู้สอน/ครูประจำวิชา`)
              .setWrap(true)
              .setHorizontalAlignment('center')
              .setVerticalAlignment('middle'); // ใช้ middle เพื่อให้ข้อความอยู่กลางแนวตั้งของพื้นที่ที่รวมไว้

        // หัวหน้าวิชาการ (ขวา)
        sheet.getRange(signatureStartRow, 8, 3, 4).merge(); // H...:K...
        sheet.getRange(signatureStartRow, 8)
              .setValue(`ลงชื่อ ................................................\n\n( ${data.headAcademicName || '.............................................'} )\nหัวหน้าวิชาการ`)
              .setWrap(true)
              .setHorizontalAlignment('center')
              .setVerticalAlignment('middle');

        // ปรับความสูงของแถวสำหรับลายเซ็นชุดบน
        sheet.setRowHeight(signatureStartRow, 25);
        sheet.setRowHeight(signatureStartRow+1, 25);
        sheet.setRowHeight(signatureStartRow+2, 25);

        // ผู้อำนวยการ (ล่าง-กลาง)
        const directorSigRow = signatureStartRow + 4; // เว้นระยะ 1 แถว
        sheet.getRange(directorSigRow, 5, 3, 4).merge(); // E...:H...
        sheet.getRange(directorSigRow, 5)
              .setValue(`ลงชื่อ ................................................\n\n( ${data.directorName} )\n${directorTitle}`)
              .setWrap(true)
              .setHorizontalAlignment('center')
              .setVerticalAlignment('middle');
        
        // ปรับความสูงของแถวสำหรับลายเซ็นผู้อำนวยการ
        sheet.setRowHeight(directorSigRow, 25);
        sheet.setRowHeight(directorSigRow+1, 25);
        sheet.setRowHeight(directorSigRow+2, 25);
        
        // วันที่อนุมัติ
        if (data.approvalDate) {
            try {
                const approvalDate = new Date(data.approvalDate);
                const monthNames = ["มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"];
                const approvalDateText = `วันที่ ${approvalDate.getDate()} เดือน ${monthNames[approvalDate.getMonth()]} พ.ศ. ${approvalDate.getFullYear() + 543}`;
                sheet.getRange(directorSigRow + 3, 5, 1, 4).merge().setValue(approvalDateText).setHorizontalAlignment('center').setFontSize(14);
            } catch(e) { /* ไม่ทำอะไรถ้าวันที่ผิดพลาด */ }
        }

        // =========================================================================
        // END: ส่วนแก้ไขปัญหา
        // =========================================================================

        SpreadsheetApp.flush();
        const fileId = newSS.getId();
        
        const pdfUrl = `https://docs.google.com/spreadsheets/d/${fileId}/export?` +
            'format=pdf' + '&size=a4' + '&portrait=false' +
            '&fitw=true' + '&sheetnames=false&printtitle=false' +
            '&gridlines=false' + '&gid=' + sheet.getSheetId();

        newSpreadsheetFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

        return { success: true, url: pdfUrl };

    } catch (error) {
        Logger.log('Error in createP5CoverPdf: ' + error.stack);
        if (newSpreadsheetFile) {
            // ในกรณีที่เกิดข้อผิดพลาด ให้ลบไฟล์ที่สร้างขึ้นเพื่อไม่ให้รก Drive
            DriveApp.getFileById(newSpreadsheetFile.getId()).setTrashed(true);
        }
        return { success: false, message: 'ไม่สามารถสร้างไฟล์ปก ปพ.5 ได้: ' + error.message };
    }
}