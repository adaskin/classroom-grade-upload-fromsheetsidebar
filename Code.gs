/**
 * update grades from google spreadsheet to google classroom: 
 * it does not include doGet, doPost...
 * it is written for personal use to just upload grades of my courses, it may include a few bugs:
 * Creates a spreadsheet menu which creates a sidebar 
 * to specify email columns and grades and upload them to specified assignment.
 * 
 * The code is mostly self explanatory, but feel free to shoot me email...
 * adaskin,2023
 */

/*
--------------------
functions to initiate addOn
--------------------
*/
/**
 * installs addon
 */
function onInstall(e) {
  onOpen(e);
}

/**
 * runs on open
 */
function onOpen(e) {
  var ui = SpreadsheetApp.getUi();
  ui.createAddonMenu()
    .addItem('Upload grades to classroom', 'showSidebar')
    .addToUi();
}

/**
 * initiates the sidebar from index.html
 *  */
function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('index').setTitle('Upload grades to classroom').setWidth(500);
  SpreadsheetApp.getUi().showSidebar(html);
}


/*
--------------------
Sidebar main functions that responds to requests
--------------------
*/

/** return the selected range from active sheet as JSON
  *  **/
function getSelectedRangeAsJSON() {
  var selected = SpreadsheetApp.getActiveSheet().getActiveRangeList(); // Gets the selected range
  var ranges = selected.getRanges();//.forEach(function(e){str = e.getA1Notation(), ranges.push(e)}); 
  str = [];
  for (var i = 0; i < ranges.length; i++) {
    str.push(ranges[i].getA1Notation());
  }
  return JSON.stringify(str);
}

/**
 * returns course list
 * from classroom names as JSON string 
 * saves courseID{}&names for future uses..
 */
function getCourseList() {
  var courses = [];
  var courseIds = [];
  try {
    Classroom.Courses.list({ "courseStates": ["ACTIVE"] }).courses.
      forEach(function (e) {
        if (e.courseState != 'ARCHIVED') {
          courses.push(e.name);
          courseIds.push(e.id);
        }
      });
  } catch (e) {
    throw ('cannot read courses', e.error);
  }
  Logger.log(courses);
  // var documentProperties = PropertiesService.getDocumentProperties();
  // documentProperties.setProperty("COURSE_IDS", JSON.stringify(courseIds));
  return JSON.stringify(courses);
}

/**
 * strdata = '{
      'course':'courseTitle',
      'assignment': 'assignmentTitle',
      'range':'A1:B100',
      'createAssignment': 0,
      'returnGrades': 0,
      };'
 */
function uploadGradesToClassroom(strdata) {

  var ss = SpreadsheetApp.getActive();

  Logger.log(strdata);
  data = JSON.parse(strdata);
  Logger.log(data);
  courseId = getCourseIdFromName(data.course);

  //create new course assignment
  if (data.createAssignment) {
    workDetails = createCourseDetails(data.assignment)
    Classroom.Courses.CourseWork.create(workDetails, courseId);
  }
  //find course assignment
  var work = Classroom.Courses
    .CourseWork
    .list(courseId)
    .courseWork.find(obj => {
      return obj.title == data.assignment
    });
  if (courseId == null || work == null) {
    throw ('could find course work, course id or course work null');
  }
  //convert A1notation to indices
  data.indices = convertRangeToIndices(data.range);
  Logger.log(data.indices);
  Logger.log(work);
  if (data.returnGrades == 1) {
    data.upDateMask = { 'updateMask': 'draftGrade. assignedGrade' };
  } else {
    data.upDateMask = { 'updateMask': 'draftGrade' };
  }
  var str = readGradesFromSheetToClassroom(courseId, work, data);
  //ss.getActiveSheet().getCurrentCell().setValue('uffffffffff');
  return str;

}

/** 
 * This reads spread sheet data row by row,
 * searches student and submission and updates its grade by calling doPatch..
 * @courseId
 * @work a single instance of CourseWork
 * @data {
      'course':'courseTitle',
      'assignment': 'assignmentTitle',
      'range':'A1:B100',
      'indices':
        [{
          "range": 'A1:B100',
          "emailCol": 0,
          "gradeCol": 1,
          "startIndex": 1,
          "endIndex": 100
        }],
      'createAssignment': 0,
      'returnGrades': 0,
      };
 */
function readGradesFromSheetToClassroom(courseId, work, data) {
  if (courseId == null || work == null) {
    throw ('could find course work, course id or course work null');
  }
  sheet = SpreadsheetApp.getActive().getActiveSheet();
  indices = data.indices;
  // This represents ALL the data
  Logger.log(indices.startIndex + "-" + indices.emailCol + "-" + indices.endIndex);
  var emails = sheet
    .getRange(indices.startIndex, indices.emailCol + 1, indices.endIndex)
    .getValues();
  var grades = sheet
    .getRange(indices.startIndex, indices.gradeCol + 1, indices.endIndex)
    .getValues();

  var studentSubmissions = Classroom.Courses.CourseWork
    .StudentSubmissions.list(courseId, work.id)
    .studentSubmissions;


  for (var i = 0; i < emails.length; i++) {
    email = emails[i][0];
    grade = grades[i][0];
    Logger.log('read email-grade:' + email + '-' + grade);

    if (email == null || email == '' || email.indexOf('@') == -1 || grade == null) {
      Logger.log('No email, skipped ' + email);
      continue;
    }

    //get student by email
    var student = Classroom.Courses.Students.get(courseId, email);
    if (student == null) {
      throw ('not found student id for email:' + email);
    }
    //find student's submission
    var submission = studentSubmissions.find(obj => {
      return obj.userId == student.userId
    });

    if (submission == null) {
      throw ('not found submission id for email:' + email);
    }
    //patch to student's grade to classroom

    doPatchOnDraftGrade(courseId, work.id, submission.id, grade,);
    Logger.log("updated grade for " + email + grade);
  }
}

