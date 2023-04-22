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
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Classroom Sync')
    .addItem('Open Sidebar', 'showSidebar')
    .addToUi();
}

/**
 * initiates the sidebar from index.html
 *  */
function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('index').setTitle('Classroom Sync').setWidth(500);
  SpreadsheetApp.getUi().showSidebar(html);
}

/** return the selected range from active sheet as JSON
  *  **/
function getSelectedRange2() {
  var selected = SpreadsheetApp.getActiveSheet().getActiveRangeList(); // Gets the selected range
  var ranges = selected.getRanges();//.forEach(function(e){str = e.getA1Notation(), ranges.push(e)}); 
  str = [];
  for (var i = 0; i < ranges.length; i++) {
    str.push(ranges[i].getA1Notation());
  }
  return JSON.stringify(str);
}
/**
 * 
 */
function getCourseList() {
  var courses = [];
  var courseIds = [];
  try {
    Classroom.Courses.list({ "courseStates": ["ACTIVE"] }).courses.
      forEach(function (e) {
        if (e.courseState != 'ARCHIVED') {
          Logger.log(e);
          courses.push(e.name);
          courseIds.push(e.id);
        }
      });
  } catch (e) {
    throw ('cannot read courses', e.error);
  }
  Logger.log(courses);
  var documentProperties = PropertiesService.getDocumentProperties();
  documentProperties.setProperty("COURSE_IDS", JSON.stringify(courseIds));
  return JSON.stringify(courses);
}


/**
 * given course name returns its id
 */
function getCourseIdFromName(courseName) {
  var course = Classroom.Courses.list().courses.find(obj => {
    return obj.name == courseName;
  });
  if (!course) throw ('uff');
  Logger.log(" final course.name: " + course.name)
  return course.id;
}

/**
 * given work title, and course id, returns work id.
 */
function getworkIdFromWorkTitle(id, courseWorkTitle) {

  var courseWork = Classroom.Courses.CourseWork.list(id)
    .courseWork.find(obj => { return obj.name == courseWorkTitle; });

  Logger.log("found course work name:" + courseWork.title);
  return courseWork.Id;
}



/**
 * given range in A1notation
 *  converts it to indices..
 * returns the following:
 *   var data = {
    "range": range,
    "emailCol": emailCol,
    "gradeCol": gradeCol,
    "startIndex": startIndex,
    "endIndex": endIndex
  }
 */
function convertRangeToIndex(range) {
  range = range.trim();
  var c = range.indexOf(",");
  if (c == -1) {
    c = range.indexOf(":");
    if (c == -1) {
      throw ('no colon, no comma in the range string');
    }

  }

  firstPart = range.slice(0, c).replaceAll(/[ \[\],:]/g, "");
  secondPart = range.slice(c, range.length).replaceAll(/[ \[\],:]/g, "");

  emailCol = firstPart.charCodeAt(0) - "A".charCodeAt(0);

  startIndex = (firstPart.length > 1) ?
    parseInt(firstPart.slice(1, firstPart.lenthg), 10) : 1;
  endIndex = (secondPart.length > 1) ?
    parseInt(secondPart.slice(1, secondPart.length), 10) : 1000;//Set to a Maksimum 1000.

  gradeCol = secondPart.charCodeAt(0) - "A".charCodeAt(0);
  var data = {
    "range": range,
    "emailCol": emailCol,
    "gradeCol": gradeCol,
    "startIndex": startIndex,
    "endIndex": endIndex
  }
  return data;
}

/**
 * given title, creates new assignment
 */
function createCourseDetails(workTitle) {
  workDetails = {
    title: workTitle,
    state: "PUBLISHED",
    description: "",
    maxPoints: 100,
    /*materials: [
      {
        driveFile:{
        driveFile: {
          id: "fileID", 
          title: "Sample Document"
  
        },
        shareMode: "STUDENT_COPY"
        }
  
      }
      ],
      */
    workType: "ASSIGNMENT"
  };



  return workDetails;
}



/**
 * the incoming data = {
      'course':course,
      'assignment': assignment,
      'range':range,
      'createAssignment": createAssignment,
      'assignedGrade': assignedGrade
      };
 */

function uploadGradesToClassroom(strdata) {
  /*data = {
    'course': 'final projects',
    'assignment': 'test12345',
    'range': 'A1:B100',
    'createAssignment': 0,
    'assignedGrade': 1
  };
  strdata = JSON.stringify(data);*/
  /////////////
  var ss = SpreadsheetApp.getActive();

  //var name = (new Date()).toLocaleString();
  //SpreadsheetApp.getActiveSpreadsheet().insertSheet("FailedUpdates"+name);


  Logger.log(strdata);
  data = JSON.parse(strdata);
  Logger.log(data);
  courseId = getCourseIdFromName(data.course);

  //find course assignment
  var work = null;
  if (data.createAssignment) {
    workDetails = createCourseDetails(data.assignment)
    work = Classroom.Courses.CourseWork.create(workDetails, courseId);
  } else {
    work = Classroom.Courses.CourseWork
      .list(courseId)
      .courseWork.find(obj => {
        return obj.title == data.assignment
      });
  }

  //convert A1notation to indices
  data.indices = { startIndex: -1, endIndex: -1, emailCol: -1, gradeCol: -1 };
  data.indices = convertRangeToIndex(data.range);
  Logger.log(work);
  uploadChosenGrades(courseId, work, data.indices);
  //ss.getActiveSheet().getCurrentCell().setValue('uffffffffff');

  return strdata;

}

/** 
 * This reads spread sheet data row by row, 
 * searches student and submission and updates its grade by calling doPatch..
 */
function uploadChosenGrades(courseId, work, indices) {
  if (courseId == null || work == null) {
    throw ('could find course work, course id and course work id null');
  }
  sheet = SpreadsheetApp.getActive().getActiveSheet();
  // This represents ALL the data
  Logger.log(indices.startIndex + "-" + indices.emailCol + "-" + indices.endIndex);
  var emails = sheet.getRange(indices.startIndex, indices.emailCol + 1, indices.endIndex).getValues();
  var grades = sheet.getRange(indices.startIndex, indices.gradeCol + 1, indices.endIndex).getValues();

  var studentSubmissions = Classroom.Courses.CourseWork.
    StudentSubmissions.list(courseId, work.id).studentSubmissions;

  for (var i = 0; i < emails.length; i++) {
    email = emails[i][0];
    grade = grades[i][0];
    Logger.log('read email-grade:' + email + '-' + grade);

    if (email == null || email == '' || null == grade) continue;

    Logger.log('working on email-grade:' + email + '-' + grade);



    var student = Classroom.Courses.Students.get(courseId, email);
    if (student == null) {
      throw ('not found student id for email-grade:' +
        email + '-' + grade);
    }
    var submission = studentSubmissions.find(obj => {
      return obj.userId == student.userId
    });

    if (submission == null) {
      throw ('not found submission id for email-grade:' +
        email + '-' + grade);
    }

    doPatchOnDraftGrade(courseId, work.id, submission.id, grade);
    Logger.log("updated grade for " + email + grade);

  }
}

/**
 * updates the grade of a given submission
 */
function doPatchOnDraftGrade(id, workId, submissionId, grade,
  upDate = { 'updateMask': 'draftGrade' }) {
  var studentSubmission = Classroom.newStudentSubmission();
  var studentSubmission = {
    'draftGrade': grade,
    'assignedGrade': grade
  };

  Classroom.Courses.CourseWork.StudentSubmissions
    .patch(
      studentSubmission,
      id,
      workId,
      submissionId,
      upDate
    );
}
