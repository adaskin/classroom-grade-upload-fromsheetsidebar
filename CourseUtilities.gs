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

/**
 * given course name returns its id
 */
function getCourseIdFromName(courseName) {
  var course = Classroom.Courses.list().courses.find(obj => {
    return obj.name == courseName;
  });
  if (!course) throw ('course name not found');
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
 * given title, creates new assignment details for a given title
 */
function createCourseDetails(workTitle){
  var workDetails={
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
  }
  return workDetails;
}
