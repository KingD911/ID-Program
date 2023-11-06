const SheetJS = require('xlsx');

// Read the Excel sheet
const workbook = SheetJS.readFile('~$Students.xlsx');

// Get the first worksheet in the workbook
const worksheet = workbook.Sheets[workbook.SheetNames];

// Extract the student information
const students = [];
for (let i = 1; i <= worksheet.getLastRow(); i++) {
  const student = {
    first_name: worksheet.getCell(i, 1).getValue(),
    last_name: worksheet.getCell(i, 2).getValue(),
    date_of_birth: worksheet.getCell(i, 3).getValue(),
    grade: worksheet.getCell(i, 4).getValue(),
    year: worksheet.getCell(i, 5).getValue(),
    student_id: worksheet.getCell(i, 6).getValue(),
  };
  students.push(student);
}

// Create `p` tags for each student
const pTags = [];
for (const student of students) {
  const pTag = document.createElement('p');
  pTag.textContent = `${student.name}: ${student.first_name} ${student.last_name}, ${student.date_of_birth}, ${student.grade}, ${student.year}, ${student.student_id}`;
  pTags.push(pTag);
}

// Append the `p` tags to the document body
document.body.appendChild(pTags);
