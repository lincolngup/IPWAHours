// spreadsheet column numbers
const COLUMN_FIRST_NAME = 22;
const COLUMN_LAST_NAME = 23;
const COLUMN_MILES_HIKED = 42;
const COLUMN_HOURS = 43;

const inputGroupFile = document.getElementById('inputGroupFile');

function Person(
  firstName,
  lastName,
  milesHiked,
  hours
) {
  this.firstName = firstName;
  this.lastName = lastName;
  this.milesHiked = milesHiked;
  this.hours = hours;
} // Person
let people = [];

const importFile = (event) => {
  var file = event.target.files[0];

  if (file) {
    let fileReader = new FileReader();
    // perform the onload when a file is selected
    fileReader.onload = progressEvent => {
      let fileContents = processExcel(progressEvent.target.result);
      // the worksheet name is the first object property
      let worksheetName = Object.keys(fileContents)[0];
      extractNames(fileContents[worksheetName]);
    }
    fileReader.readAsArrayBuffer(file);
  } else {
    console.error("importFile(): Failed to load file");
  }
} // importFile()

inputGroupFile.addEventListener('change', importFile);

const processExcel = (excelData) => {
  let workbook = XLSX.read(excelData, {
    type: 'binary'
  });

  let firstSheet = workbook.SheetNames[0];
  let jsonData = convertToJson(workbook);
  return jsonData;
}; // processExcel()

const convertToJson = (workbook) => {
  const HEADER_ROW = 1; // row 1 is assumed to be the header row
  let result = {};

  workbook.SheetNames.forEach( (sheetName) => {
    let rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {
      header: HEADER_ROW
    });
    if (rows.length) {
      result[sheetName] = rows;
    } // if
  });
  return result;
}; // convertToJson()

const extractNames = (rows) => {
  for (let row = 0; row < rows.length; row++) {
    const index = people.findIndex(entry =>
      entry.firstName === rows[row][COLUMN_FIRST_NAME] &&
      entry.lastName === rows[row][COLUMN_LAST_NAME]
    );
    // add the name if it wasn't found
    if (index === -1) {
      people.push(new Person(
        rows[row][COLUMN_FIRST_NAME],
        rows[row][COLUMN_LAST_NAME],
        rows[row][COLUMN_MILES_HIKED],
        rows[row][COLUMN_HOURS]
      ));
    } else {
      people[index].milesHiked += rows[row][COLUMN_MILES_HIKED];
      people[index].hours += rows[row][COLUMN_HOURS];
    } // if
  } // for
  console.log(people);
}; // extractNames()