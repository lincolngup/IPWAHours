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
  return result; //JSON.stringify(result, 2, 2);
}; // convertToJson()

const extractNames = (rows) => {
  const LAST_NAME = 23;
  const FIRST_NAME = 22;
  const MILES_HIKED = 42;
  const HOURS = 42;

  for (let row = 0; row < rows.length; row++) {
    people.push(new Person(
      rows[row][FIRST_NAME],
      rows[row][LAST_NAME],
      rows[row][MILES_HIKED],
    rows[row][HOURS]
    ));
  } // for
  console.log(people);
}; // extractNames()