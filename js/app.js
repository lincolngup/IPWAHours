// spreadsheet column numbers
const COLUMN_FIRST_NAME = 22;
const COLUMN_LAST_NAME = 23;
const COLUMN_MILES_HIKED = 42;
const COLUMN_HOURS = 43;
const HIKING_PARTNER = 25;
const HIKING_PARTNER_NAMES = 26;

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
      buildTable();
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
  // start at row 1 to ignore the header row
  for (let row = 1; row < rows.length; row++) {
    // look for the name
    const index = people.findIndex(entry =>
      // compare in lowecase to capture case differences between entries
      entry.firstName.toLowerCase() === rows[row][COLUMN_FIRST_NAME].toLowerCase() &&
      entry.lastName.toLowerCase() === rows[row][COLUMN_LAST_NAME].toLowerCase()
    );
    addPerson(
      rows[row][COLUMN_FIRST_NAME],
      rows[row][COLUMN_LAST_NAME],
      rows[row][COLUMN_MILES_HIKED],
      rows[row][COLUMN_HOURS],
      index
    );

    if (rows[row][HIKING_PARTNER]?.toLowerCase() === 'other ipwa volunteers') {
      rows[row][HIKING_PARTNER_NAMES].split('\n')
        .forEach((fullName) => {
          // break apart the full name into first name and last name
          const name = fullName.split(' ');
          // look for the name
          const index = people.findIndex(entry =>
            // compare in lowecase to capture case differences between entries
            entry.firstName.toLowerCase() === name[0].toLowerCase() &&
            entry.lastName.toLowerCase() === name[1].toLowerCase()
          );
          addPerson(
            name[0],
            name[1],
            rows[row][COLUMN_MILES_HIKED],
            rows[row][COLUMN_HOURS],
            index
          );        });
    } // if
  } // for
}; // extractNames()

const addPerson = (firstName, lastName, milesHiked, hours, index) => {
  // add the name if it wasn't found
  if (index === -1) {
    people.push(new Person(
      firstName,
      lastName,
      Number(milesHiked), // make sure this is saved as a number
      Number(hours) // make sure this is saved as a number
    ));
  } else {  // update the name
    people[index].milesHiked += milesHiked;
    people[index].hours += hours;
  } // if
} // addPerson()

const buildTable = () => {
  const peopleTable = document.getElementById('peopleTable');

  const header = `
      <thead>
      <tr>
        <th scope="col">First name</th>
        <th scope="col">Last name</th>
        <th scope="col">Miles hiked</th>
        <th scope="col">Hours</th>
      </tr>
    </thead>`;
  let body = '<tbody>';
  people.forEach((person) => {
    body += `
      <tr>
        <td>${person.firstName}</td>
        <td>${person.lastName}</td>
        <td>${person.milesHiked.toFixed(1)}</td>
        <td>${person.hours.toFixed(1)}</td>
      </tr>`;
  });
  body += '</tbody>';
  peopleTable.innerHTML = header + body;
} // buildTable()