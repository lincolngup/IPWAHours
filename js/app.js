const inputGroupFile = document.getElementById('inputGroupFile');

const importFile = (event) => {
  var file = event.target.files[0];

  if (file) {
    let fileReader = new FileReader();
    fileReader.onload = progressEvent => {
      let contents = processExcel(progressEvent.target.result);
      console.log(contents);
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
  let result = {};

  workbook.SheetNames.forEach( (sheetName) => {
    let rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {
      header: 1
    });
    if (rows.length) {
      result[sheetName] = rows;
    } // if
  });
  return JSON.stringify(result, 2, 2);
}; // convertToJson()