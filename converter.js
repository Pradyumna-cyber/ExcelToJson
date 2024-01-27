function convertExcelToJson() {
    var input = document.getElementById('excelFile');
    var outputDiv = document.getElementById('jsonOutput');

    var file = input.files[0];

    if (file) {
        var reader = new FileReader();

        reader.onload = function (e) {
            var data = e.target.result;
            var workbook = XLSX.read(data, { type: 'binary' });

            var jsonData = {};

            workbook.SheetNames.forEach(function (sheetName) {
                jsonData[sheetName] = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
            });

            // Convert JSON to a formatted string
            var jsonString = JSON.stringify(jsonData, null, 2);


            // Save JSON data to a file and create a download link
            saveToFile(jsonString, 'converted_data.json', 'application/json');
        };

        reader.readAsBinaryString(file);
    } else {
        outputDiv.innerText = 'Please select an Excel file.';
    }
}

function saveToFile(data, filename, type) {
    var blob = new Blob([data], { type: type });

    var a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = filename;

    
    document.body.appendChild(a);

    
    a.click();

    
    document.body.removeChild(a);
}
