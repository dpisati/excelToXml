function UploadProcess() {
    //Reference the FileUpload element.
    var fileUpload = document.getElementById("fileUpload");

    //Validate whether File is valid Excel file.
    var regex = /^([a-zA-Z0-9\s_\\.\-:])+(.xls|.xlsx)$/;
    console.log("regex", regex.test(fileUpload.value.toLowerCase()));

    if (regex.test(fileUpload.value.toLowerCase())) {
        if (typeof (FileReader) != "undefined") {
            var reader = new FileReader();

            //For Browsers other than IE.
            if (reader.readAsBinaryString) {
                reader.onload = function (e) {
                    GetTableFromExcel(e.target.result);
                };
                reader.readAsBinaryString(fileUpload.files[0]);
            } else {
                //For IE Browser.
                reader.onload = function (e) {
                    var data = "";
                    var bytes = new Uint8Array(e.target.result);
                    for (var i = 0; i < bytes.byteLength; i++) {
                        data += String.fromCharCode(bytes[i]);
                    }
                    GetTableFromExcel(data);
                };
                reader.readAsArrayBuffer(fileUpload.files[0]);
            }
        } else {
            alert("This browser does not support HTML5.");
        }
    } else {
        alert("Please upload a valid Excel file.");
    }
};
function GetTableFromExcel(data) {
    //Read the Excel File data in binary
    var workbook = XLSX.read(data, {
        type: 'binary'
    });

    //get the name of First Sheet.
    var Sheet = workbook.SheetNames[0];
    

    //Read all rows from First Sheet into an JSON array.
    var excelRows = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[Sheet]);
    
    //Create a HTML Table element.
    var myTable  = document.createElement("table");
    myTable.border = "1";

    const headers = [];

    //Add the header row.
    var row = myTable.insertRow(-1);
    var p = excelRows[0];

    for (var key in p) {
        if (p.hasOwnProperty(key)) {
            var keyName =  p[key].trim().replace(/\s/g, '_')
            var headerCell = document.createElement("TH");
            headerCell.innerHTML = keyName;
            headers.push(keyName);
            row.appendChild(headerCell);
        }
    }

    console.log("headers: ", headers);

    //Add the data rows from Excel file.
    for (var i = 1; i < excelRows.length; i++) {
        //Add the data row.
        var row = myTable.insertRow(-1);
        var test = excelRows[i];

 

        for (var key in excelRows[i]) {
            if (test.hasOwnProperty(key)) {
                var keyName =  test[key].trim().replace(/\s/g, '_');
                var columnNumber = key.trim().replace('__EMPTY_', '');
                var cellTitle = headers[columnNumber];

                if(keyName && keyName !== "") {
                    if(cellTitle) {
                        var contentCell = document.createElement("TH");
                        contentCell.innerHTML = keyName + ' ' + key.trim().replace('__EMPTY_', '');
                        console.log("========================")
                        console.log("Column Value: ", keyName)
                        console.log("Column Number: ", cellTitle, key)
                        console.log("========================")
                        row.appendChild(contentCell);
                    }
                }
            }
        }
    }

    //Add the data rows from Excel file.
    for (var i = 1; i < excelRows.length; i++) {
        //Add the data row.
        // var row = myTable.insertRow(-1);
        // var test = excelRows[i]
        // for (var key in excelRows[i]) {
        //     if (test.hasOwnProperty(key)) {
        //         var keyName =  test[key].trim().replace(/\s/g, '_')
        //         var headerCell = document.createElement("TH");
        //         headerCell.innerHTML = keyName;
        //         if(keyName !== "") {
        //             row.appendChild(headerCell);
        //             console.log("AAAAAA: ", headerCell)
        //         }
        //     }
        // }

        //Add the data cells.
        // var cell = row.insertCell(-1);
        // cell.innerHTML = excelRows[i].Id;
        // console.log("AAAAA: ", excelRows[i])

        // cell = row.insertCell(-1);
        // cell.innerHTML = excelRows[i].Name;

        // cell = row.insertCell(-1);
        // cell.innerHTML = excelRows[i].Country;
        
        // cell = row.insertCell(-1);
        // cell.innerHTML = excelRows[i].Age;
        
        // cell = row.insertCell(-1);
        // cell.innerHTML = excelRows[i].Date;
        
        // cell = row.insertCell(-1);
        // cell.innerHTML = excelRows[i].Gender;
    }
    

    var ExcelTable = document.getElementById("ExcelTable");
    ExcelTable.innerHTML = "";
    ExcelTable.appendChild(myTable);
};