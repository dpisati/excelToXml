var docFileName = "";
var xmlBody = "";
var fileUpload = document.getElementById("fileUpload");
var h1FileName = document.getElementById("filename");
var ExcelTable = document.getElementById("ExcelTable");
var codeTag = document.getElementById("xmlInHtml");
var htmlTable = document.getElementById("table");
var uploadFile = document.getElementById("fileUpload");
var deleteButton = document.getElementById("remove-btn");
var downloadButton = document.getElementById("download-btn");
var browseButton = document.getElementById("browse-btn");

fileUpload.addEventListener('change', (event) => {
    UploadProcess();
  });

function cleanPage() {   
    while (ExcelTable.firstChild) {
        ExcelTable.removeChild(ExcelTable.firstChild);
    }    
    while (codeTag.firstChild) {
        codeTag.removeChild(codeTag.firstChild);
    }    
    
    h1FileName.innerHTML = ""
    uploadFile.value = "";
    docFileName = "";
    xmlBody = "";

    deleteButton.style.display = "none";
    downloadButton.style.display = "none";
    browseButton.style.display = "inline";
}

function UploadProcess() {
    //Reference the FileUpload element.
    
    //Validate whether File is valid Excel file.
    // var regex = /^([a-zA-Z0-9\s_\\.\-:])+(.xls|.xlsx)$/;
    var isExcelFile = fileUpload.value.includes(".xls");
    
    // if (regex.test(fileUpload.value.toLowerCase())) {
        if (isExcelFile) {
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

function removeSpecialCharacters(string) {
    var newString = string
    newString = newString.replace("&", " and ");
    newString = newString.replace(">", " less than ");
    newString = newString.replace("<", " more than ");
    newString = newString.replace(/["']/g, "");
    return newString;
}

function GetTableFromExcel(data) {
    deleteButton.style.display = "inline";
    downloadButton.style.display = "inline";
    browseButton.style.display = "none";
    
    var fileName = fileUpload.files[0].name;
    fileName = removeSpecialCharacters(fileName);
    var jobName = fileName.replace(/\.[^/.]+$/, "");
    jobName = removeSpecialCharacters(jobName);
    
    h1FileName.innerHTML = jobName;
    docFileName = jobName;

    
    

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
    myTable.setAttribute("id", "table");
    myTable.border = "1";

    const headers = [];
    
    //Add the header row.
    var row = myTable.insertRow(-1);

    var p = excelRows[0];
    for (var key in p) {
        if (p.hasOwnProperty(key)) {
            var columnNumber = key.trim();
            var headerCell = document.createElement("TH");
            headerCell.innerHTML = columnNumber;
            headers.push(columnNumber);
            row.appendChild(headerCell);
        }
    }
    var codeLine = document.createElement("p");
    codeLine.innerHTML += `
    &lt;?xml version="1.0" encoding="UTF-8"?&gt;
    &lt;Root Application="Microvellum" ApplicationVersion="7.0"&gt;
        &lt;Project Name="${jobName}"&gt;
            &lt;SpecificationGroups&gt;
                &lt;SpecificationGroup Name="01-JS Standard"&gt;
                &lt;/SpecificationGroup&gt;
            &lt;/SpecificationGroups&gt;

            &lt;Products&gt;
    `;
    codeTag.appendChild(codeLine);

    //Add the data rows from Excel file.
    for (var i = 1; i < excelRows.length; i++) {

        //Add the data row.
        var rowNumber = excelRows[i];

        var codeLine = document.createElement("p");
        codeLine.setAttribute("id", `productName${i}`);``
        codeTag.appendChild(codeLine);
        var productElement = document.getElementById(`productName${i}`);

        var row = myTable.insertRow(-1);

        for (var key in excelRows[i]) {
            if (rowNumber.hasOwnProperty(key)) {
                var cellValue =  rowNumber[key].trim().replace(/\s/g, ' ');
                // var cellTitle = headers[columnNumber];
                var codeLine = document.createElement("p");
                cellValue = removeSpecialCharacters(cellValue);

                if(headers.indexOf(key) <= 6) {
                    switch(key) {
                        case "Qty":
                            codeLine.innerHTML += `&lt;Quantity&gt;${cellValue}&lt;/Quantity&gt;`
                        break;
                        case "Name":
                            productElement.innerHTML += `&lt;Product Name="${cellValue}"&gt;`;
                        break;
                        case "Width":
                            codeLine.innerHTML += `&lt;Width&gt;${cellValue}&lt;/Width&gt;`
                        break;
                        case "Height":
                            codeLine.innerHTML += `&lt;Height&gt;${cellValue}&lt;/Height&gt;`
                        break;
                        case "Depth":
                            codeLine.innerHTML += `&lt;Depth&gt;${cellValue}&lt;/Depth&gt;`
                        break;
                        case "ProductSpecGroupName":
                            codeLine.innerHTML += `&lt;LinkIDSpecificationGroup&gt;${cellValue}&lt;/LinkIDSpecificationGroup&gt;`
                        break;
                        case "Comments":
                            codeLine.innerHTML += `&lt;Comment&gt;${cellValue}&lt;/Comment&gt;&lt;Prompts&gt;`
                        break;
                    }

                    codeTag.appendChild(codeLine);

                    var contentCell = document.createElement("TH");
                    contentCell.innerHTML = cellValue;
                    row.appendChild(contentCell);
                }

                if(headers.indexOf(key) > 6) {
                    var contentCell = document.createElement("TH");
                    contentCell.innerHTML = cellValue + "  " + key;
                    row.appendChild(contentCell);
                    codeLine.innerHTML += `&lt;Prompt Name="${key}"&gt;&lt;Value&gt;${cellValue}&lt;/Value&gt;&lt;/Prompt&gt;`
                    codeTag.appendChild(codeLine);
                }
            }
        }

        codeLine.innerHTML += `&lt;/Prompts&gt;&lt;/Product&gt;` 
        codeTag.appendChild(codeLine);       
    }
    
    codeLine.innerHTML += `&lt;/Products&gt;&lt;/Project&gt;&lt;/Root&gt;`
    codeTag.appendChild(codeLine);
    
    // ExcelTable.appendChild(myTable);
};

function download(filename, text) {
    var element = document.createElement('a');
    element.style.display = 'none';
    element.setAttribute('href', 'data:text/pain;charset=utf-8' + encodeURIComponent(text))
    element.setAttribute('download', filename);
    document.body.appendChild(element);
    element.click();
    document.body.removeChild(element);
}

deleteButton.addEventListener("click", function() {
    cleanPage();
})

downloadButton.addEventListener("click", function() {
        var allPs = document.getElementsByTagName("p");
        for (var i = 0; i < allPs.length; i++) {
            if(allPs[i].innerText) {
                xmlBody += allPs[i].innerText;
            }
        }

        var finalFileName = docFileName + ".xml";
        var xml = xmlBody;

        var blob = new Blob([xml], {
            // type: "text/plain;charset=uft-8"
        });
        saveAs(blob, finalFileName)
});
