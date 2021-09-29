function UploadProcess() {
    //Reference the FileUpload element.

    //Validate whether File is valid Excel file.
    var regex = /^([a-zA-Z0-9\s_\\.\-:])+(.xls|.xlsx)$/;
    
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
    var fileUpload = document.getElementById("fileUpload");
    var fileName = fileUpload.files[0].name;
    var jobName = fileName.replace(/\.[^/.]+$/, "");

    var ExcelTable = document.getElementById("ExcelTable");
    var codeTag = document.getElementById("xmlInHtml");

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
            var cellValue =  p[key].trim().replace(/\s/g, '_')
            var columnNumber = key.trim().replace('__EMPTY', '').replace('_', '');
            
            if(columnNumber === "") {
                columnNumber = 0;
            }

            var headerCell = document.createElement("TH");
            headerCell.innerHTML = cellValue + " " + columnNumber;
            headers.push(cellValue);
            row.appendChild(headerCell);
        }
    }

    var codeLine = document.createElement("p");
    codeLine.innerHTML += `
    &lt;?xml version="1.0" encoding="UTF-8"?&gt;<br />
    &lt;Root Application="Microvellum" ApplicationVersion="7.0"&gt;<br />
        &lt;Project Name="${jobName}"&gt;<br />
            &lt;SpecificationGroups&gt;<br />
                &lt;SpecificationGroup Name="01-JS Standard"&gt;<br />
                &lt;/SpecificationGroup&gt;<br />
            &lt;/SpecificationGroups&gt;<br />

            &lt;Products&gt;<br /><br /><br />
    `;
    codeTag.appendChild(codeLine);


    //Add the data rows from Excel file.
    for (var i = 1; i < excelRows.length; i++) {

        //Add the data row.
        var rowNumber = excelRows[i];
        var isTheLineEmpty = (rowNumber.__EMPTY_18 === "" ||  rowNumber.__EMPTY === "");

        if(!isTheLineEmpty) {

            var codeLine = document.createElement("p");
            codeLine.setAttribute("id", `productName${i}`);
            codeTag.appendChild(codeLine);
            var productElement = document.getElementById(`productName${i}`);


            var row = myTable.insertRow(-1);

            for (var key in excelRows[i]) {
                if (rowNumber.hasOwnProperty(key)) {
                    var cellValue =  rowNumber[key].trim().replace(/\s/g, ' ');
                    var columnNumber = key.trim().replace('__EMPTY', '').replace('_', '');
                    if(columnNumber === "") {
                        columnNumber = 0;
                    }
                    var cellTitle = headers[columnNumber];

                    var codeLine = document.createElement("p");

                    if(columnNumber <= 6) {

                        switch(columnNumber) {
                            case 0:
                                codeLine.innerHTML += `
                                    &lt;Quantity&gt;${cellValue}&lt;/Quantity&gt;<br />
                                `
                            break;
                            case "1":
                                productElement.innerHTML += `
                                    &lt;Product Name="${cellValue}"&gt;<br />
                                `;
                            break;
                            case "2":
                                codeLine.innerHTML += `
                                    &lt;Width&gt;${cellValue}&lt;/Width&gt;<br />
                                `
                            break;
                            case "3":
                                codeLine.innerHTML += `
                                    &lt;Height&gt;${cellValue}&lt;/Height&gt;<br />
                                `
                            break;
                            case "4":
                                codeLine.innerHTML += `
                                    &lt;Depth&gt;${cellValue}&lt;/Depth&gt;<br />
                                `
                            break;
                            case "5":
                                codeLine.innerHTML += `
                                    &lt;LinkIDSpecificationGroup&gt;${cellValue}&lt;/LinkIDSpecificationGroup&gt;<br />
                                `
                            break;
                            case "6":
                                codeLine.innerHTML += `
                                    &lt;Comment&gt;${cellValue}&lt;/Comment&gt;<br />
                                    &lt;Prompts&gt;
                                `
                            break;
                        }

                        codeTag.appendChild(codeLine);

                        if(cellValue && cellValue !== "") {
                            if(cellTitle) {
                                var contentCell = document.createElement("TH");
                                contentCell.innerHTML = cellValue + ' ' + columnNumber;
                                row.appendChild(contentCell);
                            }
                        }
                    }

                    if(columnNumber > 6) {
                        if(cellValue && cellValue !== "") {
                            if(cellTitle) {
                                var contentCell = document.createElement("TH");
                                contentCell.innerHTML = cellValue + ' ' + columnNumber;
                                row.appendChild(contentCell);
                                codeTag.innerHTML += `
                                        &lt;Prompt Name="${cellTitle}"&gt;<br />
                                            &lt;Value&gt;${cellValue}&lt;/Value&gt;<br />
                                        &lt;/Prompt&gt;<br />
                                    `
                            }
                        }
                    }
                }
            }
            codeTag.innerHTML += `
                    &lt;/Prompts&gt;<br />
                &lt;/Product&gt;<br /><br /><br /><br /><br />
            `
        }
    }
    
    codeTag.innerHTML += `
            &lt;/Products&gt;<br />
        &lt;/Project&gt;<br />
    &lt;/Root&gt;<br />
    `
    ExcelTable.innerHTML = "";
    ExcelTable.appendChild(myTable);
};