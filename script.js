//Defining the initial state
var docFileName = "";
var xmlBody = "";

//Getting the elements from HTML
var fileUpload = document.getElementById("fileUpload");
var h1FileName = document.getElementById("filename");
var codeTag = document.getElementById("xmlInHtml");
var htmlTable = document.getElementById("table");
var deleteButton = document.getElementById("remove-btn");
var downloadButton = document.getElementById("download-btn");
var browseButton = document.getElementById("browse-btn");
var mainSection = document.getElementById("main");

//Listen to on load file
fileUpload.addEventListener('change', (event) => {
    UploadProcess();
});

//Function to clear and reset the page
deleteButton.addEventListener("click", function() {
    cleanPage();
})

//Download funtion
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

//Clean and reload the start state
function cleanPage() {   
    while (codeTag.firstChild) {
        codeTag.removeChild(codeTag.firstChild);
    }    
    
    h1FileName.innerHTML = ""
    fileUpload.value = "";
    docFileName = "";
    xmlBody = "";

    deleteButton.style.display = "none";
    downloadButton.style.display = "none";
    mainSection.style.display = "none";
    browseButton.style.display = "inline";
}

//Upload button function
function UploadProcess() {
    //Validate is the file is xls
    var isExcelFile = fileUpload.value.includes(".xls");
    
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

//Remove characters not acceptable in XML
function removeSpecialCharacters(string) {
    var newString = string
    newString = newString.replace("&", " and ");
    newString = newString.replace(">", " less than ");
    newString = newString.replace("<", " more than ");
    newString = newString.replace(/["']/g, "");
    return newString;
}

//Read the data from updated Excel file
function GetTableFromExcel(data) {
    const headers = [];
    const rooms = [];

    deleteButton.style.display = "inline";
    downloadButton.style.display = "inline";
    mainSection.style.display = "block";
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

    //Add all Prompt names to headers array
    var p = excelRows[0];
    for (var key in p) {
        if (p.hasOwnProperty(key)) {
            var columnNumber = key.trim();
            headers.push(columnNumber);
        }
    }
            
    //Adding XML header to file
    var codeLine = document.createElement("p");
    codeLine.innerHTML += `
    &lt;?xml version="1.0" encoding="UTF-8"?&gt;
    &lt;Root Application="Microvellum" ApplicationVersion="7.0"&gt;
        &lt;Project Name="${jobName}"&gt;
            &lt;SpecificationGroups&gt;
                &lt;SpecificationGroup Name="01-JS Standard"&gt;
                &lt;/SpecificationGroup&gt;
            &lt;/SpecificationGroups&gt;
    `;
    codeTag.appendChild(codeLine);

    // Loop to get rooms
    for (var i = 1; i < excelRows.length; i++) {
        //Add the data row.
        var rowNumber = excelRows[i];

        //Loop to get all Rooms
        for (var key in excelRows[i]) {
            if (rowNumber.hasOwnProperty(key)) {
                var cellValue =  rowNumber[key].trim().replace(/\s/g, ' ');
                cellValue = removeSpecialCharacters(cellValue);

                if(key === "Room") {
                    var isRoomInsideTheArray = (rooms.indexOf(cellValue) > -1);
                    if(!isRoomInsideTheArray) {
                        rooms.push(cellValue);
                    }
                }
            }
        }
    }

    //Add Rooms as Locations on top of the file
    
    codeLine.innerHTML += `&lt;Locations&gt;`;
    rooms.forEach(room => {
        codeLine.innerHTML += `&lt;Location Name="${room}"&gt;&lt;/Location&gt;`;
    });
    codeLine.innerHTML += `&lt;/Locations&gt;`;

    //Start The projects tag
    codeLine.innerHTML += `&lt;Products&gt;`;

    //Loop for read cabinet row
    for (var i = 1; i < excelRows.length; i++) {

        //Add the data row.
        var rowNumber = excelRows[i];

        var codeLine = document.createElement("p");
        codeLine.setAttribute("id", `productName${i}`);``
        codeTag.appendChild(codeLine);
        var productElement = document.getElementById(`productName${i}`);

        //Loop to get the data outside the Prompts tag (Width / Height / Depth... etc)
        for (var key in excelRows[i]) {
            if (rowNumber.hasOwnProperty(key)) {
                var cellValue =  rowNumber[key].trim().replace(/\s/g, ' ');
                var codeLine = document.createElement("p");
                cellValue = removeSpecialCharacters(cellValue);

                if(headers.indexOf(key) <= 7) {
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
                            codeLine.innerHTML += `&lt;Comment&gt;${cellValue}&lt;/Comment&gt;`
                        break;
                        case "Room":
                            codeLine.innerHTML += `&lt;LinkIDLocation&gt;${cellValue}&lt;/LinkIDLocation&gt;&lt;Prompts&gt;`
                        break;
                    }
                    codeTag.appendChild(codeLine);
                }

                //Loop for each prompt on cabinet
                if(headers.indexOf(key) > 7) {
                    codeLine.innerHTML += `&lt;Prompt Name="${key}"&gt;&lt;Value&gt;${cellValue}&lt;/Value&gt;&lt;/Prompt&gt;`
                    codeTag.appendChild(codeLine);
                }
            }
        }

        codeLine.innerHTML += `&lt;/Prompts&gt;&lt;/Product&gt;` 
        codeTag.appendChild(codeLine);       
    }
    
    //Closing the file with tags required (Products, Project and Root)
    codeLine.innerHTML += `&lt;/Products&gt;&lt;/Project&gt;&lt;/Root&gt;`
    codeTag.appendChild(codeLine);
};
