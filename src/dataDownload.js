var viz, sheet, table
var filename, workbook
var data = [], sheetNames = [], counter=0;

function initViz() {
    var containerDiv = document.getElementById("vizContainer"),
    url = "http://public.tableau.com/views/RegionalSampleWorkbook/Storms",
    options = {
        hideTabs: true,
        hideToolbar: true,
    };
    viz = new tableau.Viz(containerDiv, url, options);
    createExcelWorkbook();
}

function getColumnHeaders(tableColumnHeaders){
    var columnHeaderArray = [];   
    for (i = 0; i < tableColumnHeaders.length; i++)
    {
        columnHeader = tableColumnHeaders [i];                        
        value = columnHeader.getFieldName();
        columnHeaderArray.push(value);                        
    }
    data.push (columnHeaderArray);
} 

function getData(tableData){
    for (i = 0; i < tableData.length; i++)
    {
        row = tableData [i];
        rowArray = [];
        for (j = 0; j < row.length; j++)
        {
            value = row [j].formattedValue;
            rowArray.push(value);
        }
        data.push (rowArray);
    }
}

function createExcelWorkbook () {
    workbook = XLSX.utils.book_new();
}

function writeToExcelWorkbook(ws_name, data){
    /* original data */       
    var ws = XLSX.utils.aoa_to_sheet(data); 
    /* add worksheet to workbook */
    /* catches all the worksheets that have names greater than 31 characters
    try {
        XLSX.utils.book_append_sheet(workbook, ws, ws_name);
    }
    catch (error) {
        console.log (ws_name + " - " + error)
    }                            
}    

function downloadExcelWorkbook (){
    /* write workbook */
    XLSX.writeFile(workbook, filename);
}   

function getSummaryData(){
    filename = viz.getWorkbook().getActiveSheet().getName() + ".xlsx";
    var dashboard = viz.getWorkbook().getActiveSheet().getWorksheets();
    options = {
            ignoreSelection: true
        };
    
    for (k = 0; k < dashboard.length; k++) {
        sheet = dashboard[k];
        sheetNames.push (sheet.getName());
    }

    for (k = 0; k < sheetNames.length; k++) {
        // Getting the Summary Data
        sheet = dashboard[k];
        sheet.getSummaryDataAsync(options).then(function(t){
            table = t;
            data = [];

            // Retrieving data in matrix format
            var tableData = table.getData(); // Get data
            var tableColumnHeaders = table.getColumns(); // Get column headers 
            getColumnHeaders(tableColumnHeaders);
            getData(tableData);

            // Writing to Excel as a new sheet, and printing if all sheets are done
            writeToExcelWorkbook (sheetNames[counter], data);
            counter = counter+1;

            if (counter === dashboard.length) {
                downloadExcelWorkbook(); 
                counter = 0;  
                sheetNames = [];
                createExcelWorkbook();
            }
           
        });
    }
}       
            
