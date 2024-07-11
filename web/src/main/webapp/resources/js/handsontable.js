/**
 * Processing events of search OKR by emails.
 */
$(document).ready(function() {
    loadTableData();

    $('#saveDataBtn').on('click', function() {
        saveTableData();
    });
    $('#exportExcelBtn').on('click', function() {
        exportToExcel();
    });
});

/**
 * Load column width, header, initTable()
 */
function loadTableData() {
    $.ajax({
        url : _ctx + 'handsontable/loaddata',
        type : 'GET',
        dataType : 'json',
        contentType : 'application/json',
        success : function(res) {
            console.log("res=" + JSON.stringify(res));
    
            if (res) {
                tblProductData = res;
                tblProductColHeaders = res.colHeaders;
                tblProductColWidths = res.colWidths;
                okrData = res.data;
                initTable();
            }                
        },
        error : function (e) {
            console.log("Error: " + e);
        }
    });
}

function initTable() {
    var container = document.getElementById('tblProduct');
  
    hotProduct = new Handsontable(container, {
        data: tblProductData.data,
        colHeaders: tblProductData.colHeaders,
        colWidths: tblProductData.colWidths,
        height: 290,
        rowHeaders: true,
        minRows: 10,
        currentRowClassName: 'currentRow',
        currentColClassName: 'currentCol',
        manualColumnResize: true,
        manualRowResize: true,
        minSpareRows: 1,
        contextMenu: true,
        licenseKey: 'non-commercial-and-evaluation'
    });
}

function saveTableData() {
    var updatedData = hotProduct.getData();
    $.ajax({
        url: _ctx + 'handsontable/savedata',
        type: 'POST',
        data: JSON.stringify(updatedData),
        dataType: 'json',
        contentType: 'application/json',
        success: function(res) {
            console.log("Data saved successfully");
        },
        error: function(e) {
            console.log("Error: " + e);
        }
    });
}

        function exportToExcel() {
            fetch(_ctx + 'handsontable/export/excel')
                .then(response => response.blob())
                .then(blob => {
                    const url = window.URL.createObjectURL(blob);
                    const a = document.createElement('a');
                    a.href = url;
                    a.download = 'handsontable-data.xlsx';
                    document.body.appendChild(a);
                    a.click();
                    a.remove();
                })
                .catch(error => console.log("Error: " + error));
        }