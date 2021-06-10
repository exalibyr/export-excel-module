const uriXLS = 'data:application/vnd.ms-excel;base64,';
const templateXLS = '<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40"><head><xml><x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet><x:Name>{worksheet}</x:Name><x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions></x:ExcelWorksheet></x:ExcelWorksheets></x:ExcelWorkbook></xml><meta http-equiv="content-type" content="text/plain; charset=UTF-8"/></head><body>{table}</body></html>';

function exportData() {
    let worksheetName = 'test';
    let content = buildTable(parseForm());
    let filename = 'test.xls';
    let worksheet = {worksheet: worksheetName || 'Worksheet', table: content.outerHTML};
    download(uriXLS + formatData(templateXLS, worksheet), fileName);
}

function download(uri, filename) {
    let link = document.createElement("a");
    link.download = filename;
    link.href = uri;
    link.click();    
}

function formatData (template, worksheet) {
    let formattedData = template.replace(/{(\w+)}/g, function(m, p) {
        return worksheet[p]; 
    });
    return window.btoa(unescape(encodeURIComponent(formattedData)));
}

function parseForm() {
    let data = {};
    let fields = document.querySelectorAll('.exportable-field');
    if (fields != null) {
        for (let field of fields) {
            let label = findElementByAttribute(field, 'for');
            let value = findElementByAttribute(field, 'value');
            data[label.textContent] = value.getAttribute("value");
        }
    }
    return data;
}
function buildTable (data) {
    let table = document.createElement("table");

    for (let key in data) {
        let row = document.createElement("tr");

        let fieldName = document.createElement("td");
        fieldName.innerHTML = key;

        let fieldValue = document.createElement("td");
        fieldValue.innerHTML = data[key];

        row.append(fieldName, fieldValue);

        table.appendChild(row);
    }

    return table;
}
function findElementByAttribute (node, attribute) {
    if (node.hasAttribute(attribute)) {
        return node;
    } else {
        if (node.hasChildNodes()) {
            for (let child of node.children) {
                let element = findElementByAttribute(child, attribute);
                if (element != null) {
                    return element;
                }
            }
            return null;
        } else {
            return null;
        }
    }
}

function getText(label) {
    let obj = new Object(node);
    obj.children = null;
    return obj.textContent;
}

function te() {
    let data = {};
    let blocks = document.querySelectorAll('.exportable-block');
    blocks
    for (let block of blocks) {
        Object.defineProperty(data, block.innerHTML), {
          value: null,
          writable: true,
          enumerable: true,
          configurable: true
        });
        let fields = document.querySelectorAll('.exportable-field');
        if (fields != null) {
            let blockData = data[block.innerHTML];
            for (let field of fields) {
                Object.defineProperty(blockData, field.getAttribute("label"), {
                  value: field.getAttribute("value"),
                  writable: true,
                  enumerable: true,
                  configurable: true
                });
            }
        }
    }
}

//const tableToExcel = (function() {
//        let uri = 'data:application/vnd.ms-excel;base64,'
//        , template = '<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40"><head><xml><x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet><x:Name>{worksheet}</x:Name><x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions></x:ExcelWorksheet></x:ExcelWorksheets></x:ExcelWorkbook></xml><meta http-equiv="content-type" content="text/plain; charset=UTF-8"/></head><body>{table}</body></html>'
//        , base64 = function(s) { return window.btoa(unescape(encodeURIComponent(s))) }
//        , format = function(s, c) { 	    	 
//            return s.replace(/{(\w+)}/g, function(m, p) { return c[p]; }) 
//        }
//        , download = function(uri, filename) {
//            let link = document.createElement("a");
//            link.download = filename;
//            link.href = uri;
//            link.click();
//        }
//        return function(content, worksheetName, fileName) {
//            let worksheet = {worksheet: worksheetName || 'Worksheet', table: content.outerHTML};
//            download(uri + base64(format(template, worksheet)), fileName);
//        }
//    })(); 
//function exportExcel() {
//    tableToExcel(buildTable(parseForm()),'test', 'test.xls');
//}