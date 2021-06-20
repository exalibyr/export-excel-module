const uriXLS = 'data:application/vnd.ms-excel;base64,';
const templateXLS = '<html xmlns:o="urn:schemas-microsoft-com:office:office"' +
    ' xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40"><head>' +
    '<xml><x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet><x:Name>{worksheet}</x:Name><x:WorksheetOptions>' +
    '<x:DisplayGridlines/></x:WorksheetOptions></x:ExcelWorksheet></x:ExcelWorksheets></x:ExcelWorkbook></xml>' +
    '<meta http-equiv="content-type" content="text/plain; charset=UTF-8"/></head><body>{table}</body></html>';



const exportData = (sourceComponent, fileName, worksheetName, lang) => {
    let content = buildTable(
        parseForm(sourceComponent, lang),
        lang
    );
    let worksheet = {
        worksheet: worksheetName,
        table: content.outerHTML
    };

    let link = document.createElement("a");
    link.download = fileName + ".xls";
    link.href = uriXLS + formatData(templateXLS, worksheet);
    link.click();
};

const parseForm = (cmp, lang) => {
    let data = {};
    let components = cmp.querySelectorAll(".exportable-cmp");
    if (components != null) {
        for (let component of components) {
            Object.assign(
                data,
                parseComponent(component, lang)
            );
        }
    }
    return data;
};

const parseComponent = (cmp, lang) => {
    let data = {};
    let fields = cmp.querySelectorAll(".exportable-field");
    let header = cmp.querySelector(".exportable-header");
    if (fields != null) {
        for (let field of fields) {
            let label = '';
            for (let item of field.labels) {
                label += item.textContent;
            }
            Object.defineProperty(data, label, {
                enumerable: true,
                configurable: true,
                writable: true,
                value: field.classList.contains("exportable-boolean")
                    ? parseBoolean(field.checked, lang)
                    : field.value
            });
        }
    }
    if (header != null) {
        let result = {};
        Object.defineProperty(result, header.textContent, {
            enumerable: true,
            configurable: true,
            writable: true,
            value: data
        });
        return result;
    }
    return data;
};

const buildTable = (data, lang) => {
    let table = document.createElement("table");
    let mainHeaders = document.createElement("tr");

    let name = document.createElement("td");
    name.setAttribute("bgcolor", "silver");
    name.setAttribute("align", "center");
    name.setAttribute("style", "border: medium solid black; font-weight: bold;");

    let value = document.createElement("td");
    value.setAttribute("bgcolor", "silver");
    value.setAttribute("align", "center");
    value.setAttribute("style", "border: medium solid black; font-weight: bold;");

    switch (lang) {
        case "ru": {
            name.textContent = "Имя поля";
            value.textContent = "Значение поля";
            break;
        }
        case "en": {
            name.textContent = "Field name";
            value.textContent = "Field value";
            break;
        }
        default: {
            name.textContent = "Field name";
            value.textContent = "Field value";
            break;
        }
    }

    mainHeaders.appendChild(name);
    mainHeaders.appendChild(value);
    table.appendChild(mainHeaders);

    return buildRows(table, data);
};

const buildRows = (table, data) => {
    for (let field in data) {
        if (data.hasOwnProperty(field)) {
            if (isObject(data[field])) {
                let row = document.createElement("tr");
                let header = document.createElement("td");

                header.textContent = field;
                header.setAttribute("colspan", '2');
                header.setAttribute("bgcolor", "yellow");
                header.setAttribute("align", "center");
                header.setAttribute("style", "border: medium solid black;");

                row.appendChild(header);
                table.appendChild(row);

                buildRows(table, data[field]);
            } else {
                let row = document.createElement("tr");

                let fieldName = document.createElement("td");
                fieldName.textContent = field;

                let fieldValue = document.createElement("td");
                fieldValue.textContent = data[field];
                fieldValue.setAttribute("align", "left");

                row.appendChild(fieldName);
                row.appendChild(fieldValue);

                table.appendChild(row);
            }
        }
    }

    return table;
};

const formatData = (template, worksheet) => {
    let formattedData = template.replace(/{(\w+)}/g, function (m, p) {
        return worksheet[p];
    });
    return btoa(
        unescape(
            encodeURIComponent(formattedData)
        )
    );
};

const parseBoolean = (value, lang) => {
    if (Object.prototype.toString.call(value) === "[object Boolean]") {
        switch (lang) {
            case "ru": {
                return value ? "Да" : "Нет";
            }
            case "en": {
                return value ? "Yes" : "No";
            }
            default:
                return value ? "Yes" : "No";
        }
    } else {
        return value;
    }
};

const isObject = (v) => {
    return Object.prototype.toString.call(v) === "[object Object]";
};