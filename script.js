//current array of items in the services table
var currentList = [];

//current total price for invoice (sum of "amount(s)" in the services table)
var runningTotal = 0;

//number to currency formatter
var formatter = new Intl.NumberFormat('en-US', {
    style: 'currency',
    currency: 'USD',
});

var selectedRow = null;

//updates the current amount of money in the services form. time * hours
function updateAmount() {
    var hoursValue = document.getElementById("hours").value;
    var rateValue = document.getElementById("rate").value;
    var hoursXrate = hoursValue * rateValue;
    document.getElementById("curamount").innerHTML = formatter.format(hoursXrate);
}

//gathers form data, adds form data to the running list of items and to the main table then calls resetform 
function onFormSubmit() {
    var formData = readFormData();
    insertNewRecord(formData);
    resetForm();
}

//helper function to gather form field data on submit of the services form
function readFormData() {
    var formData = {};
    formData["item"] = document.getElementById("item").value;
    formData["hours"] = document.getElementById("hours").value;
    formData["rate"] = document.getElementById("rate").value;
    return formData;
}

//helper function to insert data into running list of items and the main table
function insertNewRecord(data) {
    var task = document.getElementById("item").value;
    var hoursValue = document.getElementById("hours").value;
    var rateValue = document.getElementById("rate").value;
    var hoursXrate = hoursValue * rateValue;

    currentList.push({ item: task, rate: rateValue, hours: hoursValue, amount: hoursXrate });

    var table = document.getElementById("itemList").getElementsByTagName("tbody")[0];
    var newRow = table.insertRow(table.length);

    cell1 = newRow.insertCell(0);
    cell2 = newRow.insertCell(1);
    cell3 = newRow.insertCell(2);
    cell4 = newRow.insertCell(3);
    cell5 = newRow.insertCell(4);

    cell1.innerHTML = data.item;
    if (hoursValue === "") {
        cell2.innerHTML = '0 Hours';
    } else {
        cell2.innerHTML = data.hours + ' Hours';
    }
    cell3.innerHTML = formatter.format(data.rate);
    cell4.innerHTML = formatter.format(hoursXrate);
    cell5.innerHTML = '<a onclick="onDelete(this)" class="btn btn-danger btn-sm"><i style="padding-left:13px;padding-right:12px;" class="fa fa-trash-alt"></i></a>';


    //excel table
    var excelTable = document.getElementById("excelTable").getElementsByTagName("tbody")[0];
    var excelnewRow = excelTable.insertRow(excelTable.length);

    cella = excelnewRow.insertCell(0);
    cellb = excelnewRow.insertCell(1);
    cellc = excelnewRow.insertCell(2);
    celld = excelnewRow.insertCell(3);

    cella.innerHTML = data.item;
    cellb.innerHTML = data.hours + ' Hours';
    cellc.innerHTML = formatter.format(data.rate);
    celld.innerHTML = formatter.format(hoursXrate);

    runningTotal = runningTotal + hoursXrate;
}

//resets serivce, hours, rate and amount form fields.
function resetForm() {
    document.getElementById("item").value = "";
    document.getElementById("hours").value = "";
    document.getElementById("rate").value = "";
    document.getElementById("curamount").innerHTML = "$ 0.00";
    document.getElementById("Total").innerHTML = formatter.format(runningTotal);
    selectedRow = null;
    document.getElementById("item").focus();
}

//deletes selected row, subtracts amount from total price, updates main/excel tables
function onDelete(td) {
    row = td.parentElement.parentElement;
    var getRow = document.getElementById("itemList").rows[row.rowIndex].cells[0].innerHTML;
    var getPrice = document.getElementById("itemList").rows[row.rowIndex].cells[3].innerHTML;
    var validatePrice = getPrice.replace(/[|&;$%@"<>()+,]/g, "");
    runningTotal = runningTotal - validatePrice;
    document.getElementById("Total").innerHTML = runningTotal;

    for (var i = 0; i < currentList.length; i++) {
        if (currentList[i].item === getRow) {

            currentList.splice(i, 1);
        }
    }

    grabRow = row.rowIndex;
    console.log(row.rowIndex);
    console.log(grabRow);
    console.log('split');
    document.getElementById("itemList").deleteRow(row.rowIndex);
    console.log(row.rowIndex);
    console.log(grabRow);
    document.getElementById("excelTable").deleteRow(grabRow + 8);
    resetForm();
}

//helper function for preview and export to excel/word methods
//gathers top information form field data, gets excel/preview form data
//generates dynamic list of items from main table   
function populate() {
    //reset generated list in preview to nothing so no duplicates occur
    document.getElementById("wrapper").innerHTML = "";

    var title = document.getElementById("title").value;
    var invoiceNum = document.getElementById("invoiceNum").value;
    var sender = document.getElementById("sender").value;
    var billTo = document.getElementById("billTo").value;
    var curDate = document.getElementById("curDate").value;
    var dueDate = document.getElementById("dueDate").value;
    var company = document.getElementById("company").value;
    var notes = document.getElementById("notes").value;
    var currency = document.getElementById("currency").value;

    var topList = [title, invoiceNum, sender, billTo, curDate, dueDate, company, notes];

    document.getElementById("previewInvoiceTitle").innerHTML = topList[0];
    document.getElementById("previewInvoiceNumber").innerHTML = 'Invoice #' + topList[1];
    document.getElementById("previewInvoiceSender").innerHTML = 'Sender: ' + topList[2];
    document.getElementById("previewInvoiceBillTo").innerHTML = 'Bill To: ' + topList[3];
    document.getElementById("previewInvoiceDate").innerHTML = 'Date: ' + topList[4];
    document.getElementById("previewInvoiceDueDate").innerHTML = 'Due Date: ' + topList[5];
    document.getElementById("previewInvoiceCompany").innerHTML = 'Company:  ' + topList[6];
    document.getElementById("previewInvoiceNotes").innerHTML = 'Note(s): ' + topList[7];

    if (topList[0] === "") {
        document.getElementById("excelInvoiceTitle").innerHTML = "Invoice";
    } else {
        document.getElementById("excelInvoiceTitle").innerHTML = topList[0];
    }
    document.getElementById("excelInvoiceNumber").innerHTML = 'Invoice #' + topList[1];
    document.getElementById("excelInvoiceSender").innerHTML = 'Sender: ' + topList[2];
    document.getElementById("excelInvoiceBillTo").innerHTML = 'Bill To: ' + topList[3];
    document.getElementById("excelInvoiceDate").innerHTML = topList[4].toString();
    document.getElementById("excelInvoiceDueDate").innerHTML = 'Due Date: ' + topList[5].toString();
    document.getElementById("excelInvoiceCompany").innerHTML = 'Company:  ' + topList[6];
    document.getElementById("excelInvoiceNotes").innerHTML = 'Note(s): ' + topList[7];

    //check to see whether the currency is USD or CAD
    if (currency === 'USD') {
        document.getElementById("previewTotal").innerHTML = 'Total: ' + formatter.format(runningTotal) + ' USD';
        document.getElementById("excelTotal").innerHTML = 'Total: ' + formatter.format(runningTotal) + ' USD';
    }

    if (currency === 'CAD') {
        document.getElementById("previewTotal").innerHTML = 'Total: ' + formatter.format(runningTotal) + ' CAD';
        document.getElementById("excelTotal").innerHTML = 'Total: ' + formatter.format(runningTotal) + ' CAD';
    }


    //dynamically generate li items in "wrapper" div 
    currentList.forEach(renderProductList);

    function renderProductList(element, index, arr) {
        var tr = document.createElement('tr');
        document.getElementById('wrapper').appendChild(tr);
        var td = document.createElement('td');
        tr.appendChild(td);
        td.innerHTML = `${currentList[index].item}`;

        var td = document.createElement('td');
        tr.appendChild(td);
        td.innerHTML = `${currentList[index].hours} Hours`;

        var td = document.createElement('td');
        tr.appendChild(td);
        td.innerHTML = `${formatter.format(currentList[index].rate)}`;

        var td = document.createElement('td');
        tr.appendChild(td);
        td.innerHTML = `${formatter.format(currentList[index].amount)}`;

    }

}

//exports preview element to docx format
function ExportToDoc(filename = document.getElementById("title").value) {
    populate();
    var HtmlHead = "<html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'><head><meta charset='utf-8'><title>Export HTML To Doc</title></head><body>";
    var EndHtml = "</body></html>";
    var html = HtmlHead + document.getElementById("preview").innerHTML + EndHtml;
    var blob = new Blob(['\ufeff', html], {
        type: 'application/msword'
    });

    var url = 'data:application/vnd.ms-word;charset=utf-8,' + encodeURIComponent(html);
    filename = filename ? filename + '.doc' : 'DOCXdocument.doc';
    var downloadLink = document.createElement("a");

    document.body.appendChild(downloadLink);

    if (navigator.msSaveOrOpenBlob) {
        navigator.msSaveOrOpenBlob(blob, filename);
    } else {
        downloadLink.href = url;
        downloadLink.download = filename;
        downloadLink.click();
    }

    document.body.removeChild(downloadLink);
}

//exports hidden excel table to .xlsx format
function ExportToExcel(filename = document.getElementById("title").value) {
    populate();
    var wb = XLSX.utils.table_to_book(document.getElementById('excelTable'), { sheet: "Sheet JS" });
    var wbout = XLSX.write(wb, { bookType: 'xlsx', bookSST: true, type: 'binary' });

    filename = filename ? filename + '.xlsx' : 'excelDocument.xlsx';
    saveAs(new Blob([s2ab(wbout)], { type: "application/octet-stream" }), filename);

    function s2ab(s) {
        var buf = new ArrayBuffer(s.length);
        var view = new Uint8Array(buf);
        for (var i = 0; i < s.length; i++) view[i] = s.charCodeAt(i) & 0xFF;
        return buf;
    }

}

//calls populate to get all form field elements on the page and displays the result in a modal
function preview() {
    populate();
    //show the modal after html preview element population has finished
    $('#exampleModal').modal('show');
}
