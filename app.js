function toggleReconcileButton() {
        const consentCheckbox = document.getElementById('consent-checkbox');
        const reconcileBtn = document.getElementById('reconcile-btn');
        reconcileBtn.disabled = !consentCheckbox.checked;
    }

    // --- All other JavaScript for the tool's core functionality ---
    let reconciliationData = {};

    function parseExcelFile(file){return new Promise((resolve,reject)=>{const reader=new FileReader;reader.onload=event=>{const data=new Uint8Array(event.target.result),workbook=XLSX.read(data,{type:"array"}),sheetName=workbook.SheetNames[0],worksheet=workbook.Sheets[sheetName],json=XLSX.utils.sheet_to_json(worksheet);resolve(json)},reader.onerror=reject,reader.readAsArrayBuffer(file)})}function normalizeInvoiceNumber(invNum){return"string"!=typeof invNum&&(invNum=String(invNum)),invNum.replace(/[^a-zA-Z0-9]/g,"").toUpperCase()}function parseAndFormatDate(dateInput){if(!dateInput)return"";if("number"==typeof dateInput&&dateInput>1){const jsDate=new Date(864e5*(dateInput-25569));if(jsDate instanceof Date&&!isNaN(jsDate)){let day=String(jsDate.getDate()).padStart(2,"0"),month=String(jsDate.getMonth()+1).padStart(2,"0"),year=jsDate.getFullYear();return`${day}/${month}/${year}`}}if(dateInput instanceof Date&&!isNaN(dateInput)){let day=String(dateInput.getDate()).padStart(2,"0"),month=String(dateInput.getMonth()+1).padStart(2,"0"),year=dateInput.getFullYear();return`${day}/${month}/${year}`}const parsedDate=new Date(dateInput);return parsedDate instanceof Date&&!isNaN(parsedDate)?`${String(parsedDate.getDate()).padStart(2,"0")}/${String(parsedDate.getMonth()+1).padStart(2,"0")}/${parsedDate.getFullYear()}`:String(dateInput)}function convertToDateObject(dateInput){if(!dateInput)return null;if(dateInput instanceof Date&&!isNaN(dateInput))return dateInput;if("number"==typeof dateInput&&dateInput>1)return new Date(864e5*(dateInput-25569));const parsed=new Date(dateInput);return parsed instanceof Date&&!isNaN(parsed)?parsed:dateInput}async function reconcile(){const gstr2bFile=document.getElementById("gstr2b-file").files[0],purchaseFile=document.getElementById("purchase-file").files[0];if(!gstr2bFile||!purchaseFile)return void alert("Please upload both GSTR-2B and Purchase Register files.");try{document.getElementById("results-container").style.display="none";let gstr2bData=await parseExcelFile(gstr2bFile),purchaseData=await parseExcelFile(purchaseFile);gstr2bData=gstr2bData.filter(row=>(row["Supplier GSTIN"]||row["GSTIN of supplier"])&&(row["Invoice number"]||row["Invoice No."])),purchaseData=purchaseData.filter(row=>(row["Supplier GSTIN"]||row.GSTIN)&&(row["Invoice Number"]||row["Invoice No."]));const mapped2B=gstr2bData.map(row=>({gstin:row["Supplier GSTIN"]||row["GSTIN of supplier"],invNum:row["Invoice number"]||row["Invoice No."],invDate:row["Invoice date"]||row["Invoice Date"],invValue:parseFloat(row["Invoice value"])||0,taxableValue:parseFloat(row["Taxable Value"])||0,igst:parseFloat(row["Integrated Tax"])||0,cgst:parseFloat(row["Central Tax"])||0,sgst:parseFloat(row["State/UT Tax"])||0})),mappedPurchase=purchaseData.map(row=>({gstin:row["Supplier GSTIN"]||row.GSTIN,invNum:row["Invoice Number"]||row["Invoice No."],invDate:row["Invoice Date"]||row["Invoice date"],invValue:parseFloat(row["Invoice Value"])||0,taxableValue:parseFloat(row["Taxable Value"])||0,igst:parseFloat(row.IGST)||0,cgst:parseFloat(row.CGST)||0,sgst:parseFloat(row.SGST)||0}));const matchedInvoices=[],mismatchedInvoices=[],onlyInPurchase=[],onlyIn2B=[...mapped2B];mappedPurchase.forEach(pInvoice=>{let foundMatch=!1;for(let i=onlyIn2B.length-1;i>=0;i--){const bInvoice=onlyIn2B[i];if(pInvoice.gstin===bInvoice.gstin&&normalizeInvoiceNumber(pInvoice.invNum)===normalizeInvoiceNumber(bInvoice.invNum)){let reasons=[];Math.abs(pInvoice.invValue-bInvoice.invValue)>1&&reasons.push(`Invoice Value (2B: ${bInvoice.invValue}, Purchase: ${pInvoice.invValue})`),Math.abs(pInvoice.taxableValue-bInvoice.taxableValue)>1&&reasons.push(`Taxable Value (2B: ${bInvoice.taxableValue}, Purchase: ${pInvoice.taxableValue})`),reasons.length>0?mismatchedInvoices.push({gstr2b:bInvoice,purchase:pInvoice,reasons:reasons}):matchedInvoices.push(pInvoice),onlyIn2B.splice(i,1),foundMatch=!0;break}}foundMatch||onlyInPurchase.push(pInvoice)}),reconciliationData={matched:matchedInvoices,mismatched:mismatchedInvoices,purchaseOnly:onlyInPurchase,gstr2bOnly:onlyIn2B},displayResults(matchedInvoices,mismatchedInvoices,onlyInPurchase,onlyIn2B)}catch(error){console.error("Error during reconciliation:",error),alert("An error occurred. Please ensure the Excel files are correctly formatted. Check the browser's developer console (F12) for more details.")}}function displayResults(matched,mismatched,purchaseOnly,gstr2bOnly){document.getElementById("results-container").style.display="block",document.getElementById("matched-count").innerText=matched.length,document.getElementById("mismatched-count").innerText=mismatched.length,document.getElementById("only-purchase-count").innerText=purchaseOnly.length,document.getElementById("only-2b-count").innerText=gstr2bOnly.length,populateSimpleTable("matched-table",matched),populateSimpleTable("only-purchase-table",purchaseOnly),populateSimpleTable("only-2b-table",gstr2bOnly),populateMismatchedTable("mismatched-table",mismatched)}function appendCell(row, value, className) {
    const cell = document.createElement("td");
    if (className) cell.className = className;
    cell.textContent = value == null ? "" : String(value);
    row.appendChild(cell);
    return cell;
}

function populateSimpleTable(tableId, data) {
    const tableBody = document.querySelector(`#${tableId} tbody`);
    tableBody.replaceChildren();

    data.forEach(item => {
        const row = document.createElement("tr");
        [
            item.gstin,
            item.invNum,
            parseAndFormatDate(item.invDate),
            item.invValue || 0,
            item.taxableValue || 0,
            item.igst || 0,
            item.cgst || 0,
            item.sgst || 0
        ].forEach(value => appendCell(row, value));
        tableBody.appendChild(row);
    });
}

function appendDetailLine(cell, label, value) {
    cell.appendChild(document.createTextNode(`${label}: ${value == null ? "" : String(value)}`));
    cell.appendChild(document.createElement("br"));
}

function populateMismatchedTable(tableId, data) {
    const tableBody = document.querySelector(`#${tableId} tbody`);
    tableBody.replaceChildren();

    data.forEach(item => {
        const row = document.createElement("tr");
        appendCell(row, item.purchase.gstin);
        appendCell(row, item.purchase.invNum);
        appendCell(row, item.reasons.join("\n"), "mismatch-reason");

        const gstr2bCell = document.createElement("td");
        appendDetailLine(gstr2bCell, "Inv Dt", parseAndFormatDate(item.gstr2b.invDate));
        appendDetailLine(gstr2bCell, "Inv Val", item.gstr2b.invValue);
        appendDetailLine(gstr2bCell, "Taxable Val", item.gstr2b.taxableValue);
        row.appendChild(gstr2bCell);

        const purchaseCell = document.createElement("td");
        appendDetailLine(purchaseCell, "Inv Dt", parseAndFormatDate(item.purchase.invDate));
        appendDetailLine(purchaseCell, "Inv Val", item.purchase.invValue);
        appendDetailLine(purchaseCell, "Taxable Val", item.purchase.taxableValue);
        row.appendChild(purchaseCell);
        tableBody.appendChild(row);
    });
}
function exportAllToExcel(){if(!reconciliationData||0===Object.keys(reconciliationData).length)return void alert("Please run the reconciliation first before exporting.");const wb=XLSX.utils.book_new(),matchedSheetData=reconciliationData.matched.map(item=>({"Supplier GSTIN":item.gstin,"Invoice No.":item.invNum,"Invoice Date":convertToDateObject(item.invDate),"Invoice Value":item.invValue,"Taxable Value":item.taxableValue,IGST:item.igst,CGST:item.cgst,SGST:item.sgst})),matchedWs=XLSX.utils.json_to_sheet(matchedSheetData);XLSX.utils.book_append_sheet(wb,matchedWs,"Matched Invoices");const mismatchedSheetData=reconciliationData.mismatched.map(item=>({"Supplier GSTIN":item.purchase.gstin,"Invoice No.":item.purchase.invNum,"Mismatch Reason":item.reasons.join(", "),"GSTR-2B Inv Date":convertToDateObject(item.gstr2b.invDate),"Purchase Inv Date":convertToDateObject(item.purchase.invDate),"GSTR-2B Inv Value":item.gstr2b.invValue,"Purchase Inv Value":item.purchase.invValue,"GSTR-2B Taxable Value":item.gstr2b.taxableValue,"Purchase Taxable Value":item.purchase.taxableValue})),mismatchedWs=XLSX.utils.json_to_sheet(mismatchedSheetData);XLSX.utils.book_append_sheet(wb,mismatchedWs,"Mismatched Invoices");const purchaseOnlySheetData=reconciliationData.purchaseOnly.map(item=>({"Supplier GSTIN":item.gstin,"Invoice No.":item.invNum,"Invoice Date":convertToDateObject(item.invDate),"Invoice Value":item.invValue,"Taxable Value":item.taxableValue,IGST:item.igst,CGST:item.cgst,SGST:item.sgst})),purchaseOnlyWs=XLSX.utils.json_to_sheet(purchaseOnlySheetData);XLSX.utils.book_append_sheet(wb,purchaseOnlyWs,"In Purchase Register Only");const gstr2bOnlySheetData=reconciliationData.gstr2bOnly.map(item=>({"Supplier GSTIN":item.gstin,"Invoice No.":item.invNum,"Invoice Date":convertToDateObject(item.invDate),"Invoice Value":item.invValue,"Taxable Value":item.taxableValue,IGST:item.igst,CGST:item.cgst,SGST:item.sgst})),gstr2bOnlyWs=XLSX.utils.json_to_sheet(gstr2bOnlySheetData);XLSX.utils.book_append_sheet(wb,gstr2bOnlyWs,"In GSTR-2B Only"),XLSX.writeFile(wb,"GST_Reconciliation_Report.xlsx")}
