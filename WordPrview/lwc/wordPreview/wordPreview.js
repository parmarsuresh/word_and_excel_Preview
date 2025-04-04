import { LightningElement , track } from 'lwc';
import jszip  from '@salesforce/resourceUrl/jszip';
import docxpreview  from '@salesforce/resourceUrl/docxpreview';
import thumbnail_example  from '@salesforce/resourceUrl/thumbnail_example';
import thumbnail_excss  from '@salesforce/resourceUrl/thumbnail_excss';
import excelFileReader from "@salesforce/resourceUrl/ExcelReaderPlugin";
import { ShowToastEvent } from 'lightning/platformShowToastEvent';
let XLS = {};
import { loadScript,loadStyle } from 'lightning/platformResourceLoader';
export default class WordPreview extends LightningElement {

    strAcceptedFormats = [".xls", ".xlsx"];
    strUploadFileName; //Store the name of the selected file.
    objExcelToJSON; //Javascript object to store the content of the file

    pdfLibInitialized = false;
    docxLoaded = false;
    fileSelected = false;
    
    headers = [];
    rows = [];
    isFileLoaded = false;

    openWordPreview = false;
    openExcelPreview = false;
    downloadimageName = '';
    @track SheetName = [];
    @track sheetsData = {};

    connectedCallback() {
        Promise.all([
            loadScript(this, jszip),
            loadScript(this, docxpreview),
            loadScript(this, thumbnail_example),
            loadScript(this, excelFileReader),
            loadStyle(this, thumbnail_excss)
        ])
        .then(() => {
            console.log('Libraries loaded successfully=====');
            console.log('window.XLSX--;',window.XLSX.read);
            this.docxLoaded = true;
            this.xlsx = window.XLSX; 
            XLS = XLSX;
        })
        .catch(error => console.error('Error loading libraries', error));
    }

    closeWordModal(){
        this.openWordPreview = false;
        this.openExcelPreview = false;
        this.SheetName = [];
        this.sheetsData = {};
        this.downloadimageName = '';
    }

    handleFileChange(event) {

        const file = event.target.files[0];
        this.downloadimageName = file.name;
        if (!file || !this.docxLoaded) {
            return;
        }
        this.openWordPreview = true;
        setTimeout(()=>{
            if(this.downloadimageName.toLowerCase().includes('.docx'))
            {
                this.openExcelPreview = false;
                const container = this.template.querySelector('.document-container');
                // console.log('Container:', container);
                // console.log('file', file);
                // console.log('window.docx.renderAsync:', window.docx.renderAsync);
                if (container) {
                    container.innerHTML = ''; // Clear previous content
                    const docxOptions = {
                        debug: true,
                        experimental: true
                    };

                    window.docx.renderAsync(file, container, null, docxOptions)
                        .then(() => console.log('File rendered successfully'))
                        .catch(error => console.error('Error rendering file:', error));
                }
            }
            else if(this.downloadimageName.toLowerCase().includes('.xls') || this.downloadimageName.toLowerCase().includes('.xlsx'))
            {
                this.openExcelPreview = true;
                this.handleProcessExcelFile(file);
            }

        },1000) 
    }


   
   //  rendersheetTable(data) {
   //      // console.log('data--', JSON.stringify(data));
   //      const table = this.template.querySelector('.excelTable');
   //      table.innerHTML = ''; // Clear previous content
   //      if(data.length > 0)
   //      {
   //          if (data.length === 0) {
   //              const noDataRow = document.createElement('tr');
   //              const noDataCell = document.createElement('td');
   //              noDataCell.colSpan = Object.keys(data[0] || {}).length;
   //              noDataCell.textContent = 'No data available';
   //              noDataCell.style.textAlign = 'center';
   //              noDataCell.style.padding = '8px';
   //              noDataRow.appendChild(noDataCell);
   //              table.appendChild(noDataRow);
   //              return;
   //          }

   //          // Create table header
   //          const headerRow = document.createElement('tr');
   //          headerRow.style.backgroundColor = '#f2f2f2';
   //          headerRow.style.border = '1px solid black';

   //          const headers = Object.keys(data[0]);
   //          headers.forEach(header => {
   //              const th = document.createElement('th');
   //              th.textContent = header.trim(); // Trim to avoid trailing spaces
   //              th.style.border = '1px solid black';
   //              th.style.padding = '8px';
   //              th.style.fontWeight = 'bold';
   //              headerRow.appendChild(th);
   //          });
   //          table.appendChild(headerRow);

   //          // Create table rows
   //          data.forEach(row => {
   //              const tr = document.createElement('tr');
   //              tr.style.border = '1px solid black';

   //              headers.forEach(header => {
   //              const td = document.createElement('td');
   //              td.textContent = row[header] !== undefined ? row[header] : ''; // Handle missing fields
   //              td.style.border = '1px solid black';
   //              td.style.padding = '8px';
   //              tr.appendChild(td);
   //              });

   //              table.appendChild(tr);
   //          });
   //      }
   // }


  
    rendersheetTable(data) {
		const table = this.template.querySelector('.excelTable');
		table.innerHTML = ''; // Clear previous content

		if (!data || data.length === 0) {
			const noDataRow = document.createElement('tr');
			const noDataCell = document.createElement('td');
			noDataCell.colSpan = 1;
			noDataCell.textContent = 'No data available';
			noDataCell.style.textAlign = 'center';
			noDataCell.style.padding = '8px';
			noDataRow.appendChild(noDataCell);
			table.appendChild(noDataRow);
			return;
		}

		// Collect all headers from the data
		let headers = new Set();
		data.forEach(row => Object.keys(row).forEach(key => headers.add(key)));
		headers = Array.from(headers);

		// Create table header
		const headerRow = document.createElement('tr');
		headerRow.style.backgroundColor = '#f2f2f2';
		headerRow.style.border = '1px solid black';

		headers.forEach(header => {
			const th = document.createElement('th');
			th.textContent = header.trim();
			th.style.border = '1px solid black';
			th.style.padding = '8px';
			th.style.fontWeight = 'bold';
			headerRow.appendChild(th);
		});
		table.appendChild(headerRow);

		// Create table rows
		data.forEach(row => {
			const tr = document.createElement('tr');
			tr.style.border = '1px solid black';

			headers.forEach(header => {
				const td = document.createElement('td');

				if (row[header]) {
					let cellData = row[header];
					td.textContent = cellData.value !== undefined ? cellData.value : '';

					if (cellData.bold) {
						td.style.fontWeight = 'bold';
					}

					if (cellData.highlighted) {
						td.style.backgroundColor = '#ffff99'; // Light yellow
					}
				} else {
					td.textContent = '';
				}

				td.style.border = '1px solid black';
				td.style.padding = '8px';
				tr.appendChild(td);
			});

			table.appendChild(tr);
		});
	}


    // handleProcessExcelFile(file) {
    //     let objFileReader = new FileReader();
    //     // console.log('XLS.read=',XLS.read);
    //     objFileReader.onload = (event) => {
    //             let objFiledata = event.target.result;
    //             let objFileWorkbook = XLS.read(objFiledata, { type: "binary"});

    //             objFileWorkbook.SheetNames.forEach((sheetName) => {
    //                 const sheetData = XLS.utils.sheet_to_row_object_array(objFileWorkbook.Sheets[sheetName]);
    //                 this.sheetsData[sheetName] = sheetData; // Store data for each sheet
    //             });
    //             this.SheetName = Object.keys(this.sheetsData);
    //             // console.log('All Sheets Data:', JSON.stringify(this.sheetsData));

    //             // Check if data exists
    //             if (Object.keys(this.sheetsData).length === 0) {
    //                 console.warn('No data found in the Excel file.');
    //                 return;
    //             }
    //             // Example: Render the first sheet (or modify as needed)
    //             const firstSheetName = objFileWorkbook.SheetNames[0];
    //             const firstSheetData = this.sheetsData[firstSheetName];
    //             if (firstSheetData.length > 0) {
    //                 this.rendersheetTable(firstSheetData); // Render the first sheet as a table
    //             } else {
    //                 console.warn(`Sheet "${firstSheetName}" is empty.`);
    //             }
    //         };
    //         objFileReader.onerror = function (error) {
    //             console.log('error==',error);
    //         };
    //         objFileReader.readAsBinaryString(file);
    // }

     handleProcessExcelFile(file) {
        let objFileReader = new FileReader();
        // console.log('XLS.read=',XLS.read);
        objFileReader.onload = (event) => {
        let objFiledata = event.target.result;
        let objFileWorkbook = XLS.read(base64Content, { type: "base64", cellStyles: true });
        objFileWorkbook.SheetNames.forEach((sheetName) => {
				const sheet = objFileWorkbook.Sheets[sheetName];
				const range = XLS.utils.decode_range(sheet['!ref']);
				const formattedData = [];
				let headers = [];

				const allRows = [];

				for (let R = range.s.r; R <= range.e.r; ++R) {
					const rowData = {};
					const isHeaderRow = R === range.s.r;

					for (let C = range.s.c; C <= range.e.c; ++C) {
						const cellAddress = XLS.utils.encode_cell({ r: R, c: C });
						const cell = sheet[cellAddress];

						if (isHeaderRow) {
							headers[C] = cell?.v ? String(cell.v).trim() : `Column ${C + 1}`;
						} else {
							const header = headers[C];
							if (!header) continue;

							if (!cell) {
								rowData[header] = { value: "", bold: false, highlighted: false };
								continue;
							}

							let cellValue = cell.v;
							let isDate = false;

							if (typeof cellValue === 'number' && cell.t === 'n' && cell.z?.includes('d')) {
								cellValue = this.excelDateToJSDate(cellValue);
								isDate = true;
							}

							const isBold = cell.s?.font?.bold === 1;
							const isHighlighted = !!cell.s?.fill?.fgColor?.rgb;

							rowData[header] = {
								value: cellValue,
								bold: isBold,
								highlighted: isHighlighted,
								isDate
							};
						}
					}

					if (!isHeaderRow) {
						const hasData = Object.values(rowData).some(cell => (cell?.value ?? '').toString().trim() !== '');
						allRows.push(hasData ? rowData : null); // Add null if row is empty
					}
				}

				// Remove trailing blank rows
				for (let i = allRows.length - 1; i >= 0; i--) {
					if (allRows[i] === null) {
						allRows.pop();
					} else {
						break;
					}
				}

				this.sheetsData[sheetName] = allRows.filter(row => row !== null);
			});

			this.SheetName = Object.keys(this.sheetsData);

			if (this.SheetName.length > 0) {
				this.rendersheetTable(this.sheetsData[this.SheetName[0]]);
			}   
    }

    switchSheet(event)
    {
        // console.log('event==',JSON.stringify(event.currentTarget.dataset));
        let firstSheetName = event.currentTarget.dataset.name;
        const firstSheetData = this.sheetsData[firstSheetName];
        if (firstSheetData.length > 0) {
            this.rendersheetTable(firstSheetData); // Render the first sheet as a table
        } else {
            console.warn(`Sheet "${firstSheetName}" is empty.`);
        }
    }


}
