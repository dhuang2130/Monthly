function generateReport(category) {
    console.log('generateReport function called with category:', category);
    const fileInput = document.getElementById('fileInput');

    if (fileInput.files.length === 0) {
        alert('Please select the required folders containing Excel files.');
        return;
    }

    const files = Array.from(fileInput.files).filter(file => file.name.endsWith('.xls') || file.name.endsWith('.xlsx'));
    let allReportData = [];

    const processFile = (file) => {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = (e) => {
                console.log(`FileReader onload called for file: ${file.name}`);
                const data = new Uint8Array(e.target.result);
                let workbook;

                try {
                    if (file.name.endsWith('.xls')) {
                        workbook = XLSX.read(e.target.result, { type: 'binary' });
                    } else {
                        workbook = XLSX.read(data, { type: 'array' });
                    }
                } catch (error) {
                    console.error('Error reading workbook:', error);
                    reject(error);
                    return;
                }

                console.log('Workbook loaded');
                console.log('Sheet Names:', workbook.SheetNames);

                const sheet = workbook.Sheets[workbook.SheetNames[0]];
                const sheetContents = XLSX.utils.sheet_to_json(sheet, { header: 1 });
                console.log('Sheet Contents:', sheetContents);

                const reportData = generateManufacturedReport(workbook);
                resolve(reportData);
            };

            reader.onerror = (error) => {
                console.error('File reading error:', error);
                reject(error);
            };

            if (file.name.endsWith('.xls')) {
                reader.readAsBinaryString(file);
            } else {
                reader.readAsArrayBuffer(file);
            }
        });
    };

    const processAllFiles = async () => {
        for (const file of files) {
            try {
                const reportData = await processFile(file);
                if (allReportData.length === 0) {
                    allReportData = reportData;
                } else {
                    for (let i = 1; i < reportData.length; i++) {
                        allReportData.push(reportData[i]);
                    }
                }
            } catch (error) {
                console.error('Error processing file:', error);
            }
        }

        if (allReportData.length > 0) {
            const newWorkbook = XLSX.utils.book_new();
            const reportSheet = XLSX.utils.aoa_to_sheet(allReportData);
            XLSX.utils.book_append_sheet(newWorkbook, reportSheet, 'Combined Report');

            const reportContent = XLSX.write(newWorkbook, { bookType: 'xlsx', type: 'array' });
            const blob = new Blob([reportContent], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
            const url = URL.createObjectURL(blob);
            const downloadLink = document.getElementById('downloadLink');
            downloadLink.href = url;
            downloadLink.download = `${category}Report.xlsx`;
            downloadLink.style.display = 'block';
        } else {
            alert('No data to generate report.');
        }
    };

    processAllFiles();
}

function generateManufacturedReport(workbook) {
    console.log('generateManufacturedReport function called');

    const firstSheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[firstSheetName];
    if (sheet) {
        console.log(`Sheet ${firstSheetName} found`);

        const findCellValue = (searchStrings) => {
            const range = XLSX.utils.decode_range(sheet['!ref']);
            for (let row = range.s.r; row <= range.e.r; row++) {
                for (let col = range.s.c; col <= range.e.c; col++) {
                    const cellAddress = { c: col, r: row };
                    const cellRef = XLSX.utils.encode_cell(cellAddress);
                    const cell = sheet[cellRef];
                    if (cell && cell.v) {
                        const cellValue = cell.v.toString().toLowerCase();
                        if (searchStrings.includes(cellValue)) {
                            console.log(`Found "${cell.v}" at row ${row}, col ${col}`);
                            return { row, col };
                        }
                    }
                }
            }
            return null;
        };

        const getNextCellValue = (position) => {
            if (!position) return null;
            const nextCellRef = XLSX.utils.encode_cell({ c: position.col + 1, r: position.row });
            const nextCell = sheet[nextCellRef];
            console.log(`Next cell value at row ${position.row}, col ${position.col + 1}:`, nextCell ? nextCell.v : 'null');
            return nextCell ? nextCell.v : null;
        };

        const productName = sheet['A1'] ? sheet['A1'].v : "Unknown Product";
        console.log('Product Name:', productName);

        const lotNumberPosition = findCellValue(["lot #", "lot number"]);
        const manufactureDatePosition = findCellValue(["date", "manufacture date"]);
        const kitLotSizePosition = findCellValue(["kit lot size", "current lot size", "lot size"]);

        console.log('Lot Number Position:', lotNumberPosition);
        console.log('Manufacture Date Position:', manufactureDatePosition);
        console.log('Kit Lot Size Position:', kitLotSizePosition);

        const lotNumber = getNextCellValue(lotNumberPosition);
        let manufactureDate = getNextCellValue(manufactureDatePosition);
        const kitLotSize = getNextCellValue(kitLotSizePosition);

        if (typeof manufactureDate === 'number') {
            manufactureDate = XLSX.SSF.format("mm/dd/yyyy", manufactureDate);
        }

        console.log('Lot Number:', lotNumber);
        console.log('Manufacture Date:', manufactureDate);
        console.log('Kit Lot Size:', kitLotSize);

        if (lotNumber && manufactureDate && kitLotSize) {
            const reportData = [["Product", "Date", "Lot #", "Number of kits"], [productName, manufactureDate, lotNumber, kitLotSize]];
            return reportData;
        } else {
            console.error('One or more required cells are missing');
            return [];
        }
    } else {
        console.error('Sheet not found');
        return [];
    }
}