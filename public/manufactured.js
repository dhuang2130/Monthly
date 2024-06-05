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
        const kitLotSizePosition = findCellValue(["kit lot size", "current lot size"]);

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