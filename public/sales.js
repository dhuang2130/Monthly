function generateReport(category) {
    console.log('generateReport function called with category:', category);
    const fileInput = document.getElementById('fileInput');
    const keyInput = document.getElementById('keyInput'); // Assuming you have an input for the key file

    if (fileInput.files.length === 0 || keyInput.files.length === 0) {
        alert('Please select the required Excel files.');
        return;
    }

    const file = fileInput.files[0];
    const keyFile = keyInput.files[0];
    console.log('Selected files:', file.name, keyFile.name);

    const reader = new FileReader();
    const keyReader = new FileReader();

    reader.onload = (e) => {
        console.log('FileReader onload called');
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
            alert('Error reading Excel file. Please ensure the file is not corrupted and try again.');
            return;
        }

        keyReader.onload = (ke) => {
            const keyData = new Uint8Array(ke.target.result);
            let keyWorkbook;

            try {
                if (keyFile.name.endsWith('.xls')) {
                    keyWorkbook = XLSX.read(ke.target.result, { type: 'binary' });
                } else {
                    keyWorkbook = XLSX.read(keyData, { type: 'array' });
                }
            } catch (error) {
                console.error('Error reading key workbook:', error);
                alert('Error reading key Excel file. Please ensure the file is not corrupted and try again.');
                return;
            }

            // Generate sales report using the key
            const reportData = generateSalesReport(workbook, keyWorkbook);

            // Check if reportData is generated
            if (reportData) {
                // Create a blob and link for downloading
                const blob = new Blob([reportData], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
                const url = URL.createObjectURL(blob);

                // Generate dynamic filename
                const fileName = file.name.replace(/(\.xlsx|\.xls)$/, 'Report$1');
                
                const downloadLink = document.getElementById('downloadLink');
                downloadLink.href = url;
                downloadLink.download = fileName;
                downloadLink.style.display = 'block';
                console.log('Download link created for Sales Report');
            } else {
                console.error('Failed to generate report data');
            }
        };

        if (keyFile.name.endsWith('.xls')) {
            keyReader.readAsBinaryString(keyFile);
        } else {
            keyReader.readAsArrayBuffer(keyFile);
        }
    };

    if (file.name.endsWith('.xls')) {
        reader.readAsBinaryString(file);
    } else {
        reader.readAsArrayBuffer(file);
    }

    reader.onerror = keyReader.onerror = (error) => {
        console.error('File reading error:', error);
    };
}

function generateSalesReport(workbook, keyWorkbook) {
    console.log('generateSalesReport function called');

    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const keySheet = keyWorkbook.Sheets[keyWorkbook.SheetNames[0]];
    
    if (!sheet || !keySheet) {
        console.error('Sheet not found in one of the workbooks');
        return null;
    }

    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    const keyRows = XLSX.utils.sheet_to_json(keySheet, { header: 1 });
    const validProducts = new Set(keyRows.map(row => row[0])); // Assuming the first column in the key file contains product names

    console.log('Rows extracted from sheet:', rows);
    console.log('Valid products from key:', Array.from(validProducts));

    const productQuantities = {};

    // Define the product pattern
    const productPattern = /(\d+)\s+(\w+(-\w+)*|\w+)/g;

    // Extract unique months from the data with shortened names
    const dateColumn = rows.map(row => row[0]).filter(date => typeof date === 'string' && !isNaN(Date.parse(date)));
    const uniqueMonths = [...new Set(dateColumn.map(date => new Date(date).toLocaleString('default', { month: 'short' })))];

    uniqueMonths.sort((a, b) => new Date(`01 ${a} 2020`) - new Date(`01 ${b} 2020`)); // Ensure correct ascending order
    console.log('Unique months identified:', uniqueMonths);

    // Initialize product quantities for all valid products
    validProducts.forEach(product => {
        productQuantities[product] = uniqueMonths.reduce((acc, month) => {
            acc[month] = 0;
            return acc;
        }, { Total: 0 });
    });

    // Calculate quantities for each product and month
    rows.forEach((row) => {
        const purchaseCell = row[1];  // Assuming the 'Purchase' column is the second column in the sheet
        const date = row[0]; // Assuming the 'Date' column is the first column in the sheet
        if (typeof purchaseCell === 'string' && typeof date === 'string' && !isNaN(Date.parse(date))) {
            const currentMonth = new Date(date).toLocaleString('default', { month: 'short' });
            let match;
            while ((match = productPattern.exec(purchaseCell)) !== null) {
                const [quantity, item] = [parseInt(match[1], 10), match[2]];
                if (!isNaN(quantity) && validProducts.has(item)) {
                    productQuantities[item][currentMonth] += quantity;
                    productQuantities[item].Total += quantity;
                }
            }
        }
    });

    // Sort products alphabetically
    const sortedProducts = Array.from(validProducts).sort();

    // Create report data
    const reportData = [['Product', ...uniqueMonths, 'Total']];
    sortedProducts.forEach(product => {
        const row = [product, ...uniqueMonths.map(month => productQuantities[product]?.[month] || 0), productQuantities[product]?.Total || 0];
        reportData.push(row);
    });

    console.log('Final report data:', reportData);

    const newWorkbook = XLSX.utils.book_new();
    const reportSheet = XLSX.utils.aoa_to_sheet(reportData);
    XLSX.utils.book_append_sheet(newWorkbook, reportSheet, 'Sales Report');

    const reportContent = XLSX.write(newWorkbook, { bookType: 'xlsx', type: 'array' });
    return reportContent;
}
