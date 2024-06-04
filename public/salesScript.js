function generateReport(category) {
    console.log('generateReport function called with category:', category);
    const fileInput = document.getElementById('fileInput');

    if (fileInput.files.length === 0) {
        alert('Please select the required Excel file.');
        return;
    }

    const file = fileInput.files[0];
    console.log('Selected file:', file.name);

    const reader = new FileReader();
    reader.onload = (e) => {
        console.log('FileReader onload called');
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });

        console.log('Workbook loaded');

        // Generate sales report
        const reportData = generateSalesReport(workbook);

        // Check if reportData is generated
        if (reportData) {
            // Create a blob and link for downloading
            const blob = new Blob([reportData], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
            const url = URL.createObjectURL(blob);
            const downloadLink = document.getElementById('downloadLink');
            downloadLink.href = url;
            downloadLink.download = `SalesReport.xlsx`;
            downloadLink.style.display = 'block';
            console.log('Download link created for Sales Report');
        } else {
            console.error('Failed to generate report data');
        }
    };
    reader.readAsArrayBuffer(file);
    reader.onerror = (error) => {
        console.error('File reading error:', error);
    };
}

function generateSalesReport(workbook) {
    console.log('generateSalesReport function called');

    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    
    if (!sheet) {
        console.error('No sheet found in the workbook');
        return null;
    }

    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    console.log('Rows extracted from sheet:', rows);
    
    const productQuantities = {};

    // Extract unique months from the data
    const dateColumn = rows.map(row => row[0]).filter(date => typeof date === 'string' && !isNaN(Date.parse(date)));
    const uniqueMonths = [...new Set(dateColumn.map(date => new Date(date).toLocaleString('default', { month: 'long', year: 'numeric' })))].sort((a, b) => new Date(a) - new Date(b));
    console.log('Unique months identified:', uniqueMonths);

    // Extract dynamic products from the current data
    const productPattern = /(\d+)\s+(\w+(-\w+)*|\w+)/g;
    const validProducts = new Set();
    rows.forEach(row => {
        const purchaseCell = row[1];  // Assuming the 'Purchase' column is the second column in the sheet
        if (typeof purchaseCell === 'string') {
            let match;
            while ((match = productPattern.exec(purchaseCell)) !== null) {
                const [, , item] = match;
                validProducts.add(item);
                console.log(`Identified product: ${item}`);
            }
        }
    });

    console.log('Valid products identified:', Array.from(validProducts));

    // Initialize product quantities for all valid products
    validProducts.forEach(product => {
        productQuantities[product] = uniqueMonths.reduce((acc, month) => {
            acc[month] = 0;
            return acc;
        }, { Total: 0 });
    });

    // Track detailed quantities and row indices for QQHIST10 in May
    let qqhist10MayDetails = [];

    // Calculate quantities for each product and month
    rows.forEach((row, rowIndex) => {
        const purchaseCell = row[1];  // Assuming the 'Purchase' column is the second column in the sheet
        const date = row[0]; // Assuming the 'Date' column is the first column in the sheet
        if (typeof purchaseCell === 'string' && typeof date === 'string' && !isNaN(Date.parse(date))) {
            const currentMonth = new Date(date).toLocaleString('default', { month: 'long', year: 'numeric' });
            let match;
            while ((match = productPattern.exec(purchaseCell)) !== null) {
                const [quantity, item] = [parseInt(match[1], 10), match[2]];
                if (!isNaN(quantity) && validProducts.has(item)) {
                    productQuantities[item][currentMonth] += quantity;
                    productQuantities[item].Total += quantity;
                    if (item === 'QQHIST10' && currentMonth === 'May 2024') {
                        qqhist10MayDetails.push({ quantity, rowIndex });
                        console.log(`Added QQHIST10 Quantity: ${quantity} for May 2024, Row Index: ${rowIndex}, Current Total: ${productQuantities[item][currentMonth]}`);
                    }
                }
            }
        }
    });

    // Output product quantities for debugging
    console.log('Final Product Quantities:', JSON.stringify(productQuantities, null, 2));

    // Sort products alphabetically
    const sortedProducts = Array.from(validProducts).sort();

    // Create report data
    const reportData = [['Product', ...uniqueMonths, 'Total']];
    sortedProducts.forEach(product => {
        const row = [product, ...uniqueMonths.map(month => productQuantities[product]?.[month] || 0), productQuantities[product]?.Total || 0];
        if (product === 'QQHIST10') {
            console.log(`Adding row for QQHIST10: ${row}`);
        }
        reportData.push(row);
    });

    console.log('Final report data:', reportData);

    // Log detailed quantities and row indices for QQHIST10 in May
    console.log(`Details for QQHIST10 in May 2024:`, qqhist10MayDetails);
    console.log(`Final total quantity for QQHIST10 in May 2024: ${productQuantities['QQHIST10']['May 2024']}`);

    const newWorkbook = XLSX.utils.book_new();
    const reportSheet = XLSX.utils.aoa_to_sheet(reportData);
    XLSX.utils.book_append_sheet(newWorkbook, reportSheet, 'Sales Report');

    const reportContent = XLSX.write(newWorkbook, { bookType: 'xlsx', type: 'array' });
    return reportContent;
}
