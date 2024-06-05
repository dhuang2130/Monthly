document.addEventListener('DOMContentLoaded', () => {
    const categorySelect = document.getElementById('categorySelect');
    const generateReportBtn = document.getElementById('generateReportBtn');
    const fileInput = document.getElementById('fileInput');
    const keyInput = document.getElementById('keyInput');
    const keyInputContainer = document.getElementById('keyInputContainer');

    const loadScript = (src, callback) => {
        const existingScript = document.querySelector(`script[src="${src}"]`);
        if (existingScript) {
            callback();
            return;
        }
        const script = document.createElement('script');
        script.src = src;
        script.onload = callback;
        document.head.appendChild(script);
    };

    categorySelect.addEventListener('change', () => {
        const category = categorySelect.value;
        if (category === 'Manufactured') {
            fileInput.setAttribute('webkitdirectory', '');
            keyInputContainer.style.display = 'none';
        } else {
            fileInput.removeAttribute('webkitdirectory');
            keyInputContainer.style.display = 'block';
        }
    });

    generateReportBtn.addEventListener('click', () => {
        const category = categorySelect.value;
        console.log(`Loading script for category: ${category}`);
        if (category === 'Sales') {
            loadScript('sales.js', () => {
                console.log('sales.js loaded');
                if (typeof window.generateReport === 'function') {
                    window.generateReport(category);
                } else {
                    console.error('generateReport function is not defined in sales.js');
                }
            });
        } else if (category === 'Manufactured') {
            loadScript('script.js', () => {
                console.log('script.js loaded');
                if (typeof window.generateReport === 'function') {
                    window.generateReport(category);
                } else {
                    console.error('generateReport function is not defined in script.js');
                }
            });
        }
    });
});
