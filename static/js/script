document.addEventListener('DOMContentLoaded', function() {
    // DOM elements
    const uploadArea = document.getElementById('uploadArea');
    const fileInput = document.getElementById('fileInput');
    const browseButton = document.getElementById('browseButton');
    const fileCountSpan = document.getElementById('fileCount');
    const mergeBtn = document.getElementById('mergeBtn');
    const resetBtn = document.getElementById('resetBtn');
    const downloadBtn = document.getElementById('downloadBtn');
    const progressBar = document.getElementById('progressBar');
    const progressFill = document.getElementById('progressFill');
    const summaryFileCount = document.getElementById('summaryFileCount');
    const summarySheetCount = document.getElementById('summarySheetCount');
    const summaryRowCount = document.getElementById('summaryRowCount');
    const summaryMonthCount = document.getElementById('summaryMonthCount');
    const gstinValue = document.getElementById('gstinValue');
    const fyValue = document.getElementById('fyValue');
    const monthsValue = document.getElementById('monthsValue');
    const monthCounter = document.getElementById('monthCounter');
    const counterValue = document.getElementById('counterValue');
    const previewInfo = document.getElementById('previewInfo');
    const rowCountInfo = document.getElementById('rowCountInfo');
    const modal = document.getElementById('reportModal');
    const closeModal = document.getElementById('closeModal');
    const closeModalBtn = document.getElementById('closeModalBtn');
    const confirmDownload = document.getElementById('confirmDownload');
    const modalFileCount = document.getElementById('modalFileCount');
    const modalMonthCount = document.getElementById('modalMonthCount');
    const modalRowCount = document.getElementById('modalRowCount');
    const modalGSTIN = document.getElementById('modalGSTIN');
    const previewTabs = document.getElementById('previewTabs');
    const previewTabContent = document.getElementById('previewTabContent');

    // Return type buttons
    const gstr1Btn = document.getElementById('gstr1Btn');
    const gstr3bBtn = document.getElementById('gstr3bBtn');
    const gstr2aBtn = document.getElementById('gstr2aBtn');
    let selectedReturnType = 'GSTR-3B'; // default

    // Info tabs
    const infoTabs = document.querySelectorAll('.info-tab');
    const tabPanes = document.querySelectorAll('.tab-pane-info');

    infoTabs.forEach(tab => {
        tab.addEventListener('click', () => {
            const target = tab.getAttribute('data-tab');
            infoTabs.forEach(t => t.classList.remove('active'));
            tabPanes.forEach(p => p.classList.remove('active'));
            tab.classList.add('active');
            document.getElementById(target + '-tab').classList.add('active');
        });
    });

    let currentToken = null;
    let previewData = null;

    // Stats counters
    let totalProcessed = parseInt(localStorage.getItem('totalProcessed') || '0');
    let todayProcessed = parseInt(localStorage.getItem('todayProcessed') || '0');
    const today = new Date().toDateString();
    if (localStorage.getItem('statsDate') !== today) {
        todayProcessed = 0;
        localStorage.setItem('statsDate', today);
    }
    document.getElementById('totalFilesCounter').textContent = totalProcessed;
    document.getElementById('todayFilesCounter').textContent = todayProcessed;
    document.getElementById('currentDate').textContent = new Date().toLocaleDateString();

    function updateStats() {
        totalProcessed++;
        todayProcessed++;
        localStorage.setItem('totalProcessed', totalProcessed);
        localStorage.setItem('todayProcessed', todayProcessed);
        localStorage.setItem('statsDate', new Date().toDateString());
        document.getElementById('totalFilesCounter').textContent = totalProcessed;
        document.getElementById('todayFilesCounter').textContent = todayProcessed;
    }

    // Return type toggle
    gstr1Btn.addEventListener('click', () => {
        gstr1Btn.classList.add('active', 'btn-primary');
        gstr1Btn.classList.remove('btn-secondary');
        gstr3bBtn.classList.remove('active', 'btn-primary');
        gstr3bBtn.classList.add('btn-secondary');
        gstr2aBtn.classList.remove('active', 'btn-primary');
        gstr2aBtn.classList.add('btn-secondary');
        selectedReturnType = 'GSTR-1';
    });

    gstr3bBtn.addEventListener('click', () => {
        gstr3bBtn.classList.add('active', 'btn-primary');
        gstr3bBtn.classList.remove('btn-secondary');
        gstr1Btn.classList.remove('active', 'btn-primary');
        gstr1Btn.classList.add('btn-secondary');
        gstr2aBtn.classList.remove('active', 'btn-primary');
        gstr2aBtn.classList.add('btn-secondary');
        selectedReturnType = 'GSTR-3B';
    });

    gstr2aBtn.addEventListener('click', () => {
        gstr2aBtn.classList.add('active', 'btn-primary');
        gstr2aBtn.classList.remove('btn-secondary');
        gstr1Btn.classList.remove('active', 'btn-primary');
        gstr1Btn.classList.add('btn-secondary');
        gstr3bBtn.classList.remove('active', 'btn-primary');
        gstr3bBtn.classList.add('btn-secondary');
        selectedReturnType = 'GSTR-2A';
    });

    // File selection via browse button
    browseButton.addEventListener('click', () => fileInput.click());

    fileInput.addEventListener('change', handleFileSelect);

    // Drag and drop events
    uploadArea.addEventListener('dragover', (e) => {
        e.preventDefault();
        uploadArea.classList.add('drag-over');
    });

    uploadArea.addEventListener('dragleave', () => {
        uploadArea.classList.remove('drag-over');
    });

    uploadArea.addEventListener('drop', (e) => {
        e.preventDefault();
        uploadArea.classList.remove('drag-over');
        fileInput.files = e.dataTransfer.files;
        handleFileSelect();
    });

    function handleFileSelect() {
        const files = fileInput.files;
        fileCountSpan.textContent = files.length;
        summaryFileCount.textContent = files.length;
        mergeBtn.disabled = files.length === 0;
    }

    // Reset button
    resetBtn.addEventListener('click', () => {
        fileInput.value = '';
        fileCountSpan.textContent = '0';
        summaryFileCount.textContent = '0';
        summarySheetCount.textContent = '0';
        summaryRowCount.textContent = '0';
        summaryMonthCount.textContent = '0';
        mergeBtn.disabled = true;
        downloadBtn.disabled = true;
        gstinValue.textContent = 'Not Available';
        fyValue.textContent = 'Not Available';
        monthsValue.textContent = '0';
        monthCounter.style.display = 'none';
        previewInfo.textContent = 'No data parsed. Upload JSON files and click "Consolidate Files".';
        rowCountInfo.textContent = 'Showing 0 rows';
        previewTabs.innerHTML = '';
        previewTabContent.innerHTML = '';
        currentToken = null;
        previewData = null;
    });

    // Merge / Consolidate button with loading animation
    mergeBtn.addEventListener('click', async () => {
        const files = fileInput.files;
        if (files.length === 0) return;

        // Show progress and loading state on button
        progressBar.style.display = 'block';
        progressFill.style.width = '10%';
        mergeBtn.disabled = true;
        const originalText = mergeBtn.innerHTML;
        mergeBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Processing...';

        const formData = new FormData();
        for (let file of files) {
            formData.append('files[]', file);
        }
        formData.append('returnType', selectedReturnType);

        try {
            const response = await fetch('/upload', {
                method: 'POST',
                body: formData
            });
            const result = await response.json();
            if (result.error) {
                alert('Error: ' + result.error);
                mergeBtn.disabled = false;
                mergeBtn.innerHTML = originalText;
                progressBar.style.display = 'none';
                return;
            }

            progressFill.style.width = '100%';
            setTimeout(() => {
                progressBar.style.display = 'none';
                progressFill.style.width = '0%';
            }, 500);

            // Update summary cards
            summaryFileCount.textContent = result.fileCount;
            summarySheetCount.textContent = result.sheetCount;
            summaryRowCount.textContent = result.rowCount;
            summaryMonthCount.textContent = result.monthCount;
            monthsValue.textContent = result.monthCount;
            gstinValue.textContent = result.preview.meta.gstin || 'Not Available';
            fyValue.textContent = result.preview.meta.financial_year || 'Not Available';
            monthCounter.style.display = 'flex';
            counterValue.textContent = result.monthCount;

            // Build preview based on type
            if (result.preview.type === 'GSTR-1') {
                buildGSTR1Preview(result.preview);
            } else if (result.preview.type === 'GSTR-2A') {
                buildGSTR2APreview(result.preview);
            } else {
                buildGSTR3BPreview(result.preview);
            }

            currentToken = result.token;
            previewData = result.preview;
            downloadBtn.disabled = false;
            mergeBtn.disabled = false;
            mergeBtn.innerHTML = originalText;

            updateStats();

        } catch (err) {
            console.error(err);
            alert('Upload failed. Check console.');
            mergeBtn.disabled = false;
            mergeBtn.innerHTML = originalText;
            progressBar.style.display = 'none';
        }
    });

    function buildGSTR3BPreview(data) {
        const months = data.meta.months;
        const rows = data.rows;

        previewTabs.innerHTML = '';
        previewTabContent.innerHTML = '';

        let headerHtml = '<tr><th>S.No</th><th>Particulars</th>';
        months.forEach(m => {
            headerHtml += `<th>${m}</th>`;
        });
        headerHtml += '<th>Total</th></tr>';

        let bodyHtml = '';
        rows.forEach((r, idx) => {
            bodyHtml += '<tr>';
            bodyHtml += `<td>${idx + 1}</td>`;
            bodyHtml += `<td>${r.label}</td>`;
            r.values.forEach(v => {
                const val = formatIndianNumber(v);
                bodyHtml += `<td class="${v === 0 ? 'zero-value' : ''}">${val}</td>`;
            });
            const total = formatIndianNumber(r.total);
            bodyHtml += `<td class="${r.total === 0 ? 'zero-value' : ''}">${total}</td>`;
            bodyHtml += '</tr>';
        });

        previewTabContent.innerHTML = `
            <div class="table-wrapper">
                <table class="preview-table">
                    <thead>${headerHtml}</thead>
                    <tbody>${bodyHtml}</tbody>
                </table>
            </div>
        `;

        previewInfo.textContent = `Consolidated report for ${months.length} months.`;
        rowCountInfo.textContent = `Showing ${rows.length} rows`;
    }

    function buildGSTR1Preview(data) {
        previewTabs.innerHTML = '';
        previewTabContent.innerHTML = '';

        const sheetNames = ['b2b', 'b2cs', 'cdnr', 'hsn'];
        sheetNames.forEach((name, idx) => {
            const sheet = data.sheets[name];
            if (!sheet) return;

            const tabBtn = document.createElement('button');
            tabBtn.className = `preview-tab ${idx === 0 ? 'active' : ''}`;
            tabBtn.setAttribute('data-tab', name);
            let displayName = name.toUpperCase();
            if (name === 'b2b') displayName = 'B2B, SEZ, DE';
            tabBtn.textContent = displayName;
            previewTabs.appendChild(tabBtn);

            const pane = document.createElement('div');
            pane.className = `tab-pane ${idx === 0 ? 'active' : ''}`;
            pane.id = `tab-${name}`;

            let tableHtml = '<div class="table-wrapper"><table class="preview-table"><thead><tr>';
            sheet.columns.forEach(col => {
                tableHtml += `<th>${col}</th>`;
            });
            tableHtml += '</tr></thead><tbody>';
            sheet.rows.forEach(row => {
                tableHtml += '<tr>';
                row.forEach(cell => {
                    tableHtml += `<td>${cell}</td>`;
                });
                tableHtml += '</tr>';
            });
            tableHtml += '</tbody></table></div>';
            pane.innerHTML = tableHtml;
            previewTabContent.appendChild(pane);
        });

        document.querySelectorAll('.preview-tab').forEach(tab => {
            tab.addEventListener('click', () => {
                document.querySelectorAll('.preview-tab').forEach(t => t.classList.remove('active'));
                document.querySelectorAll('.tab-pane').forEach(p => p.classList.remove('active'));
                tab.classList.add('active');
                document.getElementById(`tab-${tab.dataset.tab}`).classList.add('active');
            });
        });

        previewInfo.textContent = `Consolidated GSTR-1 for ${data.meta.months.length} months.`;
        rowCountInfo.textContent = `Showing multiple sheets.`;
    }

    function buildGSTR2APreview(data) {
        previewTabs.innerHTML = '';
        previewTabContent.innerHTML = '';

        const sheets = ['b2b', 'cdn', 'tcs'];
        sheets.forEach((name, idx) => {
            const sheet = data.sheets[name];
            if (!sheet) return;

            const tabBtn = document.createElement('button');
            tabBtn.className = `preview-tab ${idx === 0 ? 'active' : ''}`;
            tabBtn.setAttribute('data-tab', name);
            tabBtn.textContent = name.toUpperCase();
            previewTabs.appendChild(tabBtn);

            const pane = document.createElement('div');
            pane.className = `tab-pane ${idx === 0 ? 'active' : ''}`;
            pane.id = `tab-${name}`;

            let tableHtml = '<div class="table-wrapper"><table class="preview-table"><thead><tr>';
            sheet.columns.forEach(col => {
                tableHtml += `<th>${col}</th>`;
            });
            tableHtml += '</tr></thead><tbody>';
            sheet.rows.forEach(row => {
                tableHtml += '<tr>';
                row.forEach(cell => {
                    const val = (typeof cell === 'number') ? formatIndianNumber(cell) : (cell || '');
                    tableHtml += `<td>${val}</td>`;
                });
                tableHtml += '</tr>';
            });
            tableHtml += '</tbody></table></div>';
            pane.innerHTML = tableHtml;
            previewTabContent.appendChild(pane);
        });

        document.querySelectorAll('.preview-tab').forEach(tab => {
            tab.addEventListener('click', () => {
                document.querySelectorAll('.preview-tab').forEach(t => t.classList.remove('active'));
                document.querySelectorAll('.tab-pane').forEach(p => p.classList.remove('active'));
                tab.classList.add('active');
                document.getElementById(`tab-${tab.dataset.tab}`).classList.add('active');
            });
        });

        previewInfo.textContent = `Consolidated GSTR-2A for ${data.meta.months.length} months.`;
        rowCountInfo.textContent = `Showing multiple sheets.`;
    }

    // Indian number formatting for preview
    function formatIndianNumber(num) {
        if (num === null || num === undefined) return "0.00";
        let n = parseFloat(num);
        if (isNaN(n)) return num;
        let s = n.toFixed(2);
        let [intPart, decPart] = s.split('.');
        if (intPart.length > 3) {
            let lastThree = intPart.slice(-3);
            let rest = intPart.slice(0, -3);
            let restGroups = [];
            while (rest.length > 0) {
                restGroups.push(rest.slice(-2));
                rest = rest.slice(0, -2);
            }
            restGroups.reverse();
            intPart = restGroups.join(',') + ',' + lastThree;
        }
        return intPart + '.' + decPart;
    }

    // Download button click -> show modal
    downloadBtn.addEventListener('click', () => {
        if (!currentToken || !previewData) return;
        modalFileCount.textContent = summaryFileCount.textContent;
        modalMonthCount.textContent = previewData.meta.months.length;
        modalRowCount.textContent = summaryRowCount.textContent;
        modalGSTIN.textContent = previewData.meta.gstin || '-';
        modal.style.display = 'flex';
    });

    // Modal close
    closeModal.addEventListener('click', () => modal.style.display = 'none');
    closeModalBtn.addEventListener('click', () => modal.style.display = 'none');
    window.addEventListener('click', (e) => {
        if (e.target === modal) modal.style.display = 'none';
    });

    // Confirm download
    confirmDownload.addEventListener('click', () => {
        if (!currentToken) return;
        window.location.href = `/download?token=${currentToken}`;
        modal.style.display = 'none';
    });

    // Instructions popup (opens automatically on first visit)
    const instructionModal = document.getElementById('instructionModal');
    const closeInstructionModal = document.getElementById('closeInstructionModal');
    const viewInstructionsBtn = document.getElementById('viewInstructionsBtn');

    if (!localStorage.getItem('instructionsShown')) {
        instructionModal.style.display = 'flex';
        localStorage.setItem('instructionsShown', 'true');
    }

    closeInstructionModal.addEventListener('click', () => {
        instructionModal.style.display = 'none';
    });

    viewInstructionsBtn.addEventListener('click', () => {
        instructionModal.style.display = 'flex';
    });

    window.addEventListener('click', (e) => {
        if (e.target === instructionModal) {
            instructionModal.style.display = 'none';
        }
    });

    // Initialize: set summary sheet count to 0
    summarySheetCount.textContent = '0';
});
