document.addEventListener('DOMContentLoaded', () => {
    // DOM Elements - Config
    const apiKeyInput = document.getElementById('apiKeyInput');
    const togglePasswordBtn = document.getElementById('togglePasswordBtn');

    // DOM Elements - File Dropzones
    const pdfDropZone = document.getElementById('pdfDropZone');
    const pdfFileInput = document.getElementById('pdfFileInput');
    const pdfFileInfo = document.getElementById('pdfFileInfo');
    const removePdfBtn = document.getElementById('removePdfBtn');

    const xlsxDropZone = document.getElementById('xlsxDropZone');
    const xlsxFileInput = document.getElementById('xlsxFileInput');
    const xlsxFileInfo = document.getElementById('xlsxFileInfo');
    const removeXlsxBtn = document.getElementById('removeXlsxBtn');

    // DOM Elements - Actions & Sections
    const processBtn = document.getElementById('processBtn');
    const uploadSection = document.querySelector('.upload-section');
    const configCard = document.querySelector('.config-card');
    const loadingPanel = document.getElementById('loadingPanel');
    const errorPanel = document.getElementById('errorPanel');
    const errorMsgText = document.getElementById('errorMsgText');
    const closeErrorBtn = document.getElementById('closeErrorBtn');
    const resultsPanel = document.getElementById('resultsPanel');
    const newProcessBtn = document.getElementById('newProcessBtn');

    // DOM Elements - Results Card
    const summaryPatientName = document.getElementById('summaryPatientName');
    const summaryCollectionDate = document.getElementById('summaryCollectionDate');
    const summaryAge = document.getElementById('summaryAge');
    const summaryBirthDate = document.getElementById('summaryBirthDate');
    const downloadLink = document.getElementById('downloadLink');
    const resultsCountBadge = document.getElementById('resultsCountBadge');
    const resultsTableBody = document.getElementById('resultsTableBody');

    // Local State
    let pdfFile = null;
    let xlsxFile = null;

    // 1. Load Saved API Key
    const savedKey = localStorage.getItem('gemini_api_key');
    if (savedKey) {
        apiKeyInput.value = savedKey;
    }

    // 2. Toggle API Key Visibility
    togglePasswordBtn.addEventListener('click', () => {
        const isPassword = apiKeyInput.type === 'password';
        apiKeyInput.type = isPassword ? 'text' : 'password';
        togglePasswordBtn.innerHTML = isPassword ? 
            '<i class="fa-solid fa-eye-slash"></i>' : 
            '<i class="fa-solid fa-eye"></i>';
    });

    // 3. Setup Drag and Drop Events for PDF
    setupDragDropZone(pdfDropZone, pdfFileInput, 'application/pdf', (file) => {
        pdfFile = file;
        showFileInfo(pdfDropZone, pdfFileInfo, file.name, 'pdf-icon-big fa-regular fa-file-pdf');
        validateInputs();
    });

    removePdfBtn.addEventListener('click', (e) => {
        e.stopPropagation(); // prevent triggering dropzone click
        pdfFile = null;
        resetDropZone(pdfDropZone, pdfFileInfo);
        pdfFileInput.value = '';
        validateInputs();
    });

    // 4. Setup Drag and Drop Events for Excel
    setupDragDropZone(xlsxDropZone, xlsxFileInput, '.xlsx', (file) => {
        xlsxFile = file;
        showFileInfo(xlsxDropZone, xlsxFileInfo, file.name, 'excel-icon-big fa-regular fa-file-excel');
        validateInputs();
    }, ['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet']);

    removeXlsxBtn.addEventListener('click', (e) => {
        e.stopPropagation();
        xlsxFile = null;
        resetDropZone(xlsxDropZone, xlsxFileInfo);
        xlsxFileInput.value = '';
        validateInputs();
    });

    // 5. Drag and Drop Engine Helper
    function setupDragDropZone(zoneEl, inputEl, acceptType, onFileSelected, altMimeTypes = []) {
        // Clicking on zone opens file browser
        zoneEl.addEventListener('click', () => {
            inputEl.click();
        });

        inputEl.addEventListener('change', (e) => {
            if (e.target.files.length > 0) {
                const file = e.target.files[0];
                onFileSelected(file);
            }
        });

        // Drag events
        ['dragenter', 'dragover'].forEach(eventName => {
            zoneEl.addEventListener(eventName, (e) => {
                e.preventDefault();
                e.stopPropagation();
                zoneEl.classList.add('dragover');
            }, false);
        });

        ['dragleave', 'drop'].forEach(eventName => {
            zoneEl.addEventListener(eventName, (e) => {
                e.preventDefault();
                e.stopPropagation();
                zoneEl.classList.remove('dragover');
            }, false);
        });

        zoneEl.addEventListener('drop', (e) => {
            const dt = e.dataTransfer;
            if (dt.files.length > 0) {
                const file = dt.files[0];
                const fileExtension = '.' + file.name.split('.').pop().toLowerCase();
                
                // Validate file type
                if (file.type === acceptType || fileExtension === acceptType || altMimeTypes.includes(file.type)) {
                    onFileSelected(file);
                } else {
                    alert(`Tipo de arquivo inválido. Por favor envie um arquivo ${acceptType}.`);
                }
            }
        });
    }

    function showFileInfo(zoneEl, infoEl, name, iconClasses) {
        const contentEl = zoneEl.querySelector('.drop-zone-content');
        contentEl.classList.add('hidden');
        infoEl.classList.remove('hidden');
        infoEl.querySelector('.file-name').textContent = name;
        
        // update icon
        const iconEl = infoEl.querySelector('i');
        iconEl.className = iconClasses;
    }

    function resetDropZone(zoneEl, infoEl) {
        const contentEl = zoneEl.querySelector('.drop-zone-content');
        contentEl.classList.remove('hidden');
        infoEl.classList.add('hidden');
    }

    // 6. Validate inputs to enable process button
    function validateInputs() {
        const hasApiKey = apiKeyInput.value.trim().length > 0;
        const hasPdf = pdfFile !== null;
        const hasXlsx = xlsxFile !== null;

        processBtn.disabled = !(hasApiKey && hasPdf && hasXlsx);
    }

    apiKeyInput.addEventListener('input', validateInputs);

    // 7. Process files form submission
    processBtn.addEventListener('click', async () => {
        const api_key = apiKeyInput.value.trim();
        if (!api_key || !pdfFile || !xlsxFile) return;

        // Save key locally
        localStorage.setItem('gemini_api_key', api_key);

        // UI transitions
        configCard.classList.add('hidden');
        uploadSection.classList.add('hidden');
        loadingPanel.classList.remove('hidden');
        errorPanel.classList.add('hidden');
        resultsPanel.classList.add('hidden');

        // Form data packaging
        const formData = new FormData();
        formData.append('api_key', api_key);
        formData.append('pdf', pdfFile);
        formData.append('xlsx', xlsxFile);

        try {
            const response = await fetch('/process', {
                method: 'POST',
                body: formData
            });

            const data = await response.json();

            if (!response.ok) {
                throw new Error(data.error || 'Erro desconhecido ao processar arquivos.');
            }

            // Success! Render results
            summaryPatientName.textContent = data.patient_name;
            summaryCollectionDate.textContent = data.collection_date;
            summaryAge.textContent = data.age || 'Não informado';
            summaryBirthDate.textContent = data.birth_date || 'Não informado';
            
            // Set download link
            downloadLink.href = `/download/${data.filename}`;
            resultsCountBadge.textContent = `${data.results.length} exames`;

            // Populate table
            resultsTableBody.innerHTML = '';
            data.results.forEach(res => {
                const row = document.createElement('tr');
                
                const paramTd = document.createElement('td');
                paramTd.textContent = res.parameter;
                paramTd.style.fontWeight = '500';
                
                const valueTd = document.createElement('td');
                valueTd.textContent = res.value;
                
                row.appendChild(paramTd);
                row.appendChild(valueTd);
                resultsTableBody.appendChild(row);
            });

            // UI Transitions
            loadingPanel.classList.add('hidden');
            resultsPanel.classList.remove('hidden');

        } catch (err) {
            console.error(err);
            errorMsgText.textContent = err.message;
            loadingPanel.classList.add('hidden');
            errorPanel.classList.remove('hidden');
        }
    });

    // 8. Close Error Panel
    closeErrorBtn.addEventListener('click', () => {
        errorPanel.classList.add('hidden');
        configCard.classList.remove('hidden');
        uploadSection.classList.remove('hidden');
    });

    // 9. Reset process for another exam
    newProcessBtn.addEventListener('click', () => {
        // Keep key but clear files
        pdfFile = null;
        xlsxFile = null;
        resetDropZone(pdfDropZone, pdfFileInfo);
        resetDropZone(xlsxDropZone, xlsxFileInfo);
        pdfFileInput.value = '';
        xlsxFileInput.value = '';
        validateInputs();

        // UI transitions
        resultsPanel.classList.add('hidden');
        configCard.classList.remove('hidden');
        uploadSection.classList.remove('hidden');
    });
});
