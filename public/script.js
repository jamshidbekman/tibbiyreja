document.addEventListener('DOMContentLoaded', () => {
    const dropArea = document.getElementById('drop-area');
    const fileElem = document.getElementById('fileElem');
    const fileNameDisplay = document.getElementById('file-name');
    const addRangeBtn = document.getElementById('addRangeBtn');
    const rangesContainer = document.getElementById('ranges-container');
    const rangeTemplate = document.getElementById('range-template');
    const generateBtn = document.getElementById('generateBtn');
    const loadingArea = document.getElementById('loading-area');
    const resultArea = document.getElementById('result-area');
    const downloadLink = document.getElementById('downloadLink');

    let selectedFile = null;
    let currentYearCounts = {}; // Stores year analysis data

    // --- File Handling & Analysis ---

    // File Input Change
    fileElem.addEventListener('change', () => {
        handleFiles(fileElem.files);
    });

    // Drag & Drop
    ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
        dropArea.addEventListener(eventName, preventDefaults, false);
    });

    function preventDefaults(e) {
        e.preventDefault();
        e.stopPropagation();
    }

    ['dragenter', 'dragover'].forEach(eventName => {
        dropArea.addEventListener(eventName, () => dropArea.classList.add('highlight'), false);
    });

    ['dragleave', 'drop'].forEach(eventName => {
        dropArea.addEventListener(eventName, () => dropArea.classList.remove('highlight'), false);
    });

    dropArea.addEventListener('drop', (e) => {
        const dt = e.dataTransfer;
        const files = dt.files;
        handleFiles(files);
    }, false);

    // Main File Handler
    async function handleFiles(files) {
        if (!files.length) return;
        selectedFile = files[0];
        fileNameDisplay.textContent = `Tanlandi: ${selectedFile.name}`;
        fileNameDisplay.style.display = 'block';

        // Upload to analyze immediately
        const formData = new FormData();
        formData.append('file', selectedFile);

        try {
            const response = await fetch('/api/analyze', {
                method: 'POST',
                body: formData
            });

            if (response.ok) {
                const result = await response.json();
                currentYearCounts = result.yearCounts || {};
                console.log('Analysis Result:', currentYearCounts);


                // Display Total Patients
                if (result.totalPatients !== undefined) {
                    let fileInfo = document.getElementById('file-info');
                    if (!fileInfo) {
                        const div = document.createElement('div');
                        div.id = 'file-info';
                        div.className = 'file-info-badge';
                        document.getElementById('drop-area').appendChild(div);
                        fileInfo = div; // Assign the newly created div
                    }
                    fileInfo.textContent = `Jami topilgan aholi: ${result.totalPatients}`;
                }

                // Trigger validation updates for all existing ranges
                document.querySelectorAll('.range-item').forEach(updateRangeValidation);
            }
        } catch (e) {
            console.error("Analysis failed", e);
        }
    }

    // --- Range Management & Validation ---

    function updateRangeValidation(rangeItem) {
        const startInput = rangeItem.querySelector('.start-year');
        const endInput = rangeItem.querySelector('.end-year');
        const monthInputs = rangeItem.querySelectorAll('.month-input input');

        let statsDiv = rangeItem.querySelector('.range-stats');
        if (!statsDiv) {
            statsDiv = document.createElement('div');
            statsDiv.className = 'range-stats';
            // Insert after title/remove button row
            rangeItem.insertBefore(statsDiv, rangeItem.children[1]);
        }

        const startYear = parseInt(startInput.value) || 0;
        const endYear = parseInt(endInput.value) || 0;
        const gender = rangeItem.querySelector('.gender-select').value;

        // Calculate Available
        let available = 0;
        for (let y = startYear; y <= endYear; y++) {
            const counts = currentYearCounts[y];
            if (counts) {
                // Check if old structure (number) or new structure (object)
                if (typeof counts === 'number') {
                    available += counts;
                } else {
                    if (gender === 'male') available += (counts.male || 0);
                    else if (gender === 'female') available += (counts.female || 0);
                    else available += (counts.total || 0);
                }
            }
        }

        // Calculate Planned
        let planned = 0;
        monthInputs.forEach(input => planned += (parseInt(input.value) || 0));

        // Update UI
        const isError = planned > available;
        const isWarning = planned < available;
        const isOk = planned > 0 && planned === available;

        let statusText = '';
        let statusClass = '';


        const visitCount = parseInt(rangeItem.querySelector('.visit-count').value) || 1;
        // Skip validation if visitCount > 1 (Auto Plan)
        if (visitCount > 1) {
            statsDiv.textContent = `Mavjud: ${available} | Reja: HAMMASI (${available})`;
            statsDiv.className = 'range-stats stat-ok';
            return;
        }

        if (available === 0 && startYear > 0) {
            statusText = `Mavjud: 0 (Bunday yilda tug'ilgan aholi topilmadi)`;
            statusClass = 'stat-error';
        } else if (startYear === 0 && endYear === 0) {
            statusText = 'Yil oralig\'ini kiriting';
            statusClass = '';
        } else {
            statusText = `Mavjud: ${available} | Reja: ${planned}`;
            if (isError) statusClass = 'stat-error';
            else if (isWarning) statusClass = 'stat-warning';
            else if (isOk) statusClass = 'stat-ok';
        }

        statsDiv.textContent = statusText;
        statsDiv.className = `range-stats ${statusClass}`;
    }

    function addRange() {
        const clone = rangeTemplate.content.cloneNode(true);
        const rangeItem = clone.querySelector('.range-item');
        const rangeId = Date.now();
        rangeItem.dataset.id = rangeId;

        // Index
        const currentRangeCount = rangesContainer.querySelectorAll('.range-item').length + 1;
        rangeItem.querySelector('.range-index').textContent = currentRangeCount;

        // Delete handler
        rangeItem.querySelector('.delete-range').addEventListener('click', () => {
            rangeItem.remove();
            // Re-index remaining ranges
            rangesContainer.querySelectorAll('.range-item').forEach((item, idx) => {
                item.querySelector('.range-index').textContent = idx + 1;
            });
        });


        // Listen for Visit Count Changes
        const visitCountSelect = rangeItem.querySelector('.visit-count');
        const monthInputsContainer = rangeItem.querySelector('.months-grid');
        const birthdayOption = rangeItem.querySelector('.birthday-option');
        const startInput = rangeItem.querySelector('.start-year');
        const endInput = rangeItem.querySelector('.end-year');
        const genderSelect = rangeItem.querySelector('.gender-select');

        const updateUIState = () => {
            const count = parseInt(visitCountSelect.value);

            // Birthday Option Visibility
            if (count > 1) {
                birthdayOption.classList.remove('hidden');
                monthInputsContainer.style.opacity = '0.5';
                monthInputsContainer.style.pointerEvents = 'none';

                // Reset inputs to 0 to avoid confusion, they are now "Auto"
                // But we need to handle validation differently. 
                // Let's just visually disable them. Logic will handle it.
            } else {
                birthdayOption.classList.add('hidden');
                monthInputsContainer.style.opacity = '1';
                monthInputsContainer.style.pointerEvents = 'auto';
            }
            updateRangeValidation(rangeItem);
        };

        visitCountSelect.addEventListener('change', updateUIState);
        genderSelect.addEventListener('change', () => updateRangeValidation(rangeItem));

        // Auto Distribute Logic
        const distributeBtn = rangeItem.querySelector('.distribute-btn');
        distributeBtn.addEventListener('click', () => {
            // 1. Get Available Count (Logic duplicated from updateRangeValidation, should prob extract but for now inline is safe)
            const sYear = parseInt(startInput.value) || 0;
            const eYear = parseInt(endInput.value) || 0;
            const g = genderSelect.value;
            let totalAvailable = 0;

            for (let y = sYear; y <= eYear; y++) {
                const counts = currentYearCounts[y];
                if (counts) {
                    if (typeof counts === 'number') totalAvailable += counts;
                    else {
                        if (g === 'male') totalAvailable += (counts.male || 0);
                        else if (g === 'female') totalAvailable += (counts.female || 0);
                        else totalAvailable += (counts.total || 0);
                    }
                }
            }

            if (totalAvailable <= 0) {
                alert("Taqsimlash uchun aholi mavjud emas (Mavjud: 0)");
                return;
            }

            // 2. Distribute Evenly across 12 months
            const base = Math.floor(totalAvailable / 12);
            let remainder = totalAvailable % 12;

            const inputs = rangeItem.querySelectorAll('.month-input input');
            inputs.forEach((input, index) => {
                let val = base;
                if (remainder > 0) {
                    val++;
                    remainder--;
                }
                input.value = val;
            });

            // 3. Update Validation Status
            updateRangeValidation(rangeItem);
        });

        // Add Listeners for Real-time validation
        const inputs = rangeItem.querySelectorAll('input');
        inputs.forEach(input => {
            input.addEventListener('input', () => updateRangeValidation(rangeItem));
        });

        rangesContainer.appendChild(rangeItem);

        // Initial check
        if (Object.keys(currentYearCounts).length > 0) {
            updateRangeValidation(rangeItem);
        }

        // Initial UI State
        updateUIState();
    }

    addRangeBtn.addEventListener('click', addRange);

    // Initialize with one empty range
    addRange();

    // --- Generation Logic ---

    generateBtn.addEventListener('click', async () => {
        if (!selectedFile) {
            alert('Iltimos, avval Excel faylni yuklang.');
            return;
        }

        // Collect Config
        const ranges = [];
        const rangeItems = document.querySelectorAll('.range-item');
        let hasError = false;

        for (const item of rangeItems) {
            const startYear = parseInt(item.querySelector('.start-year').value);
            const endYear = parseInt(item.querySelector('.end-year').value);
            const visitCount = parseInt(item.querySelector('.visit-count').value) || 1;
            const gender = item.querySelector('.gender-select').value;
            const useBirthday = item.querySelector('.use-birthday').checked;

            if (isNaN(startYear) || isNaN(endYear)) {
                alert('Iltimos, barcha oraliqlar uchun boshlanish va tugash yillarini to\'g\'ri kiriting.');
                return;
            }

            const counts = [];
            for (let i = 0; i < 12; i++) {
                const val = parseInt(item.querySelector(`.m-${i}`).value) || 0;
                counts.push(val);
            }

            ranges.push({
                startYear,
                endYear,
                visitCount,
                gender,
                useBirthday,
                counts
            });
        }

        if (ranges.length === 0) {
            alert('Iltimos, kamida bitta oraliq qo\'shing.');
            return;
        }

        const targetYear = parseInt(document.getElementById('targetYear').value) || 2026;
        const saturdayWorking = document.getElementById('saturdayWorking').checked;
        const holidaysText = document.getElementById('holidays').value;
        const holidays = holidaysText.split(',').map(s => s.trim()).filter(s => s);

        const config = {
            ranges,
            targetYear,
            saturdayWorking,
            holidays
        };

        // UI State
        generateBtn.disabled = true;
        resultArea.classList.add('hidden');
        loadingArea.classList.remove('hidden');

        // Send Data
        const formData = new FormData();
        formData.append('file', selectedFile);
        formData.append('config', JSON.stringify(config));

        try {
            const response = await fetch('/api/process', {
                method: 'POST',
                body: formData
            });

            if (!response.ok) {
                const err = await response.json();
                throw new Error(err.error || 'Qayta ishlashda xatolik yuz berdi');
            }

            const data = await response.json();

            // Success
            loadingArea.classList.add('hidden');
            resultArea.style.display = 'block'; // Explicitly show
            resultArea.classList.remove('hidden');
            resultArea.classList.add('fade-in-up');

            // Scroll to result
            resultArea.scrollIntoView({ behavior: 'smooth' });

            const dlLink = document.getElementById('downloadLink') || document.getElementById('download-link');
            if (dlLink) {
                dlLink.href = data.downloadUrl;
                dlLink.textContent = "ZIP Faylni Yuklab Olish";
            } else {
                console.error("Download link element not found!");
                alert("Natija tayyor, lekin yuklab olish tugmasi topilmadi. Sahifani yangilang.");
            }

            // Show warnings if any
            if (data.warnings && data.warnings.length > 0) {
                let warnHtml = '<h4>Ogohlantirishlar:</h4><ul>';
                data.warnings.forEach(w => warnHtml += `<li>${w}</li>`);
                warnHtml += '</ul>';

                // Create or reuse warning container
                let warnDiv = document.getElementById('warnings-area');
                if (!warnDiv) {
                    warnDiv = document.createElement('div');
                    warnDiv.id = 'warnings-area';
                    warnDiv.className = 'warning-message';
                    resultArea.appendChild(warnDiv);
                }
                warnDiv.innerHTML = warnHtml;
                warnDiv.style.display = 'block';
            } else {
                const warnDiv = document.getElementById('warnings-area');
                if (warnDiv) warnDiv.style.display = 'none';
            }

            resultArea.classList.remove('hidden');
        } catch (error) {
            alert('Xatolik: ' + error.message);
        } finally {
            generateBtn.disabled = false;
            loadingArea.classList.add('hidden');
        }
    });
});
