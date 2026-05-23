// DOM Elements
const dropArea = document.getElementById('drop-area');
const fileInput = document.getElementById('file-input');
const uploadSection = document.getElementById('upload-section');
const workspace = document.getElementById('workspace');
const addMoreFilesBtn = document.getElementById('add-more-files-btn');
const uploadedFilesList = document.getElementById('uploaded-files-list');

// Column Selection & Reorder
const columnsList = document.getElementById('columns-list');
const selectAllBtn = document.getElementById('select-all-btn');
const deselectAllBtn = document.getElementById('deselect-all-btn');

// Filters & Formats
const filtersContainer = document.getElementById('filters-container');
const addFilterBtn = document.getElementById('add-filter-btn');
const formatsContainer = document.getElementById('formats-container');
const addFormatBtn = document.getElementById('add-format-btn');

// Sort & Duplicates
const sortColumnSelect = document.getElementById('sort-column');
const sortOrderSelect = document.getElementById('sort-order');
const applyDupBtn = document.getElementById('apply-dup-btn');
const cancelDupBtn = document.getElementById('cancel-dup-btn');
const dupStatus = document.getElementById('dup-status');

// Summary
const summaryColumnSelect = document.getElementById('summary-column');
const summaryTypeSelect = document.getElementById('summary-type');
const rsContainer = document.getElementById('rs-container');
const addRsBtn = document.getElementById('add-rs-btn');

// Preview & Export
const previewThead = document.getElementById('preview-thead');
const previewTbody = document.getElementById('preview-tbody');
const emptyPreview = document.getElementById('empty-preview');
const rowCount = document.getElementById('row-count');
const downloadBtn = document.getElementById('download-btn');
const resetBtn = document.getElementById('reset-btn');
const toastEl = document.getElementById('toast');
const processingStatus = document.getElementById('processing-status');

// Scroll Containers
const topScrollContainer = document.getElementById('top-scroll-container');
const topScrollDummy = document.getElementById('top-scroll-dummy');
const tableContainer = document.getElementById('table-container');

// State
let uploadedFiles = []; // {name, data: ArrayBuffer}
let rawHeaders = [];
let rawData = [];
let currentFileName = "customized_data.xlsx";
let filterIdCounter = 0;
let formatIdCounter = 0;
let rsIdCounter = 0;
let isDuplicateRemovalApplied = false;
let globalDataId = 0;
let pinnedRowIds = new Set();
let rawDataHistory = [];
let isSelecting = false;
let startCoords = null; // {r, c}
let endCoords = null;   // {r, c}
let selectedCells = []; // array of td elements

// Initialize
function init() {
    setupEventListeners();
    setupAccordions();
}

function showToast(message) {
    toastEl.textContent = message;
    toastEl.classList.add('show');
    setTimeout(() => toastEl.classList.remove('show'), 3000);
}

// UI Accordions
function setupAccordions() {
    document.querySelectorAll('.accordion-header').forEach(header => {
        header.addEventListener('click', () => {
            const content = header.nextElementSibling;
            content.classList.toggle('hidden');
        });
    });
}

// Event Listeners setup
function setupEventListeners() {
    // Drag & Drop
    ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
        dropArea.addEventListener(eventName, e => { e.preventDefault(); e.stopPropagation(); }, false);
    });
    ['dragenter', 'dragover'].forEach(eventName => dropArea.addEventListener(eventName, () => dropArea.classList.add('drag-over'), false));
    ['dragleave', 'drop'].forEach(eventName => dropArea.addEventListener(eventName, () => dropArea.classList.remove('drag-over'), false));
    dropArea.addEventListener('drop', handleDrop, false);
    fileInput.addEventListener('change', handleFileSelect, false);
    
    addMoreFilesBtn.addEventListener('click', () => fileInput.click());
    
    selectAllBtn.addEventListener('click', () => {
        document.querySelectorAll('.column-checkbox').forEach(cb => cb.checked = true);
        scheduleUpdate();
    });
    
    deselectAllBtn.addEventListener('click', () => {
        document.querySelectorAll('.column-checkbox').forEach(cb => cb.checked = false);
        scheduleUpdate();
    });

    sortColumnSelect.addEventListener('change', scheduleUpdate);
    sortOrderSelect.addEventListener('change', scheduleUpdate);
    
    applyDupBtn.addEventListener('click', () => {
        isDuplicateRemovalApplied = true;
        dupStatus.textContent = "상태: 적용됨 (미리보기 반영완료)";
        dupStatus.style.color = "var(--success)";
        scheduleUpdate();
    });
    
    cancelDupBtn.addEventListener('click', () => {
        isDuplicateRemovalApplied = false;
        dupStatus.textContent = "상태: 미적용";
        dupStatus.style.color = "var(--text-muted)";
        scheduleUpdate();
    });
    
    addFilterBtn.addEventListener('click', createFilterRow);
    addFormatBtn.addEventListener('click', createFormatRow);
    addRsBtn.addEventListener('click', createRowSummaryRule);

    summaryColumnSelect.addEventListener('change', scheduleUpdate);
    summaryTypeSelect.addEventListener('change', () => {
        const countValInput = document.getElementById('summary-count-val');
        if (countValInput) {
            if (summaryTypeSelect.value === 'count') {
                countValInput.classList.remove('hidden');
            } else {
                countValInput.classList.add('hidden');
            }
        }
        scheduleUpdate();
    });
    const summaryCountValInput = document.getElementById('summary-count-val');
    if (summaryCountValInput) {
        summaryCountValInput.addEventListener('input', scheduleUpdate);
    }
    
    downloadBtn.addEventListener('click', downloadExcel);
    resetBtn.addEventListener('click', resetApp);
    
    // Scroll Sync (Top)
    topScrollContainer.addEventListener('scroll', () => {
        tableContainer.scrollLeft = topScrollContainer.scrollLeft;
    });
    tableContainer.addEventListener('scroll', () => {
        topScrollContainer.scrollLeft = tableContainer.scrollLeft;
    });
    
    // Editable cell sync
    previewTbody.addEventListener('input', handleCellEdit);

    // 마우스 드래그를 통한 셀 범위 선택 (Event Delegation)
    tableContainer.addEventListener('mousedown', (e) => {
        const td = e.target.closest('td');
        if (!td || td.dataset.colIdx === undefined) return;
        if (e.button !== 0) return; // 좌클릭만 허용
        
        isSelecting = true;
        const table = document.getElementById('preview-table');
        if (table) table.classList.add('selecting');
        
        startCoords = { r: parseInt(td.dataset.rowIdx), c: parseInt(td.dataset.colIdx) };
        endCoords = { ...startCoords };
        highlightSelection();
    });
    
    tableContainer.addEventListener('mouseover', (e) => {
        if (!isSelecting) return;
        const td = e.target.closest('td');
        if (!td || td.dataset.colIdx === undefined) return;
        
        const currentR = parseInt(td.dataset.rowIdx);
        const currentC = parseInt(td.dataset.colIdx);
        
        if (endCoords.r !== currentR || endCoords.c !== currentC) {
            endCoords = { r: currentR, c: currentC };
            highlightSelection();
        }
    });
    
    window.addEventListener('mouseup', () => {
        if (isSelecting) {
            isSelecting = false;
            const table = document.getElementById('preview-table');
            if (table) table.classList.remove('selecting');
        }
    });
    
    // 복사 이벤트 처리 (Ctrl + C)
    document.addEventListener('copy', (e) => {
        if (selectedCells.length === 0) return;
        e.preventDefault();
        
        const cellsByRow = {};
        selectedCells.forEach(td => {
            const r = td.dataset.rowIdx;
            if (!cellsByRow[r]) cellsByRow[r] = [];
            cellsByRow[r].push(td);
        });
        
        const textRows = Object.keys(cellsByRow).sort((a,b)=>a-b).map(r => {
            const rowCells = cellsByRow[r].sort((a,b) => a.dataset.colIdx - b.dataset.colIdx);
            return rowCells.map(td => td.textContent).join('\t');
        });
        
        const copyText = textRows.join('\n');
        e.clipboardData.setData('text/plain', copyText);
        showToast("선택한 영역이 클립보드에 복사되었습니다.");
    });
    
    // 실행 취소 (되돌리기) 버튼
    const undoBtn = document.getElementById('undo-btn');
    if (undoBtn) {
        undoBtn.addEventListener('click', undo);
    }
    
    // 키보드 실행 취소 단축키 (Ctrl + Z)
    window.addEventListener('keydown', (e) => {
        if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'z') {
            const activeEl = document.activeElement;
            if (activeEl && activeEl.tagName === 'TD' && activeEl.contentEditable === "true") {
                activeEl.blur(); // 포커스를 해제하여 현재 값 저장 유도 후 실행 취소 진행
            }
            e.preventDefault();
            undo();
        }
    });
    
    // 셀 포커스 진입 시 변경 전 상태를 히스토리에 기록 ( keystone 단위 기록 방지 )
    previewTbody.addEventListener('focusin', (e) => {
        if (e.target.tagName === 'TD' && e.target.contentEditable === "true") {
            pushHistory();
        }
    });

    // 복사한 셀 모양대로 붙여넣기 (Ctrl + V)
    tableContainer.addEventListener('paste', (e) => {
        const td = e.target.closest('td');
        if (!td || td.dataset.colIdx === undefined) return;
        
        e.preventDefault();
        
        const clipboardData = e.clipboardData || window.clipboardData;
        const pastedText = clipboardData.getData('text/plain');
        if (!pastedText) return;
        
        const rowsText = pastedText.split(/\r?\n/);
        if (rowsText.length > 1 && rowsText[rowsText.length - 1] === "") {
            rowsText.pop();
        }
        
        const gridData = rowsText.map(line => line.split('\t'));
        if (gridData.length === 0) return;
        
        const startR = parseInt(td.dataset.rowIdx);
        const startC = parseInt(td.dataset.colIdx);
        
        pushHistory();
        
        const finalHeaders = [];
        document.querySelectorAll('.column-item').forEach(item => {
            if (item.querySelector('.column-checkbox').checked) finalHeaders.push(item.querySelector('.column-checkbox').value);
        });
        const rowSummaries = getActiveRowSummaries();
        rowSummaries.forEach(rs => {
            if (!finalHeaders.includes(rs.name)) finalHeaders.push(rs.name);
        });
        
        const renderedRows = document.querySelectorAll('#preview-tbody tr:not(.summary-row)');
        
        let updateCount = 0;
        gridData.forEach((rowVals, rOffset) => {
            const targetR = startR + rOffset;
            const tr = renderedRows[targetR];
            if (!tr || !tr.dataset.id) return;
            
            const idStr = String(tr.dataset.id);
            let dbRow;
            if (idStr.startsWith('empty-')) {
                const newId = globalDataId++;
                dbRow = { _id: newId };
                rawHeaders.forEach(h => {
                    dbRow[h] = "";
                });
                rawData.push(dbRow);
                tr.dataset.id = newId; // Update DOM ID immediately for focus restoration mapping
            } else {
                const rowId = parseInt(tr.dataset.id);
                dbRow = rawData.find(r => r._id === rowId);
            }
            if (!dbRow) return;
            
            rowVals.forEach((val, cOffset) => {
                const targetC = startC + cOffset;
                const colName = finalHeaders[targetC];
                const isCalculatedCol = rowSummaries.some(rs => rs.name === colName);
                if (colName && !isCalculatedCol) {
                    dbRow[colName] = val;
                    updateCount++;
                }
            });
        });
        
        if (updateCount > 0) {
            showToast(`${updateCount}개 셀에 붙여넣었습니다.`);
            scheduleUpdate();
        }
    });
    
    // 선택 행 삭제 버튼 이벤트 처리 (체크박스 선택 행 또는 드래그 선택 영역 행 일괄 삭제)
    const deleteSelectedRowsBtn = document.getElementById('delete-selected-rows-btn');
    if (deleteSelectedRowsBtn) {
        deleteSelectedRowsBtn.addEventListener('click', () => {
            const idsToDelete = new Set();
            
            // 1. 드래그 범위 선택된 셀들의 행 추가
            selectedCells.forEach(td => {
                const tr = td.parentElement;
                if (tr && tr.dataset.id && !tr.dataset.id.startsWith('empty-')) {
                    idsToDelete.add(parseInt(tr.dataset.id));
                }
            });
            
            // 2. 📌 체크박스로 고정/선택된 행들 추가
            pinnedRowIds.forEach(id => {
                idsToDelete.add(id);
            });
            
            if (idsToDelete.size === 0) {
                showToast("삭제할 행의 📌 체크박스를 선택하거나 셀 영역을 마우스로 드래그해주세요.");
                return;
            }
            
            pushHistory();
            
            rawData = rawData.filter(row => !idsToDelete.has(row._id));
            
            // 삭제된 행들을 고정 Set에서 제거
            idsToDelete.forEach(id => {
                pinnedRowIds.delete(id);
            });
            
            selectedCells = [];
            startCoords = null;
            endCoords = null;
            
            showToast(`${idsToDelete.size}개의 행을 삭제했습니다.`);
            scheduleUpdate();
        });
    }
}

function highlightSelection() {
    if (!startCoords || !endCoords) return;
    
    document.querySelectorAll('#preview-tbody td').forEach(td => {
        td.classList.remove('selected-cell');
    });
    
    const rStart = Math.min(startCoords.r, endCoords.r);
    const rEnd = Math.max(startCoords.r, endCoords.r);
    const cStart = Math.min(startCoords.c, endCoords.c);
    const cEnd = Math.max(startCoords.c, endCoords.c);
    
    selectedCells = [];
    const rows = document.querySelectorAll('#preview-tbody tr:not(.summary-row)');
    
    for (let r = rStart; r <= rEnd; r++) {
        const tr = rows[r];
        if (!tr) continue;
        for (let c = cStart; c <= cEnd; c++) {
            const td = tr.querySelector(`td[data-col-idx="${c}"]`);
            if (td) {
                td.classList.add('selected-cell');
                selectedCells.push(td);
            }
        }
    }
}

function pushHistory() {
    const snapshot = JSON.stringify(rawData);
    if (rawDataHistory.length > 0 && rawDataHistory[rawDataHistory.length - 1] === snapshot) {
        return;
    }
    rawDataHistory.push(snapshot);
    if (rawDataHistory.length > 30) {
        rawDataHistory.shift();
    }
    updateUndoButtonState();
}

function undo() {
    if (rawDataHistory.length === 0) {
        showToast("되돌릴 작업이 없습니다.");
        return;
    }
    const prevDataStr = rawDataHistory.pop();
    rawData = JSON.parse(prevDataStr);
    showToast("이전 상태로 되돌렸습니다.");
    updateUndoButtonState();
    scheduleUpdate();
}

function updateUndoButtonState() {
    const undoBtn = document.getElementById('undo-btn');
    if (undoBtn) {
        undoBtn.disabled = rawDataHistory.length === 0;
    }
}

// File Handling
async function handleDrop(e) { handleFiles(e.dataTransfer.files); }
async function handleFileSelect(e) { handleFiles(e.target.files); }

async function handleFiles(files) {
    if (!files || files.length === 0) return;
    processingStatus.textContent = "파일 읽는 중...";
    processingStatus.classList.add('loading');
    
    let addedCount = 0;
    for (let i = 0; i < files.length; i++) {
        const file = files[i];
        const fileExt = file.name.split('.').pop().toLowerCase();
        if (['xlsx', 'xls', 'csv'].includes(fileExt)) {
            const buffer = await readFileAsArrayBuffer(file);
            const wb = XLSX.read(buffer, {type: 'array', bookSheets: true});
            const sheetNames = wb.SheetNames || ['Sheet1'];
            uploadedFiles.push({ 
                name: file.name, 
                buffer: buffer,
                sheetNames: sheetNames,
                selectedSheet: sheetNames[0]
            });
            addedCount++;
        }
    }
    
    if (addedCount > 0) {
        if (uploadedFiles.length === 1) {
            currentFileName = uploadedFiles[0].name.replace(/\.[^/.]+$/, "") + "_custom.xlsx";
        } else {
            currentFileName = `merged_${uploadedFiles.length}_files_custom.xlsx`;
        }
        uploadSection.classList.add('hidden');
        workspace.classList.remove('hidden');
        renderFileList();
        processFiles();
    } else {
        showToast("유효한 엑셀 파일이 없습니다.");
        processingStatus.classList.remove('loading');
    }
}

function readFileAsArrayBuffer(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = e => resolve(e.target.result);
        reader.onerror = e => reject(e);
        reader.readAsArrayBuffer(file);
    });
}

function renderFileList() {
    const uploadedFilesList = document.getElementById('uploaded-files-list');
    uploadedFilesList.innerHTML = '';
    uploadedFiles.forEach((f, index) => {
        const div = document.createElement('div');
        div.className = 'file-item';
        div.style.display = 'flex';
        div.style.alignItems = 'center';
        div.style.gap = '0.5rem';
        div.style.padding = '0.2rem 0';
        
        const delBtn = document.createElement('button');
        delBtn.innerHTML = '&times;';
        delBtn.className = 'btn-delete-file';
        delBtn.title = '파일 제거';
        delBtn.style.background = 'none';
        delBtn.style.border = 'none';
        delBtn.style.color = 'var(--danger)';
        delBtn.style.cursor = 'pointer';
        delBtn.style.fontSize = '1.2rem';
        delBtn.onclick = () => {
            uploadedFiles.splice(index, 1);
            if (uploadedFiles.length === 0) resetApp();
            else { renderFileList(); processFiles(); }
        };
        
        const nameSpan = document.createElement('span');
        nameSpan.textContent = f.name;
        nameSpan.style.flex = '1';
        nameSpan.style.overflow = 'hidden';
        nameSpan.style.textOverflow = 'ellipsis';
        nameSpan.style.whiteSpace = 'nowrap';
        
        const sheetSelect = document.createElement('select');
        sheetSelect.className = 'select-input';
        sheetSelect.style.width = 'auto';
        sheetSelect.style.padding = '2px 6px';
        sheetSelect.style.fontSize = '0.8rem';
        f.sheetNames.forEach(s => {
            const opt = document.createElement('option');
            opt.value = s;
            opt.textContent = s;
            sheetSelect.appendChild(opt);
        });
        sheetSelect.value = f.selectedSheet;
        sheetSelect.onchange = (e) => {
            f.selectedSheet = e.target.value;
            processFiles();
        };
        
        div.appendChild(delBtn);
        div.appendChild(nameSpan);
        div.appendChild(sheetSelect);
        uploadedFilesList.appendChild(div);
    });
}

function processFiles() {
    rawHeaders = [];
    rawData = [];
    globalDataId = 0;
    let isFirstFile = true;
    
    for (const f of uploadedFiles) {
        try {
            const wb = XLSX.read(f.buffer, {type: 'array'});
            let wsName = f.selectedSheet;
            if (!wb.Sheets[wsName]) wsName = wb.SheetNames[0];
            const ws = wb.Sheets[wsName];
            
            const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
            
            if (aoa.length > 0) {
                let headers = aoa[0];
                let maxCols = 0;
                aoa.forEach(row => { if (row.length > maxCols) maxCols = row.length; });
                for (let i = 0; i < maxCols; i++) {
                    if (headers[i] == null || String(headers[i]).trim() === '') headers[i] = XLSX.utils.encode_col(i);
                }
                const seen = new Set();
                headers = headers.map((h) => {
                    let name = String(h);
                    if (seen.has(name)) {
                        let count = 1;
                        while (seen.has(`${name}_${count}`)) count++;
                        name = `${name}_${count}`;
                    }
                    seen.add(name);
                    return name;
                });
                
                if (isFirstFile) {
                    rawHeaders = [...headers];
                    if (uploadedFiles.length > 1) {
                        rawHeaders.unshift("[원본 파일명]"); // 여러 파일 병합 시 원본 파일명 열 추가
                    }
                    isFirstFile = false;
                } else {
                    // 두 번째 이후 파일의 헤더 중 새로운 것이 있으면 누적 추가
                    headers.forEach(h => {
                        if (!rawHeaders.includes(h)) {
                            rawHeaders.push(h);
                        }
                    });
                }
                
                const fileData = aoa.slice(1)
                    .filter(row => row.some(cell => String(cell).trim() !== '')) // 내용이 아예 없는 빈 행(공백 행) 완벽 제거
                    .map(row => {
                        const obj = {};
                        if (uploadedFiles.length > 1) {
                            obj["[원본 파일명]"] = f.name; // 각 행에 파일명 삽입
                        }
                        
                        // 현재 파일의 헤더를 기준으로 값 매핑
                        headers.forEach((h, idx) => {
                            obj[h] = row[idx] !== undefined ? row[idx] : "";
                        });
                        
                        obj._id = globalDataId++; // 행 고유 식별자 추가
                        return obj;
                    });
                rawData = rawData.concat(fileData);
            }
        } catch (e) {
            console.error("Error processing file", f.name, e);
        }
    }
    
    showToast(uploadedFiles.length > 1 ? `${uploadedFiles.length}개의 파일을 병합했습니다.` : "파일을 불러왔습니다.");
    renderUIControls();
    scheduleUpdate();
}

function renderUIControls() {
    // 이전 컬럼 순서 및 선택 상태 캡처
    const previousCols = [];
    document.querySelectorAll('.column-item').forEach(item => {
        const cb = item.querySelector('.column-checkbox');
        previousCols.push({ name: cb.value, checked: cb.checked });
    });
    
    columnsList.innerHTML = '';
    
    // 이전 상태와 새로운 헤더 병합 (순서 및 체크 상태 보존)
    const headersToRender = [];
    previousCols.forEach(col => {
        if (rawHeaders.includes(col.name)) {
            headersToRender.push(col);
        }
    });
    rawHeaders.forEach(h => {
        if (!headersToRender.find(c => c.name === h)) {
            headersToRender.push({ name: h, checked: true }); // 새로운 열은 기본으로 체크
        }
    });

    headersToRender.forEach((colObj) => {
        const header = colObj.name;
        const item = document.createElement('div');
        item.className = 'column-item';
        item.draggable = true;
        
        const dragHandle = document.createElement('span');
        dragHandle.className = 'drag-handle';
        dragHandle.innerHTML = '≡';
        
        const checkbox = document.createElement('input');
        checkbox.type = 'checkbox';
        checkbox.className = 'column-checkbox';
        checkbox.value = header;
        checkbox.checked = colObj.checked;
        checkbox.addEventListener('change', scheduleUpdate);
        
        const span = document.createElement('span');
        span.className = 'column-label';
        span.textContent = header;
        
        item.appendChild(dragHandle);
        item.appendChild(checkbox);
        item.appendChild(span);
        columnsList.appendChild(item);
        
        // 전체 영역 클릭 시 체크박스 토글
        item.addEventListener('click', (e) => {
            if (e.target !== checkbox && e.target !== dragHandle) {
                checkbox.checked = !checkbox.checked;
                scheduleUpdate();
            }
        });
        
        item.addEventListener('dragstart', () => { item.classList.add('dragging'); });
        item.addEventListener('dragend', () => {
            item.classList.remove('dragging');
            scheduleUpdate();
        });
    });
    
    columnsList.addEventListener('dragover', e => {
        e.preventDefault();
        const afterElement = getDragAfterElement(columnsList, e.clientY);
        const draggable = document.querySelector('.dragging');
        if (afterElement == null) columnsList.appendChild(draggable);
        else columnsList.insertBefore(draggable, afterElement);
    });

    const populateSelect = (selectEl, defaultText) => {
        const prev = selectEl.value;
        selectEl.innerHTML = `<option value="">${defaultText}</option>`;
        rawHeaders.forEach(h => {
            const opt = document.createElement('option');
            opt.value = h; opt.textContent = h;
            selectEl.appendChild(opt);
        });
        if (rawHeaders.includes(prev)) selectEl.value = prev;
    };
    
    populateSelect(sortColumnSelect, "정렬 안함");
    populateSelect(summaryColumnSelect, "요약할 열 선택 안함");
}

function getDragAfterElement(container, y) {
    const draggableElements = [...container.querySelectorAll('.column-item:not(.dragging)')];
    return draggableElements.reduce((closest, child) => {
        const box = child.getBoundingClientRect();
        const offset = y - box.top - box.height / 2;
        if (offset < 0 && offset > closest.offset) return { offset: offset, element: child };
        else return closest;
    }, { offset: Number.NEGATIVE_INFINITY }).element;
}

// Format Rules
function createFormatRow() {
    const fid = `format-${formatIdCounter++}`;
    const row = document.createElement('div');
    row.className = 'format-item';
    row.id = fid;
    
    let colOptions = rawHeaders.map(h => `<option value="${h}">${h}</option>`).join('');
    row.innerHTML = `
        <button class="format-remove" title="삭제">&times;</button>
        <div class="format-row">
            <div class="format-row-top">
                <select class="select-input format-col">${colOptions}</select>
                <select class="select-input format-op">
                    <option value="trim">공백 제거</option>
                    <option value="date">날짜 통일 (YYYY-MM-DD)</option>
                    <option value="number">숫자 쉼표 추가</option>
                    <option value="prefix">앞에 글자 추가</option>
                    <option value="suffix">뒤에 글자 추가</option>
                    <option value="delete">특정 단어 삭제</option>
                </select>
            </div>
            <input type="text" class="text-input format-val hidden" placeholder="추가/삭제할 단어 입력">
        </div>
    `;
    
    const opSelect = row.querySelector('.format-op');
    const valInput = row.querySelector('.format-val');
    
    opSelect.addEventListener('change', () => {
        if (['prefix', 'suffix', 'delete'].includes(opSelect.value)) {
            valInput.classList.remove('hidden');
        } else {
            valInput.classList.add('hidden');
        }
        scheduleUpdate();
    });
    
    row.querySelector('.format-remove').addEventListener('click', () => {
        row.remove();
        scheduleUpdate();
    });
    
    row.querySelectorAll('select, input').forEach(el => el.addEventListener('input', scheduleUpdate));
    formatsContainer.appendChild(row);
}

function getActiveFormats() {
    const formats = [];
    document.querySelectorAll('.format-item').forEach(item => {
        const col = item.querySelector('.format-col').value;
        const op = item.querySelector('.format-op').value;
        const val = item.querySelector('.format-val').value;
        formats.push({col, op, val});
    });
    return formats;
}

// Filter Rules
function createFilterRow() {
    const filterId = `filter-${filterIdCounter++}`;
    const row = document.createElement('div');
    row.className = 'filter-item';
    row.id = filterId;
    
    let colOptions = `<option value="__all__">[전체 열]</option>` + rawHeaders.map(h => `<option value="${h}">${h}</option>`).join('');
    row.innerHTML = `
        <button class="filter-remove" title="삭제">&times;</button>
        <div class="filter-row">
            <select class="select-input filter-col">${colOptions}</select>
            <select class="select-input filter-op">
                <option value="contains">포함</option>
                <option value="equals">일치</option>
                <option value="greater">이상 (숫자)</option>
                <option value="less">이하 (숫자)</option>
            </select>
            <input type="text" class="text-input filter-val" placeholder="검색어 입력">
        </div>
    `;
    
    row.querySelector('.filter-remove').addEventListener('click', () => { row.remove(); scheduleUpdate(); });
    row.querySelectorAll('select, input').forEach(el => el.addEventListener('input', scheduleUpdate));
    filtersContainer.appendChild(row);
}

function getActiveFilters() {
    const filters = [];
    document.querySelectorAll('.filter-item').forEach(item => {
        const col = item.querySelector('.filter-col').value;
        const op = item.querySelector('.filter-op').value;
        const val = item.querySelector('.filter-val').value.trim();
        if (val) filters.push({col, op, val});
    });
    return filters;
}

// Row Summary Rules (가로 요약)
function createRowSummaryRule() {
    const rsId = `rs-${rsIdCounter++}`;
    const row = document.createElement('div');
    row.className = 'rs-item';
    row.id = rsId;
    
    let checkboxesHtml = rawHeaders.map(h => `
        <label class="rs-col-lbl">
            <input type="checkbox" class="rs-col-cb" value="${h}"> ${h}
        </label>
    `).join('');
    
    row.innerHTML = `
        <button class="rs-remove" title="삭제">&times;</button>
        <div class="rs-row-top">
            <input type="text" class="text-input rs-name" placeholder="새로운 열 이름 (예: 총점)">
            <select class="select-input rs-type">
                <option value="sum">합계</option>
                <option value="avg">평균</option>
            </select>
        </div>
        <p style="font-size: 0.8rem; color: var(--text-muted); margin: 0;">계산에 포함할 열 선택:</p>
        <div class="rs-cols-container">
            ${checkboxesHtml}
        </div>
    `;
    
    row.querySelector('.rs-remove').addEventListener('click', () => { row.remove(); scheduleUpdate(); });
    row.querySelectorAll('select, input').forEach(el => el.addEventListener('input', scheduleUpdate));
    rsContainer.appendChild(row);
}

function getActiveRowSummaries() {
    const rules = [];
    document.querySelectorAll('.rs-item').forEach(item => {
        const name = item.querySelector('.rs-name').value.trim();
        const type = item.querySelector('.rs-type').value;
        const cols = [];
        item.querySelectorAll('.rs-col-cb:checked').forEach(cb => cols.push(cb.value));
        if (name && cols.length > 0) {
            rules.push({ name, type, cols });
        }
    });
    return rules;
}

// Cell Edit Handler (Preview editing updates Raw Data)
function handleCellEdit(e) {
    if (e.target.tagName !== 'TD') return;
    const td = e.target;
    const tr = td.parentElement;
    if (tr.classList.contains('summary-row')) return;
    
    const colName = td.dataset.col;
    let newVal = td.innerText;
    
    const idStr = String(tr.dataset.id);
    if (idStr.startsWith('empty-')) {
        pushHistory();
        
        const newId = globalDataId++;
        const newRow = { _id: newId };
        rawHeaders.forEach(h => {
            newRow[h] = "";
        });
        newRow[colName] = newVal;
        rawData.push(newRow);
        
        tr.dataset.id = newId;
        
        // 빈 행이 활성화되었으므로 체크박스 및 삭제 버튼 활성화
        const cb = tr.querySelector('.row-pin-checkbox');
        if (cb) {
            cb.disabled = false;
            cb.addEventListener('click', (e) => e.stopPropagation());
            cb.addEventListener('change', () => {
                if (cb.checked) pinnedRowIds.add(newId);
                else pinnedRowIds.delete(newId);
                scheduleUpdate();
            });
        }
        
        const btnDel = tr.querySelector('.row-delete-btn');
        if (btnDel) {
            btnDel.disabled = false;
            btnDel.addEventListener('click', (e) => {
                e.stopPropagation();
                pushHistory();
                rawData = rawData.filter(r => r._id !== newId);
                showToast("행을 삭제했습니다.");
                scheduleUpdate();
            });
        }
    } else {
        const id = parseInt(tr.dataset.id);
        const row = rawData.find(r => r._id === id);
        if (row) {
            row[colName] = newVal;
        }
    }
}

// Data Processing Pipeline
let updateTimeout = null;
function scheduleUpdate() {
    processingStatus.textContent = "처리 중...";
    processingStatus.classList.add('loading');
    clearTimeout(updateTimeout);
    updateTimeout = setTimeout(() => {
        updatePreview();
        processingStatus.classList.remove('loading');
    }, 100);
}

function getProcessedData() {
    let workingData = rawData.map(row => {
        // 모든 rawHeaders 열에 대해 값이 없으면 빈 문자열("") 보장
        let obj = { _id: row._id };
        rawHeaders.forEach(h => {
            obj[h] = row[h] !== undefined ? row[h] : "";
        });
        return obj;
    });

    const filters = getActiveFilters();
    if (filters.length > 0) {
        workingData = workingData.filter(row => {
            // Pinned rows always bypass filtering conditions
            if (pinnedRowIds.has(row._id)) return true;

            return filters.every(f => {
                const searchStr = f.val.toLowerCase();
                
                // If filtering on "All Columns"
                if (f.col === '__all__') {
                    return rawHeaders.some(h => {
                        const cellVal = row[h];
                        if (cellVal === null || cellVal === undefined) return false;
                        const cellStr = String(cellVal).toLowerCase();
                        
                        if (f.op === 'contains') return cellStr.includes(searchStr);
                        if (f.op === 'equals') return cellStr === searchStr;
                        if (f.op === 'greater') return !isNaN(cellVal) && cellVal !== "" && Number(cellVal) >= Number(f.val);
                        if (f.op === 'less') return !isNaN(cellVal) && cellVal !== "" && Number(cellVal) <= Number(f.val);
                        return false;
                    });
                }
                
                const cellVal = row[f.col];
                if (cellVal === null || cellVal === undefined) return false;
                const cellStr = String(cellVal).toLowerCase();
                
                if (f.op === 'contains') return cellStr.includes(searchStr);
                if (f.op === 'equals') return cellStr === searchStr;
                if (f.op === 'greater') return !isNaN(cellVal) && cellVal !== "" && Number(cellVal) >= Number(f.val);
                if (f.op === 'less') return !isNaN(cellVal) && cellVal !== "" && Number(cellVal) <= Number(f.val);
                return true;
            });
        });
    }

    const formats = getActiveFormats();
    if (formats.length > 0) {
        workingData = workingData.map(row => {
            let newRow = {...row};
            formats.forEach(f => {
                let val = newRow[f.col];
                if (val !== null && val !== undefined && val !== "") {
                    let strVal = String(val);
                    if (f.op === 'trim') val = strVal.trim();
                    else if (f.op === 'date') {
                        if (/^\d{4}[\./-]\d{1,2}[\./-]\d{1,2}$/.test(strVal)) {
                            strVal = strVal.replace(/[\./]/g, '-');
                            let parts = strVal.split('-');
                            if(parts[1].length === 1) parts[1] = '0' + parts[1];
                            if(parts[2].length === 1) parts[2] = '0' + parts[2];
                            val = parts.join('-');
                        }
                    } else if (f.op === 'number' && !isNaN(strVal.replace(/,/g,''))) {
                        val = Number(strVal.replace(/,/g,'')).toLocaleString('ko-KR');
                    } else if (f.op === 'prefix') {
                        if (!strVal.startsWith(f.val)) val = f.val + strVal;
                    } else if (f.op === 'suffix') {
                        if (!strVal.endsWith(f.val)) val = strVal + f.val;
                    } else if (f.op === 'delete' && f.val) {
                        val = strVal.split(f.val).join('');
                    }
                    newRow[f.col] = val;
                }
            });
            return newRow;
        });
    }

    const sortCol = sortColumnSelect.value;
    if (sortCol) {
        const order = sortOrderSelect.value === 'asc' ? 1 : -1;
        workingData.sort((a, b) => {
            let valA = a[sortCol];
            let valB = b[sortCol];
            
            if (typeof valA === 'string') valA = valA.replace(/,/g, '');
            if (typeof valB === 'string') valB = valB.replace(/,/g, '');
            
            if (valA === "" || valA == null) return 1;
            if (valB === "" || valB == null) return -1;
            if (!isNaN(valA) && !isNaN(valB)) return (Number(valA) - Number(valB)) * order;
            return String(valA).localeCompare(String(valB)) * order;
        });
    }

    const finalHeaders = [];
    document.querySelectorAll('.column-item').forEach(item => {
        if (item.querySelector('.column-checkbox').checked) finalHeaders.push(item.querySelector('.column-checkbox').value);
    });

    // Add Row Summary Headers
    const rowSummaries = getActiveRowSummaries();
    rowSummaries.forEach(rs => {
        if (!finalHeaders.includes(rs.name)) finalHeaders.push(rs.name);
    });

    if (isDuplicateRemovalApplied && finalHeaders.length > 0) {
        const seen = new Set();
        workingData = workingData.filter(row => {
            const key = finalHeaders.map(h => row[h]).join('|||');
            if (seen.has(key)) return false;
            seen.add(key);
            return true;
        });
    }

    let finalDataArrays = workingData.map(row => {
        let arr = finalHeaders.map(h => {
            const rs = rowSummaries.find(r => r.name === h);
            if (rs) {
                let total = 0, count = 0;
                rs.cols.forEach(c => {
                    let val = row[c];
                    if (typeof val === 'string') val = val.replace(/,/g, '');
                    if (!isNaN(val) && val !== "" && val !== null) { total += Number(val); count++; }
                });
                if (rs.type === 'sum') return total;
                if (rs.type === 'avg') return count === 0 ? 0 : Math.round((total / count) * 100) / 100;
            }
            return row[h];
        });
        arr._id = row._id; // pass _id down to render
        return arr;
    });

    const sumCol = summaryColumnSelect.value;
    if (sumCol && finalHeaders.includes(sumCol) && finalDataArrays.length > 0) {
        const sumIndex = finalHeaders.indexOf(sumCol);
        const sumType = summaryTypeSelect.value;
        const countFilterInput = document.getElementById('summary-count-val');
        const countFilterVal = (sumType === 'count' && countFilterInput) ? countFilterInput.value.trim() : "";
        
        let total = 0, count = 0;
        
        finalDataArrays.forEach(row => {
            let val = row[sumIndex];
            if (typeof val === 'string') val = val.replace(/,/g, '');
            if (!isNaN(val) && val !== "" && val !== null) { total += Number(val); }
            
            if (val !== "" && val !== null) {
                if (sumType === 'count' && countFilterVal !== "") {
                    if (String(row[sumIndex]).trim().toLowerCase() === countFilterVal.toLowerCase()) {
                        count++;
                    }
                } else {
                    count++;
                }
            }
        });
        
        let result = 0;
        let summaryLabel = "";
        let formattedResult = "";
        
        if (sumType === 'sum') {
            result = total;
            summaryLabel = "[총 합계]";
            formattedResult = result.toLocaleString('ko-KR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
        } else if (sumType === 'avg') {
            result = count === 0 ? 0 : total / count;
            summaryLabel = "[평 균]";
            formattedResult = result.toLocaleString('ko-KR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
        } else if (sumType === 'count') {
            result = count;
            summaryLabel = countFilterVal ? `[건수: ${countFilterVal}]` : "[데이터 건수]";
            formattedResult = result.toLocaleString('ko-KR');
        }

        const summaryRow = new Array(finalHeaders.length).fill("");
        summaryRow[0] = summaryLabel;
        summaryRow[sumIndex] = formattedResult;
        finalDataArrays.push(summaryRow);
    }

    return { headers: finalHeaders, data: finalDataArrays };
}

function updatePreview() {
    // 1. Save focus and caret selection
    const activeEl = document.activeElement;
    let savedFocus = null;
    if (activeEl && activeEl.tagName === 'TD' && activeEl.closest('#preview-tbody')) {
        const tr = activeEl.parentElement;
        const colName = activeEl.dataset.col;
        const rowId = tr.dataset.id; // Could be a number string or 'empty-X'
        
        let selectionOffset = 0;
        try {
            const selection = window.getSelection();
            if (selection.rangeCount > 0) {
                const range = selection.getRangeAt(0);
                const preCaretRange = range.cloneRange();
                preCaretRange.selectNodeContents(activeEl);
                preCaretRange.setEnd(range.endContainer, range.endOffset);
                selectionOffset = preCaretRange.toString().length;
            }
        } catch (e) {
            console.error("Error saving caret position:", e);
        }
        
        savedFocus = { rowId, colName, selectionOffset };
    }

    const { headers, data } = getProcessedData();
    
    previewThead.innerHTML = '';
    previewTbody.innerHTML = '';
    
    if (headers.length === 0) {
        emptyPreview.classList.remove('hidden');
        rowCount.textContent = '0 rows';
        return;
    }
    
    emptyPreview.classList.add('hidden');
    rowCount.textContent = `${data.length.toLocaleString()} rows`;

    const displayData = data.slice(0, 500);
    
    const trHead = document.createElement('tr');
    
    // 📌 고정 열 헤더 추가
    const thPin = document.createElement('th');
    thPin.innerHTML = '📌';
    thPin.style.width = '45px';
    thPin.style.textAlign = 'center';
    trHead.appendChild(thPin);
    
    // 🗑️ 삭제 열 헤더 추가
    const thDel = document.createElement('th');
    thDel.innerHTML = '🗑️';
    thDel.style.width = '45px';
    thDel.style.textAlign = 'center';
    trHead.appendChild(thDel);
    
    headers.forEach(header => {
        const th = document.createElement('th');
        th.textContent = header;
        trHead.appendChild(th);
    });
    previewThead.appendChild(trHead);
    
    displayData.forEach((row, idx) => {
        const tr = document.createElement('tr');
        const isSummaryRow = (idx === data.length - 1 && data.length <= 500 && (row[0] === "[총 합계]" || row[0] === "[평 균]" || row[0] === "[데이터 건수]" || (typeof row[0] === 'string' && row[0].startsWith("[건수:"))));
        
        const isPinned = !isSummaryRow && pinnedRowIds.has(row._id);
        
        if (isSummaryRow) {
            tr.className = 'summary-row';
        } else {
            tr.dataset.id = row._id; // Attach rawData ID
            if (isPinned) {
                tr.classList.add('pinned-row');
            }
        }
        
        // 📌 고정 열 체크박스 추가
        const tdPin = document.createElement('td');
        tdPin.style.textAlign = 'center';
        if (!isSummaryRow) {
            const cb = document.createElement('input');
            cb.type = 'checkbox';
            cb.className = 'row-pin-checkbox';
            cb.checked = isPinned;
            cb.addEventListener('click', (e) => {
                e.stopPropagation(); // 셀 수정 에디터 등의 이벤트를 막기 위해 정지
            });
            cb.addEventListener('change', () => {
                if (cb.checked) {
                    pinnedRowIds.add(row._id);
                } else {
                    pinnedRowIds.delete(row._id);
                }
                scheduleUpdate();
            });
            tdPin.appendChild(cb);
        }
        tr.appendChild(tdPin);
        
        // 🗑️ 삭제 열 버튼 추가
        const tdDel = document.createElement('td');
        tdDel.style.textAlign = 'center';
        if (!isSummaryRow) {
            const btnDel = document.createElement('button');
            btnDel.innerHTML = '&times;';
            btnDel.className = 'row-delete-btn';
            btnDel.title = '이 행 삭제';
            btnDel.style.background = 'none';
            btnDel.style.border = 'none';
            btnDel.style.color = 'var(--danger)';
            btnDel.style.cursor = 'pointer';
            btnDel.style.fontSize = '1.2rem';
            btnDel.style.lineHeight = '1';
            btnDel.addEventListener('click', (e) => {
                e.stopPropagation();
                pushHistory();
                rawData = rawData.filter(r => r._id !== row._id);
                showToast("행을 삭제했습니다.");
                scheduleUpdate();
            });
            tdDel.appendChild(btnDel);
        }
        tr.appendChild(tdDel);
        
        row.forEach((cell, cIdx) => {
            const td = document.createElement('td');
            td.textContent = cell;
            td.title = cell;
            if (!isSummaryRow) {
                td.dataset.col = headers[cIdx]; // Attach column name for editing
                td.dataset.rowIdx = idx;       // Row index for selection coordinate
                td.dataset.colIdx = cIdx;      // Col index for selection coordinate
                td.contentEditable = "true"; // Make editable!
            }
            tr.appendChild(td);
        });
        previewTbody.appendChild(tr);
    });
    
    // 자료가 없더라도 500개 행까지 빈 행으로 채워넣기
    if (displayData.length < 500) {
        const targetCount = 500;
        const currentLength = displayData.length;
        const hasSummary = (currentLength > 0 && (displayData[currentLength - 1][0] === "[총 합계]" || displayData[currentLength - 1][0] === "[평 균]" || displayData[currentLength - 1][0] === "[데이터 건수]" || (typeof displayData[currentLength - 1][0] === 'string' && displayData[currentLength - 1][0].startsWith("[건수:"))));
        const renderCount = targetCount - (hasSummary ? 1 : 0);
        
        for (let idx = currentLength - (hasSummary ? 1 : 0); idx < renderCount; idx++) {
            const tr = document.createElement('tr');
            tr.dataset.id = `empty-${idx}`;
            
            // 📌 고정 열 체크박스 추가 (빈 행용 - 비활성화)
            const tdPin = document.createElement('td');
            tdPin.style.textAlign = 'center';
            const cb = document.createElement('input');
            cb.type = 'checkbox';
            cb.className = 'row-pin-checkbox';
            cb.disabled = true;
            tdPin.appendChild(cb);
            tr.appendChild(tdPin);
            
            // 🗑️ 삭제 열 버튼 추가 (빈 행용 - 비활성화)
            const tdDel = document.createElement('td');
            tdDel.style.textAlign = 'center';
            const btnDel = document.createElement('button');
            btnDel.innerHTML = '&times;';
            btnDel.className = 'row-delete-btn';
            btnDel.style.background = 'none';
            btnDel.style.border = 'none';
            btnDel.style.color = 'var(--danger)';
            btnDel.style.opacity = '0.3';
            btnDel.disabled = true;
            tdDel.appendChild(btnDel);
            tr.appendChild(tdDel);
            
            headers.forEach((h, cIdx) => {
                const td = document.createElement('td');
                td.textContent = "";
                td.dataset.col = h;
                td.dataset.rowIdx = idx;
                td.dataset.colIdx = cIdx;
                td.contentEditable = "true";
                tr.appendChild(td);
            });
            
            if (hasSummary && previewTbody.lastChild) {
                previewTbody.insertBefore(tr, previewTbody.lastChild);
            } else {
                previewTbody.appendChild(tr);
            }
        }
    }
    
    if (data.length > 500) {
        const trInfo = document.createElement('tr');
        const tdInfo = document.createElement('td');
        tdInfo.colSpan = headers.length + 2; // 📌 고정 열 및 삭제 열 포함하여 2 증가
        tdInfo.style.textAlign = 'center';
        tdInfo.style.color = 'var(--text-muted)';
        tdInfo.style.fontStyle = 'italic';
        tdInfo.style.backgroundColor = 'rgba(15, 23, 42, 0.3)';
        tdInfo.textContent = `... 두 번째 파일 등 이어지는 데이터 포함, 전체 ${data.length.toLocaleString()}행 중 맨 위 500행만 미리보기 중입니다 ...`;
        trInfo.appendChild(tdInfo);
        previewTbody.appendChild(trInfo);
        
        const lastDataRow = data[data.length-1];
        if (lastDataRow[0] === "[총 합계]" || lastDataRow[0] === "[평 균]" || lastDataRow[0] === "[데이터 건수]" || (typeof lastDataRow[0] === 'string' && lastDataRow[0].startsWith("[건수:"))) {
            const trSum = document.createElement('tr');
            trSum.className = 'summary-row';
            
            // 📌 고정 열 정렬을 위한 빈 셀 추가
            const tdPinEmpty = document.createElement('td');
            trSum.appendChild(tdPinEmpty);
            
            // 🗑️ 삭제 열 정렬을 위한 빈 셀 추가
            const tdDelEmpty = document.createElement('td');
            trSum.appendChild(tdDelEmpty);
            
            lastDataRow.forEach(cell => {
                const td = document.createElement('td');
                td.textContent = cell;
                trSum.appendChild(td);
            });
            previewTbody.appendChild(trSum);
        }
    }
    
    // Update scroll dummies
    setTimeout(() => {
        const table = document.getElementById('preview-table');
        if (table) {
            topScrollDummy.style.width = table.offsetWidth + 'px';
        }
    }, 50);

    // 2. Restore cell range selection highlight
    highlightSelection();

    // 3. Restore focused cell and caret position
    if (savedFocus) {
        const targetTr = Array.from(previewTbody.querySelectorAll('tr')).find(tr => String(tr.dataset.id) === String(savedFocus.rowId));
        if (targetTr) {
            const targetTd = targetTr.querySelector(`td[data-col="${savedFocus.colName}"]`);
            if (targetTd) {
                targetTd.focus();
                
                // Restore selection caret
                try {
                    const range = document.createRange();
                    const selection = window.getSelection();
                    
                    let charCount = 0;
                    let nodeToSet = null;
                    let offsetInNode = 0;
                    
                    function traverseNodes(node) {
                        if (node.nodeType === Node.TEXT_NODE) {
                            const nextCount = charCount + node.length;
                            if (savedFocus.selectionOffset >= charCount && savedFocus.selectionOffset <= nextCount) {
                                nodeToSet = node;
                                offsetInNode = savedFocus.selectionOffset - charCount;
                                return true;
                            }
                            charCount = nextCount;
                        } else {
                            for (let child of node.childNodes) {
                                if (traverseNodes(child)) return true;
                            }
                        }
                        return false;
                    }
                    
                    traverseNodes(targetTd);
                    
                    if (nodeToSet) {
                        range.setStart(nodeToSet, offsetInNode);
                        range.collapse(true);
                        selection.removeAllRanges();
                        selection.addRange(range);
                    } else {
                        range.selectNodeContents(targetTd);
                        range.collapse(false);
                        selection.removeAllRanges();
                        selection.addRange(range);
                    }
                } catch (err) {
                    console.error("Error restoring caret selection:", err);
                }
            }
        }
    }
}

function getCellWidth(value) {
    if (value === null || value === undefined || value === "") return 0;
    const str = String(value);
    let width = 0;
    for (let i = 0; i < str.length; i++) {
        width += (str.charCodeAt(i) > 128) ? 1.5 : 1;
    }
    return width;
}

function downloadExcel() {
    const { headers, data } = getProcessedData();
    if (headers.length === 0) return showToast("다운로드할 데이터가 없습니다.");

    // Remove _id property from data arrays before export
    const exportData = data.map(row => [...row]);
    
    const aoa = [headers, ...exportData];
    const newWb = XLSX.utils.book_new();
    const newWs = XLSX.utils.aoa_to_sheet(aoa);
    
    const colWidths = headers.map((header, colIndex) => {
        let maxWidth = getCellWidth(header);
        exportData.forEach(row => {
            const cellWidth = getCellWidth(row[colIndex]);
            if (cellWidth > maxWidth) maxWidth = cellWidth;
        });
        return { wch: Math.min(Math.max(Math.ceil(maxWidth) + 1, 3), 100) };
    });
    newWs['!cols'] = colWidths;
    
    XLSX.utils.book_append_sheet(newWb, newWs, "Result");
    XLSX.writeFile(newWb, currentFileName);
    showToast("파일이 다운로드 되었습니다.");
}

function resetApp() {
    uploadedFiles = [];
    rawHeaders = [];
    rawData = [];
    filterIdCounter = 0;
    formatIdCounter = 0;
    rsIdCounter = 0;
    globalDataId = 0;
    isDuplicateRemovalApplied = false;
    pinnedRowIds.clear();
    rawDataHistory = [];
    updateUndoButtonState();
    isSelecting = false;
    startCoords = null;
    endCoords = null;
    selectedCells = [];
    fileInput.value = "";
    filtersContainer.innerHTML = '';
    formatsContainer.innerHTML = '';
    rsContainer.innerHTML = '';
    
    if(dupStatus) {
        dupStatus.textContent = "상태: 미적용";
        dupStatus.style.color = "var(--text-muted)";
    }
    
    workspace.classList.add('hidden');
    uploadSection.classList.remove('hidden');
}

// App Start
init();
