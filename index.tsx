/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

// --- Type Definitions ---
interface StationData {
    "基地台名稱": string;
    "轄區": string;
    "型式": string;
    "高度(米)": number;
    "上期維護": string;
    "本次檢查": string;
    "檢查與改善": string;
    "備 註": string;
}

interface DistrictStats {
    [district: string]: { countA: number; countB: number; countC: number; }
}

// --- Global State ---
let allData: StationData[] = [];
let specialStationsData: StationData[] = [];
let allFilteredStationsData: StationData[] = [];
let currentYearlyStats: DistrictStats = {};

const defaultKeywords = ['更換', '鬆脫', '除鏽', '漏水', '破裂', '裂縫'];
let currentKeywords = new Set<string>(defaultKeywords);

/**
 * Initializes the application, gets references to DOM elements,
 * and attaches event listeners.
 */
document.addEventListener('DOMContentLoaded', () => {
    // --- Data Input ---
    const fileInput = document.getElementById('json-upload') as HTMLInputElement;
    const jsonPasteArea = document.getElementById('json-paste-area') as HTMLTextAreaElement;
    const pasteJsonBtn = document.getElementById('paste-json-btn') as HTMLButtonElement;
    const clearJsonBtn = document.getElementById('clear-json-btn') as HTMLButtonElement;

    // --- Controls & Actions ---
    const analyzeBtn = document.getElementById('analyze-btn') as HTMLButtonElement;
    const dateFilterRadios = document.querySelectorAll<HTMLInputElement>('input[name="date-filter"]');
    const customDateRange = document.getElementById('custom-date-range') as HTMLDivElement;
    const addKeywordBtn = document.getElementById('add-keyword-btn') as HTMLButtonElement;
    const keywordListContainer = document.getElementById('keyword-list-container') as HTMLDivElement;
    const resultsContainer = document.getElementById('results') as HTMLElement;
    const clearDistrictBtn = document.getElementById('clear-district-filter-btn') as HTMLButtonElement;

    // --- Event Listeners ---
    fileInput?.addEventListener('change', handleFileUpload);
    jsonPasteArea?.addEventListener('input', () => processJsonData(jsonPasteArea.value));
    pasteJsonBtn?.addEventListener('click', handlePasteFromJsonButton);
    clearJsonBtn?.addEventListener('click', handleClearJsonTextarea);
    
    analyzeBtn?.addEventListener('click', handleAnalysis);
    addKeywordBtn?.addEventListener('click', handleAddKeyword);
    clearDistrictBtn?.addEventListener('click', handleClearDistrictSelection);

    keywordListContainer.addEventListener('click', (event) => {
        const target = event.target as HTMLElement;
        if (target.classList.contains('remove-keyword')) {
            const keyword = target.dataset.keyword;
            if (keyword) {
                handleRemoveKeyword(keyword);
            }
        }
    });

    resultsContainer.addEventListener('click', (event) => {
        const target = event.target as HTMLButtonElement;
        if (target.classList.contains('export-btn')) {
            const format = target.dataset.format as 'csv' | 'json';
            const dataType = target.dataset.target as 'special' | 'all';
            handleExport(dataType, format);
        }
    });

    resultsContainer.addEventListener('input', (event) => {
        const target = event.target as HTMLInputElement;
        if (target.classList.contains('target-input')) {
            handleTargetInputChange();
        }
    });

    dateFilterRadios.forEach(radio => {
        radio.addEventListener('change', () => {
            customDateRange.style.display = radio.value === 'custom' ? 'block' : 'none';
        });
    });

    renderKeywords(); // Initial render of default keywords
});

/**
 * Handles the file upload event. Reads the file content and passes it to the processor.
 * @param event The file input change event.
 */
async function handleFileUpload(event: Event) {
    const target = event.target as HTMLInputElement;
    if (target.files && target.files.length > 0) {
        const file = target.files[0];
        const fileContent = await file.text();
        processJsonData(fileContent, true);
    }
}

/**
 * Handles pasting from the clipboard into the textarea.
 */
async function handlePasteFromJsonButton() {
    try {
        const text = await navigator.clipboard.readText();
        const jsonPasteArea = document.getElementById('json-paste-area') as HTMLTextAreaElement;
        if (jsonPasteArea) {
            jsonPasteArea.value = text;
            processJsonData(text);
        }
    } catch (err) {
        console.error('Failed to read clipboard contents: ', err);
        showMessage('無法讀取剪貼簿內容。', 'error');
    }
}

/**
 * Clears the JSON textarea and resets associated state.
 */
function handleClearJsonTextarea() {
    const jsonPasteArea = document.getElementById('json-paste-area') as HTMLTextAreaElement;
    if (jsonPasteArea) {
        jsonPasteArea.value = '';
        processJsonData(''); // Trigger reset
    }
}


/**
 * Central function to parse, validate, and load JSON data from a string.
 * @param jsonString The string content to parse.
 * @param fromFile Indicates if the source was a file upload, to clear the other input.
 */
function processJsonData(jsonString: string, fromFile: boolean = false) {
    const analyzeBtn = document.getElementById('analyze-btn') as HTMLButtonElement;
    const districtFilter = document.getElementById('district-filter') as HTMLSelectElement;

    // Clear the other input to avoid confusion
    if (fromFile) {
        const jsonPasteArea = document.getElementById('json-paste-area') as HTMLTextAreaElement;
        if(jsonPasteArea) jsonPasteArea.value = '';
    } else {
        const fileInput = document.getElementById('json-upload') as HTMLInputElement;
        if(fileInput) fileInput.value = '';
    }

    if (!jsonString.trim()) {
        allData = [];
        districtFilter.innerHTML = '';
        districtFilter.disabled = true;
        analyzeBtn.disabled = true;
        return;
    }

    try {
        const parsedData = JSON.parse(jsonString);
        if (!Array.isArray(parsedData) || parsedData.length === 0) {
            throw new Error("JSON is not a valid array or is empty.");
        }
        const firstItem = parsedData[0];
        if (!firstItem || !('基地台名稱' in firstItem && '轄區' in firstItem && '本次檢查' in firstItem)) {
            throw new Error("JSON structure is incorrect. Missing required fields.");
        }

        allData = parsedData;
        populateDistrictFilter(allData);
        districtFilter.disabled = false;
        analyzeBtn.disabled = false;
        showMessage(`成功載入 ${allData.length} 筆資料。`, 'success');
    } catch (error) {
        console.error("JSON processing error:", error);
        showMessage(`資料解析失敗: ${(error as Error).message}`, 'error');
        allData = [];
        districtFilter.innerHTML = '';
        districtFilter.disabled = true;
        analyzeBtn.disabled = true;
    }
}


/**
 * Populates the district filter select element with unique districts from the data.
 * @param data The full dataset.
 */
function populateDistrictFilter(data: StationData[]) {
    const districtFilter = document.getElementById('district-filter') as HTMLSelectElement;
    const districts = new Set(data.map(item => item["轄區"]).filter(Boolean));
    districtFilter.innerHTML = '';
    districts.forEach(district => {
        const option = document.createElement('option');
        option.value = district;
        option.textContent = district;
        districtFilter.appendChild(option);
    });
}

/**
 * Main function to trigger data analysis based on user-selected filters.
 */
function handleAnalysis() {
    if (allData.length === 0) {
        showMessage("請先上傳或貼上有效的 JSON 資料。", 'info');
        return;
    }

    try {
        const dateRange = getSelectedDateRange();
        if (!dateRange) return;

        // --- Yearly Stats Calculation ---
        const now = new Date();
        const startOfYear = new Date(now.getFullYear(), 0, 1);
        const endOfRangeA = new Date(dateRange.start);
        endOfRangeA.setDate(endOfRangeA.getDate() - 1);
        endOfRangeA.setHours(23, 59, 59, 999);
        const yearlyStats = calculateYearlyDistrictStats(allData, startOfYear, endOfRangeA, dateRange.start, dateRange.end);
        currentYearlyStats = yearlyStats; // Cache for dynamic completion chart

        // --- Main Data Filtering ---
        const selectedDistricts = getSelectedDistricts();
        const dateFilteredData = allData.filter(item => {
            const itemDate = rocToGregorian(item["本次檢查"]);
            return itemDate && itemDate >= dateRange.start && itemDate <= dateRange.end;
        });

        allFilteredStationsData = selectedDistricts.length === 0
            ? dateFilteredData
            : dateFilteredData.filter(item => selectedDistricts.includes(item["轄區"]));

        if (allFilteredStationsData.length === 0 && Object.keys(yearlyStats).length === 0) {
            showMessage("在選定的篩選條件下找不到資料。", 'info');
            clearResults();
            return;
        }
        
        specialStationsData = allFilteredStationsData.filter(item =>
            Array.from(currentKeywords).some(keyword => item["檢查與改善"]?.includes(keyword))
        );
        
        renderResults(allFilteredStationsData, specialStationsData, yearlyStats);

        showMessage(`分析完成，共找到 ${allFilteredStationsData.length} 筆符合條件的資料。`, 'success');

    } catch (error)
    {
        console.error("Analysis Error:", error);
        showMessage(`分析過程中發生錯誤: ${(error as Error).message}`, 'error');
    }
}

/** Renders all charts and tables with the provided filtered data. */
function renderResults(allData: StationData[], specialData: StationData[], yearlyStats: DistrictStats) {
    // Render charts
    renderBarChart(allData, document.getElementById('chart-container') as HTMLDivElement, "轄區", "轄區統計");
    renderBarChart(allData, document.getElementById('station-type-chart-container') as HTMLDivElement, "型式", "站台型式統計");

    // Render new yearly stats
    renderYearlyStats(yearlyStats, document.getElementById('yearly-stats-container') as HTMLDivElement);

    // Render tables
    renderDataTable(specialData, document.getElementById('table-container') as HTMLDivElement, '未找到符合關鍵字的特殊改善情形。');
    renderDataTable(allData, document.getElementById('all-stations-table-container') as HTMLDivElement, '無符合條件的站台資料。');
    
    // Hide completion chart on new analysis
    const completionContainer = document.getElementById('yearly-completion-container');
    if (completionContainer) completionContainer.style.display = 'none';

    // Show result cards
    (document.getElementById('chart-card') as HTMLDivElement).style.display = 'flex';
    (document.getElementById('station-type-card') as HTMLDivElement).style.display = 'flex';
    (document.getElementById('yearly-stats-card') as HTMLDivElement).style.display = 'flex';
    (document.getElementById('table-card') as HTMLDivElement).style.display = 'flex';
    (document.getElementById('all-stations-card') as HTMLDivElement).style.display = 'flex';
}

/**
 * Determines the selected date range from the UI controls.
 * @returns An object with start and end Date objects, or null if invalid.
 */
function getSelectedDateRange(): { start: Date, end: Date } | null {
    const selectedFilter = (document.querySelector('input[name="date-filter"]:checked') as HTMLInputElement)?.value;
    const now = new Date();

    switch (selectedFilter) {
        case 'week': {
            const startOfWeek = new Date(now);
            startOfWeek.setDate(now.getDate() - now.getDay()); // Sunday
            startOfWeek.setHours(0, 0, 0, 0);
            return { start: startOfWeek, end: new Date() };
        }
        case 'month': {
            const startOfMonth = new Date(now.getFullYear(), now.getMonth(), 1);
            return { start: startOfMonth, end: new Date() };
        }
        case 'month-last': {
            const startOfLastMonth = new Date(now.getFullYear(), now.getMonth() - 1, 1);
            const endOfLastMonth = new Date(now.getFullYear(), now.getMonth(), 0);
            endOfLastMonth.setHours(23, 59, 59, 999);
            return { start: startOfLastMonth, end: endOfLastMonth };
        }
        case 'year': {
            const startOfYear = new Date(now.getFullYear(), 0, 1);
            return { start: startOfYear, end: new Date() };
        }
        case 'custom': {
            const startDateStr = (document.getElementById('start-date') as HTMLInputElement).value;
            const endDateStr = (document.getElementById('end-date') as HTMLInputElement).value;
            if (!startDateStr || !endDateStr) {
                showMessage("請選擇自訂的起訖日期。", 'error');
                return null;
            }
            const start = new Date(startDateStr);
            const end = new Date(endDateStr);
            start.setHours(0, 0, 0, 0);
            end.setHours(23, 59, 59, 999);
            if (start > end) {
                showMessage("起始日期不能晚於結束日期。", 'error');
                return null;
            }
            return { start, end };
        }
        default:
            return null;
    }
}

/** Gets the selected districts from the multi-select dropdown. */
function getSelectedDistricts(): string[] {
    const districtFilter = document.getElementById('district-filter') as HTMLSelectElement;
    return Array.from(districtFilter.selectedOptions).map(option => option.value);
}

/** Clears the selection in the district multi-select dropdown. */
function handleClearDistrictSelection() {
    const districtFilter = document.getElementById('district-filter') as HTMLSelectElement;
    Array.from(districtFilter.options).forEach(option => {
        option.selected = false;
    });
}

/**
 * Converts a Republic of China (ROC) date string to a Gregorian Date object.
 * @param rocDateStr The date string in "YYY/MM/DD" format.
 * @returns A Date object or null if the format is invalid.
 */
function rocToGregorian(rocDateStr: string): Date | null {
    if (!rocDateStr || typeof rocDateStr !== 'string') return null;
    const parts = rocDateStr.split(/[\/\-]/);
    if (parts.length !== 3) return null;

    const rocYear = parseInt(parts[0], 10);
    const month = parseInt(parts[1], 10);
    const day = parseInt(parts[2], 10);

    if (isNaN(rocYear) || isNaN(month) || isNaN(day)) return null;

    const gregorianYear = rocYear + 1911;
    return new Date(gregorianYear, month - 1, day);
}

// --- Keyword Management ---

function handleAddKeyword() {
    const input = document.getElementById('custom-keyword-input') as HTMLInputElement;
    const newKeywords = input.value.split(/[,，\s]+/).map(k => k.trim()).filter(Boolean);
    if (newKeywords.length > 0) {
        newKeywords.forEach(k => currentKeywords.add(k));
        renderKeywords();
        input.value = '';
    }
}

function handleRemoveKeyword(keywordToRemove: string) {
    currentKeywords.delete(keywordToRemove);
    renderKeywords();
}

function renderKeywords() {
    const container = document.getElementById('keyword-list-container') as HTMLDivElement;
    container.innerHTML = '';
    currentKeywords.forEach(keyword => {
        const tag = document.createElement('span');
        tag.className = 'keyword-tag';
        tag.textContent = keyword;
        
        const removeBtn = document.createElement('span');
        removeBtn.className = 'remove-keyword';
        removeBtn.textContent = '×';
        removeBtn.dataset.keyword = keyword;
        removeBtn.setAttribute('aria-label', `移除關鍵字 ${keyword}`);

        tag.appendChild(removeBtn);
        container.appendChild(tag);
    });
}

// --- Calculation Functions ---

function calculateYearlyDistrictStats(
    data: StationData[],
    startA: Date,
    endA: Date,
    startB: Date,
    endB: Date
): DistrictStats {
    const stats: DistrictStats = {};

    data.forEach(item => {
        const district = item["轄區"];
        if (!district) return; // Skip items without a district

        if (!stats[district]) {
            stats[district] = { countA: 0, countB: 0, countC: 0 };
        }

        const itemDate = rocToGregorian(item["本次檢查"]);
        if (!itemDate) return; // Skip items without a valid date

        // Check for Range A (Year-to-date before selected range)
        if (itemDate >= startA && itemDate <= endA) {
            stats[district].countA++;
        }
        // Check for Range B (The selected range)
        else if (itemDate >= startB && itemDate <= endB) {
            stats[district].countB++;
        }
    });

    // Calculate C = A + B for all districts
    for (const district in stats) {
        stats[district].countC = stats[district].countA + stats[district].countB;
    }

    return stats;
}


// --- Rendering Functions ---

/**
 * Aggregates data by a given property and renders a bar chart.
 * @param data The filtered data to be visualized.
 * @param container The container element for the chart.
 * @param property The property of StationData to group by.
 * @param title The title for the chart.
 */
function renderBarChart(data: StationData[], container: HTMLElement, property: keyof StationData, title: string) {
    const counts = data.reduce((acc, item) => {
        const key = String(item[property] || "未知");
        acc[key] = (acc[key] || 0) + 1;
        return acc;
    }, {} as Record<string, number>);

    container.innerHTML = '';
    const maxValue = Math.max(...Object.values(counts), 0);
    if (maxValue === 0) {
        container.innerHTML = '<p>無資料可顯示圖表。</p>';
        return;
    }

    Object.entries(counts)
        .sort(([, a], [, b]) => b - a)
        .forEach(([key, count]) => {
            const barWrapper = document.createElement('div');
            barWrapper.className = 'chart-bar-wrapper';
            barWrapper.setAttribute('aria-label', `${key}: ${count} 次`);

            const label = document.createElement('div');
            label.className = 'chart-label';
            label.textContent = key;
            label.title = key;

            const bar = document.createElement('div');
            bar.className = 'chart-bar';
            bar.style.width = `${(count / maxValue) * 100}%`;
            bar.textContent = String(count);

            barWrapper.appendChild(label);
            barWrapper.appendChild(bar);
            container.appendChild(barWrapper);
        });
}

/** Renders the yearly district statistics into a table. */
function renderYearlyStats(stats: DistrictStats, container: HTMLElement) {
    const sortedDistricts = Object.keys(stats).sort();

    if (sortedDistricts.length === 0) {
        container.innerHTML = '<p>無資料可顯示統計。</p>';
        return;
    }

    const table = document.createElement('table');
    const thead = document.createElement('thead');
    const tbody = document.createElement('tbody');
    const tfoot = document.createElement('tfoot');

    thead.innerHTML = `<tr>
        <th>轄區</th>
        <th>年度累計 (A)</th>
        <th>本期新增 (B)</th>
        <th>總計 (C=A+B)</th>
        <th>年度目標</th>
    </tr>`;

    let totalA = 0, totalB = 0, totalC = 0;

    sortedDistricts.forEach(district => {
        const districtStats = stats[district];
        totalA += districtStats.countA;
        totalB += districtStats.countB;
        totalC += districtStats.countC;
        const row = document.createElement('tr');
        row.innerHTML = `
            <td>${district}</td>
            <td>${districtStats.countA}</td>
            <td>${districtStats.countB}</td>
            <td>${districtStats.countC}</td>
            <td><input type="number" class="target-input" min="0" data-district="${district}" placeholder="輸入目標"></td>
        `;
        tbody.appendChild(row);
    });

    const footerRow = document.createElement('tr');
    footerRow.innerHTML = `
        <td><strong>全轄區總計</strong></td>
        <td>${totalA}</td>
        <td>${totalB}</td>
        <td>${totalC}</td>
        <td></td>
    `;
    tfoot.appendChild(footerRow);

    table.appendChild(thead);
    table.appendChild(tbody);
    table.appendChild(tfoot);
    container.innerHTML = '';
    container.appendChild(table);
}

/**
 * Handles input changes on target fields and triggers completion chart rendering.
 */
function handleTargetInputChange() {
    const statsContainer = document.getElementById('yearly-stats-container') as HTMLElement;
    const inputs = Array.from(statsContainer.querySelectorAll<HTMLInputElement>('.target-input'));
    const completionContainer = document.getElementById('yearly-completion-container') as HTMLElement;

    const allFilled = inputs.every(input => input.value && Number(input.value) > 0);

    if (allFilled) {
        const targets = new Map<string, number>();
        inputs.forEach(input => {
            if (input.dataset.district) {
                targets.set(input.dataset.district, Number(input.value));
            }
        });
        renderCompletionVisuals(currentYearlyStats, targets);
        completionContainer.style.display = 'block';
    } else {
        completionContainer.style.display = 'none';
    }
}

/**
 * Renders the completion rate chart based on stats and user-provided targets.
 * Ensures the district sort order matches the yearly stats table.
 */
function renderCompletionVisuals(stats: DistrictStats, targets: Map<string, number>) {
    const container = document.getElementById('yearly-completion-chart-container') as HTMLElement;
    const completionData: { key: string, percentage: number }[] = [];
    
    let totalC = 0;
    let totalTarget = 0;

    // Use the same sorting as the table to ensure consistent order
    const sortedDistricts = Object.keys(stats).sort();

    sortedDistricts.forEach(district => {
        const target = targets.get(district);
        if (target && target > 0) {
            completionData.push({
                key: district,
                percentage: (stats[district].countC / target) * 100
            });
            totalC += stats[district].countC;
            totalTarget += target;
        }
    });

    if (totalTarget > 0) {
        completionData.push({
            key: '總轄區',
            percentage: (totalC / totalTarget) * 100
        });
    }
    
    container.innerHTML = '';
    if (completionData.length === 0) {
        container.innerHTML = '<p>無法計算完成率。</p>';
        return;
    }
    
    const maxValue = Math.max(100, ...completionData.map(d => d.percentage), 0);

    completionData.forEach(({ key, percentage }) => {
        const barWrapper = document.createElement('div');
        barWrapper.className = 'chart-bar-wrapper';
        barWrapper.setAttribute('aria-label', `${key}: ${percentage.toFixed(1)}%`);

        const label = document.createElement('div');
        label.className = 'chart-label';
        label.textContent = key;
        label.title = key;

        const bar = document.createElement('div');
        bar.className = 'chart-bar';
        const barWidthPercentage = (percentage / maxValue) * 100;
        bar.style.width = `${Math.min(barWidthPercentage, 100)}%`;
        bar.textContent = `${percentage.toFixed(1)}%`;
        if (percentage > 100) {
            bar.style.backgroundColor = 'var(--success-color)';
        }

        barWrapper.appendChild(label);
        barWrapper.appendChild(bar);
        container.appendChild(barWrapper);
    });
}


/**
 * Renders data into a table.
 * @param data The data to render.
 * @param container The container element for the table.
 * @param emptyMessage Message to display if data is empty.
 */
function renderDataTable(data: StationData[], container: HTMLElement, emptyMessage: string) {
    if (data.length === 0) {
        container.innerHTML = `<p>${emptyMessage}</p>`;
        return;
    }

    const table = document.createElement('table');
    const thead = document.createElement('thead');
    const tbody = document.createElement('tbody');

    const headers = ["基地台名稱", "轄區", "型式", "本次檢查", "檢查與改善"];
    thead.innerHTML = `<tr>${headers.map(h => `<th>${h}</th>`).join('')}</tr>`;

    data.forEach(item => {
        const row = document.createElement('tr');
        row.innerHTML = `
            <td>${item["基地台名稱"] || ''}</td>
            <td>${item["轄區"] || ''}</td>
            <td>${item["型式"] || ''}</td>
            <td>${item["本次檢查"] || ''}</td>
            <td>${item["檢查與改善"] || ''}</td>
        `;
        tbody.appendChild(row);
    });

    table.appendChild(thead);
    table.appendChild(tbody);
    container.innerHTML = '';
    container.appendChild(table);
}

// --- Utility Functions ---

/** Displays a message to the user. */
function showMessage(message: string, type: 'error' | 'success' | 'info') {
    const container = document.getElementById('message-container') as HTMLDivElement;
    container.textContent = message;
    container.className = `message-${type}`;
    container.style.display = 'block';
}

/** Clears all rendered results from the DOM. */
function clearResults() {
    ['chart-container', 'station-type-chart-container', 'yearly-stats-container', 'table-container', 'all-stations-table-container'].forEach(id => {
        const el = document.getElementById(id);
        if (el) el.innerHTML = '';
    });
    ['chart-card', 'station-type-card', 'yearly-stats-card', 'table-card', 'all-stations-card'].forEach(id => {
        const el = document.getElementById(id);
        if (el) el.style.display = 'none';
    });
}

// --- Data Export ---

function handleExport(dataType: 'special' | 'all', format: 'csv' | 'json') {
    const data = dataType === 'special' ? specialStationsData : allFilteredStationsData;
    if (data.length === 0) {
        showMessage('沒有可匯出的資料。', 'info');
        return;
    }

    const filename = `${dataType}_export_${new Date().toISOString().slice(0, 10)}`;
    if (format === 'json') {
        const jsonString = JSON.stringify(data, null, 2);
        downloadFile(jsonString, 'application/json', `${filename}.json`);
    } else if (format === 'csv') {
        const csvString = convertToCSV(data);
        downloadFile(csvString, 'text/csv;charset=utf-8;', `${filename}.csv`);
    }
}

function convertToCSV(data: StationData[]): string {
    if (data.length === 0) return '';
    const headers = Object.keys(data[0]) as (keyof StationData)[];
    const headerRow = headers.join(',');
    const rows = data.map(row => {
        return headers.map(header => {
            let value = String(row[header] ?? '');
            if (value.includes(',')) {
                value = `"${value.replace(/"/g, '""')}"`;
            }
            return value;
        }).join(',');
    });
    return [headerRow, ...rows].join('\r\n');
}

function downloadFile(content: string, mimeType: string, filename: string) {
    const blob = new Blob([content], { type: mimeType });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
}