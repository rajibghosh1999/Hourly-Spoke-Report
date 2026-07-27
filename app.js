/*
===========================================
 Hourly Spoke Performance Logic & ZIP Export
 North 24 Parganas
===========================================
*/

// Direct Google Apps Script Master URL
const MASTER_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbyYJ88WTK1IoS3w2MU1S96AnipUPmlem926W07VNqYUdWRoTG64KM_zoo3Ark2PgFkk/exec";

const hourlyState = {
    masterData: [],
    uploadedSummary: [],
    selectedGroup: "ALL",
    currentReportTime: "10 AM",
    
    // Block Groups
    cho1Blocks: ["BAGDAH", "DEGANGA", "GAIGHATA", "HABRA-I", "HABRA-II", "HABRA - I", "HABRA - II", "RAJARHAT"],
    cho2Blocks: ["AMDANGA", "BARASAT-I", "BARASAT-II", "BARASAT - I", "BARASAT - II", "BARRACKPORE-I", "BARRACKPORE-II", "BARRACKPUR - 1", "BARRACKPUR - 2", "BONGAON"]
};

document.addEventListener('DOMContentLoaded', () => {
    initHourlyDashboard();
});

function initHourlyDashboard() {
    initDateDisplay();
    setupEventListeners();
    fetchHourlyMasterData();
}

// ১. তারিখ সেট করার ফাংশন
function initDateDisplay() {
    const today = new Date();
    const formattedDate = today.toLocaleDateString('en-IN', { 
        day: '2-digit', month: 'short', year: 'numeric' 
    });
    
    if (document.getElementById('reportDate')) {
        document.getElementById('reportDate').textContent = formattedDate;
    }
    
    updateReportTime();
}

// ফাইল নেমিং এর জন্য ফর্মেটেড ডেট
function getFilenameDate() {
    return new Date().toLocaleDateString('en-IN', { 
        day: '2-digit', month: 'short', year: 'numeric' 
    }).replace(/ /g, '_');
}

// ফাইল নেমিং ফরম্যাট জেনারেটর
function getStandardFilename(ext) {
    const dateStr = getFilenameDate(); 
    const timeStr = hourlyState.currentReportTime.replace(/\s/g, ''); 
    const group = hourlyState.selectedGroup;
    return `${dateStr}_N24PGS_Spoke_Report_${timeStr}_${group}.${ext}`;
}

// ২. টাইম সিলেক্ট করলে আপডেট ফাংশন
function updateReportTime() {
    const timeSelect = document.getElementById('timeSlotSelect');
    const selectedTime = timeSelect ? timeSelect.value : "10 AM";
    
    hourlyState.currentReportTime = selectedTime;
    
    if (document.getElementById('reportTime')) {
        document.getElementById('reportTime').textContent = selectedTime;
    }
}

/**
 * Fetch Spoke Master Data
 */
async function fetchHourlyMasterData() {
    showLoading(true, "Loading...");

    try {
        const res = await fetch(MASTER_SCRIPT_URL, { method: 'GET', redirect: 'follow' });
        if (res.ok) {
            const data = await res.json();
            hourlyState.masterData = parseSSKMasterArray(data);
            if (hourlyState.uploadedSummary.length > 0) {
                processAndRenderReport();
            }
        }
    } catch (err) {
        console.error('Master Data Fetch Error:', err);
    } finally {
        showLoading(false);
    }
}

function parseSSKMasterArray(rawList) {
    if (!Array.isArray(rawList)) return [];
    return rawList.map(item => {
        const keys = Object.keys(item);
        const spokeKey = keys.find(k => /aam|spoke|ssk|facility/i.test(k)) || keys[0];
        const blockKey = keys.find(k => /block/i.test(k)) || keys[1];
        const choKey   = keys.find(k => /cho|officer|provider/i.test(k)) || keys[2];

        const rawSpoke = String(item[spokeKey] || '').trim();
        const rawBlock = String(item[blockKey] || '').trim();
        const rawCho   = String(item[choKey] || '').trim();

        return {
            spokeName: normalizeName(rawSpoke),
            originalSpokeName: rawSpoke,
            blockName: rawBlock || 'Unassigned',
            choName: rawCho || 'N/A'
        };
    }).filter(item => item.originalSpokeName !== '');
}

function normalizeName(str) {
    if (!str) return '';
    return str.toString()
        .toUpperCase()
        .replace(/\bAAM\b/g, '')
        .replace(/\bSC\b/g, '')
        .replace(/SUBCENTER|SUB-CENTER/g, '')
        .replace(/[^A-Z0-9]/g, '')
        .trim();
}

function setupEventListeners() {
    document.getElementById('generateReport')?.addEventListener('click', handleFileUpload);
    document.getElementById('timeSlotSelect')?.addEventListener('change', updateReportTime);
    document.getElementById('btnDownloadJpg')?.addEventListener('click', downloadReportAsJpg);
    
    document.getElementById('groupSelect')?.addEventListener('change', (e) => {
        hourlyState.selectedGroup = e.target.value;
        if (hourlyState.uploadedSummary.length > 0) {
            processAndRenderReport();
        }
    });

    document.getElementById('btnExportExcel')?.addEventListener('click', exportExcel);
    document.getElementById('btnDownloadZip')?.addEventListener('click', generateBlockWiseZipJPGs);
    document.getElementById('btnPrint')?.addEventListener('click', () => window.print());
}

function handleFileUpload() {
    const fileInput = document.getElementById('excelFile');
    const file = fileInput?.files[0];

    if (!file) {
        alert('অনুগ্রহ করে Excel ফাইলটি নির্বাচন করুন।');
        return;
    }

    showLoading(true, "Processing Uploaded File...");

    const reader = new FileReader();
    reader.onload = function (e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const sheet = workbook.Sheets[workbook.SheetNames[0]];
            hourlyState.uploadedSummary = XLSX.utils.sheet_to_json(sheet, { defval: "" });

            processAndRenderReport();
        } catch (err) {
            console.error(err);
            alert('Excel ফাইল প্রসেস করতে সমস্যা হয়েছে।');
        } finally {
            showLoading(false);
        }
    };

    reader.readAsArrayBuffer(file);
}

function processAndRenderReport() {
    const consultationMap = new Map();

    hourlyState.uploadedSummary.forEach(row => {
        const keys = Object.keys(row);
        const spokeCol = keys.find(k => /aam|spoke|ssk|phc|facility|health/i.test(k));
        const countCol = keys.find(k => /completed|consultation|count|total/i.test(k));

        if (spokeCol) {
            const cleanKey = normalizeName(row[spokeCol]);
            let count = countCol ? (parseInt(row[countCol], 10) || 0) : 1;
            if (cleanKey) {
                consultationMap.set(cleanKey, (consultationMap.get(cleanKey) || 0) + count);
            }
        }
    });

    updateTableHeaders(hourlyState.selectedGroup);

    if (hourlyState.selectedGroup === "BLOCK_WISE") {
        const blockSummary = {};

        hourlyState.masterData.forEach(m => {
            if (!blockSummary[m.blockName]) {
                blockSummary[m.blockName] = { performed: 0, totalConsultation: 0, totalSKs: 0 };
            }
            blockSummary[m.blockName].totalSKs++;
            
            const count = consultationMap.get(m.spokeName) || 0;
            if (count > 0) {
                blockSummary[m.blockName].performed++;
                blockSummary[m.blockName].totalConsultation += count;
            }
        });

        renderBlockSummaryTable(blockSummary);
        return; 
    }

    const isBlockMatchingGroup = (blockName, group) => {
        if (!blockName || group === "ALL") return true;
        
        const cleanBlk = blockName.toUpperCase().replace(/[^A-Z]/g, '');
        
        if (group === "CHO_1") {
            return hourlyState.cho1Blocks.some(b => {
                const target = b.toUpperCase().replace(/[^A-Z]/g, '');
                return cleanBlk.includes(target) || target.includes(cleanBlk);
            });
        } else if (group === "CHO_2") {
            return hourlyState.cho2Blocks.some(b => {
                const target = b.toUpperCase().replace(/[^A-Z]/g, '');
                return cleanBlk.includes(target) || target.includes(cleanBlk);
            });
        }
        return true;
    };

    let mergedList = [];

    if (hourlyState.masterData.length > 0) {
        let filteredMaster = hourlyState.masterData.filter(m => 
            isBlockMatchingGroup(m.blockName, hourlyState.selectedGroup)
        );

        mergedList = filteredMaster.map(master => {
            const completedCount = consultationMap.get(master.spokeName) || 0;
            return {
                spokeName: master.originalSpokeName,
                blockName: master.blockName,
                choName: master.choName,
                completedConsultation: completedCount
            };
        });
    } else {
        hourlyState.uploadedSummary.forEach(row => {
            const keys = Object.keys(row);
            const spokeCol = keys.find(k => /aam|spoke|ssk|phc|facility|health/i.test(k));
            const blockCol = keys.find(k => /block/i.test(k));
            const choCol   = keys.find(k => /cho|officer|provider/i.test(k));
            const countCol = keys.find(k => /completed|consultation|count|total/i.test(k));

            if (spokeCol && row[spokeCol]) {
                const blkName = blockCol ? String(row[blockCol]).trim() : 'Unassigned';
                if (isBlockMatchingGroup(blkName, hourlyState.selectedGroup)) {
                    mergedList.push({
                        spokeName: row[spokeCol],
                        blockName: blkName,
                        choName: choCol ? row[choCol] : 'N/A',
                        completedConsultation: countCol ? (parseInt(row[countCol], 10) || 0) : 0
                    });
                }
            }
        });
    }

    mergedList.sort((a, b) => {
        if (a.blockName.localeCompare(b.blockName) !== 0) {
            return a.blockName.localeCompare(b.blockName);
        }
        return a.completedConsultation - b.completedConsultation;
    });

    renderHourlyTable(mergedList);
}

function renderHourlyTable(reportList) {
    const tbody = document.getElementById('hourlyReportBody');
    if (!tbody) return;
    tbody.innerHTML = '';

    if (reportList.length === 0) {
        tbody.innerHTML = `<tr><td colspan="5" class="text-center py-4 text-danger fw-bold">কোনো ডাটা পাওয়া যায়নি। ফাইল বা গ্রুপ ঠিক আছে কিনা চেক করুন।</td></tr>`;
        return;
    }

    reportList.forEach((row, index) => {
        const tr = document.createElement('tr');
        const rowColor = getRowColorByPerformance(row.completedConsultation);
        const style = `background-color: ${rowColor} !important; -webkit-print-color-adjust: exact; print-color-adjust: exact;`;

        tr.innerHTML = `
            <td class="text-center border-dark fw-bold" style="${style}">${index + 1}</td>
            <td class="border-dark" style="${style}">${row.spokeName}</td>
            <td class="border-dark fw-bold" style="${style}">${row.blockName}</td>
            <td class="border-dark" style="${style}">${row.choName}</td>
            <td class="text-center fw-bold border-dark" style="${style}">${row.completedConsultation}</td>
        `;
        tbody.appendChild(tr);
    });
}

function getRowColorByPerformance(count) {
    if (count <= 5) return "#f8d7da";
    else if (count <= 9) return "#ffe8cc";
    else return "#d4edda";
}

/**
 * Generate Block Canvas & Image Array
 */
async function generateBlockCanvasList() {
    try {
        const tbody = document.getElementById('hourlyReportBody');
        const rows = Array.from(tbody.querySelectorAll('tr'));

        if (rows.length === 0 || rows[0].cells.length < 5) return null;

        const blockGroups = {};
        rows.forEach(tr => {
            const blockName = (hourlyState.selectedGroup === "BLOCK_WISE") ? tr.cells[1]?.textContent.trim() : tr.cells[2]?.textContent.trim();
            if (blockName) {
                if (!blockGroups[blockName]) blockGroups[blockName] = [];
                blockGroups[blockName].push({
                    spoke: tr.cells[1]?.textContent || '',
                    block: blockName,
                    cho: tr.cells[3]?.textContent || 'N/A',
                    count: tr.cells[4]?.textContent || '0',
                    bgColor: tr.cells[0]?.style.backgroundColor || '#ffffff'
                });
            }
        });

        const renderContainer = document.getElementById('zipRenderContainer');
        const dateStr = getFilenameDate();
        const timeStr = hourlyState.currentReportTime;
        const blockNames = Object.keys(blockGroups);
        const resultFiles = [];

        for (let i = 0; i < blockNames.length; i++) {
            const blk = blockNames[i];
            const list = blockGroups[blk];

            const cardDiv = document.createElement('div');
            cardDiv.className = "p-3 bg-white border border-dark";
            cardDiv.style.width = "800px";

            let tableRowsHtml = list.map((item, idx) => `
                <tr style="background-color: ${item.bgColor} !important;">
                    <td style="text-align:center; border: 1px solid #000; padding: 5px; font-weight: bold; background-color: ${item.bgColor} !important;">${idx + 1}</td>
                    <td style="border: 1px solid #000; padding: 5px; background-color: ${item.bgColor} !important;">${item.spoke}</td>
                    <td style="border: 1px solid #000; padding: 5px; font-weight: bold; background-color: ${item.bgColor} !important;">${item.block}</td>
                    <td style="border: 1px solid #000; padding: 5px; background-color: ${item.bgColor} !important;">${item.cho}</td>
                    <td style="text-align:center; border: 1px solid #000; padding: 5px; font-weight: bold; background-color: ${item.bgColor} !important;">${item.count}</td>
                </tr>
            `).join('');

            cardDiv.innerHTML = `
                <div style="text-align: center; margin-bottom: 15px;">
                    <h3 style="color: #0f766e; margin: 0; font-weight: bold;">Spoke Performance Report</h3>
                    <h4 style="margin: 5px 0; color: #1e3a8a;">Block: ${blk} (North 24 Parganas)</h4>
                    <p style="margin: 0; color: #555;">Report Date: <b>${dateStr.replace(/_/g, ' ')}</b> | Time: <b>${timeStr}</b></p>
                </div>
                <table style="width: 100%; border-collapse: collapse; font-family: sans-serif; font-size: 13px;">
                    <thead>
                        <tr style="background-color: #14532d; color: #ffffff; text-align: center; vertical-align: middle;">
                            <th style="border: 1px solid #000; padding: 6px; width: 8%;">Sl No.</th>
                            <th style="border: 1px solid #000; padding: 6px; width: 32%;">AAM SSK Name</th>
                            <th style="border: 1px solid #000; padding: 6px; width: 18%;">Block Name</th>
                            <th style="border: 1px solid #000; padding: 6px; width: 27%;">CHO Name</th>
                            <th style="border: 1px solid #000; padding: 6px; width: 15%;">Total Completed Consultation</th>
                        </tr>
                    </thead>
                    <tbody>${tableRowsHtml}</tbody>
                </table>
            `;

            renderContainer.appendChild(cardDiv);
            const canvas = await html2canvas(cardDiv, { scale: 2, useCORS: true });
            const fileName = `${blk}_Spoke_Performance_${timeStr}_${dateStr}.jpg`;

            resultFiles.push({
                blockName: blk,
                fileName: fileName,
                base64: canvas.toDataURL('image/jpeg', 0.9).split(',')[1]
            });

            renderContainer.removeChild(cardDiv);
        }
        return resultFiles;
    } catch (err) {
        console.error("Canvas Generation Error:", err);
        alert("ইমেজ জেনারেট করতে সমস্যা হচ্ছে: " + err.message);
        return null;
    }
}

/**
 * ZIP Download
 */
async function generateBlockWiseZipJPGs() {
    showLoading(true, "Creating Block Wise JPGs & Packing ZIP...");

    const fileList = await generateBlockCanvasList();

    if (!fileList || fileList.length === 0) {
        alert("ডাউনলোড করার মতো কোনো রিপোর্ট ডাটা নেই!");
        showLoading(false);
        return;
    }

    try {
        const zip = new JSZip();
        fileList.forEach(item => {
            zip.file(item.fileName, item.base64, { base64: true });
        });

        const zipFilename = getStandardFilename('zip');
        zip.generateAsync({ type: "blob" }).then(content => {
            saveAs(content, zipFilename);
            showLoading(false);
        });
    } catch (err) {
        console.error("ZIP Creation Error:", err);
        alert("ZIP তৈরি করতে সমস্যা হয়েছে।");
        showLoading(false);
    }
}

function exportExcel() {
    const table = document.querySelector("#mainReportCard table");
    if (!table) return;

    const fileName = getStandardFilename('xlsx');
    const wb = XLSX.utils.table_to_book(table, { sheet: "Spoke Performance" });
    XLSX.writeFile(wb, fileName);
}

function showLoading(isLoading, text = "Processing Data...") {
    const loadingElem = document.getElementById('loading');
    const loadingText = document.getElementById('loadingText');
    if (loadingElem) {
        if (loadingText) loadingText.textContent = text;
        loadingElem.classList.toggle('d-none', !isLoading);
    }
}

// হেডারের ডিজাইন বদলানোর জন্য
function updateTableHeaders(mode) {
    const thead = document.querySelector('.subhead-dark-green');
    if (mode === 'BLOCK_WISE') {
        thead.innerHTML = `<tr><th>Sl No.</th><th>Block Name</th><th>Total AAM SKs</th><th>Performed</th><th>Consultation</th><th>% Active</th></tr>`;
    } else {
        thead.innerHTML = `<tr><th>Sl No.</th><th>AAM SSK Name</th><th>Block Name</th><th>CHO Name</th><th>Total Completed Consultation</th></tr>`;
    }
}

// % Active কালার কোডিং
function getPercentColor(pct) {
    if (pct < 50) return "#d9534f"; // DEEP লাল
    if (pct < 70) return "#f8d7da"; // হালকা লাল
    if (pct < 80) return "#ffe8cc"; // হালকা কমলা
    return "#d4edda";               // সবুজ
}

function renderBlockSummaryTable(data) {
    const tbody = document.getElementById('hourlyReportBody');
    tbody.innerHTML = ''; 

    Object.keys(data).sort().forEach((block, i) => {
        const item = data[block];
        const pct = item.totalSKs > 0 ? ((item.performed / item.totalSKs) * 100).toFixed(1) : 0;
        const rowColor = getPercentColor(parseFloat(pct));
        
        const style = `background-color: ${rowColor} !important; -webkit-print-color-adjust: exact; print-color-adjust: exact;`;
        
        tbody.innerHTML += `
            <tr>
                <td class="text-center border-dark fw-bold" style="${style}">${i+1}</td>
                <td class="border-dark fw-bold" style="${style}">${block}</td>
                <td class="text-center border-dark fw-bold" style="${style}">${item.totalSKs}</td>
                <td class="text-center border-dark fw-bold text-success" style="${style}">${item.performed}</td>
                <td class="text-center border-dark fw-bold text-primary" style="${style}">${item.totalConsultation}</td>
                <td class="text-center border-dark fw-bold" style="${style}">${pct}%</td>
            </tr>`;
    });
}

// রিপোর্ট JPG ডাউনলোড
async function downloadReportAsJpg() {
    showLoading(true, "Generating Report Image...");

    // ১. যদি BLOCK_WISE সিলেক্ট করা থাকে, তবে আমরা পুরো টেবিলটিই নেব
    if (hourlyState.selectedGroup === "BLOCK_WISE") {
        const reportElement = document.getElementById('mainReportCard');
        
        try {
            // পুরো কার্ডটি (টেবিলসহ) ক্যাপচার করবে
            const canvas = await html2canvas(reportElement, { 
                scale: 2, 
                useCORS: true,
                backgroundColor: "#ffffff" 
            });
            
            const link = document.createElement('a');
            link.download = getStandardFilename('jpg'); 
            link.href = canvas.toDataURL('image/jpeg', 0.9);
            link.click();
        } catch (err) {
            console.error("JPG Export Error:", err);
            alert("ইমেজ তৈরি করতে সমস্যা হয়েছে।");
        } finally {
            showLoading(false);
        }
    } else {
        // যদি অন্য গ্রুপ হয়, তবে আগের মতো ব্লক-ওয়াইজ আলাদা ইমেজ ডাউনলোড হবে
        const fileList = await generateBlockCanvasList();
        if (fileList) {
            fileList.forEach(item => {
                const link = document.createElement('a');
                link.href = 'data:image/jpeg;base64,' + item.base64;
                link.download = item.fileName;
                link.click();
            });
        }
        showLoading(false);
    }
}