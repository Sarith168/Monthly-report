import { initializeApp } from "https://www.gstatic.com/firebasejs/11.0.2/firebase-app.js";
import { getAnalytics } from "https://www.gstatic.com/firebasejs/11.0.2/firebase-analytics.js";
import { getFirestore, doc, setDoc, onSnapshot, serverTimestamp } from "https://www.gstatic.com/firebasejs/11.0.2/firebase-firestore.js";

// Firebase Configuration
const firebaseConfig = {
    apiKey: "AIzaSyB5N68qMLC-KolKBszkzUjV7KPZHor-WcQ",
    authDomain: "monthly-report-a1509.firebaseapp.com",
    projectId: "monthly-report-a1509",
    storageBucket: "monthly-report-a1509.firebasestorage.app",
    messagingSenderId: "614298328019",
    appId: "1:614298328019:web:ea8c365306a66b95066d26",
    measurementId: "G-FNYH77L51J"
};

// Initialize Firebase, Analytics & Firestore
const app = initializeApp(firebaseConfig);
const analytics = getAnalytics(app);
const db = getFirestore(app);
const REPORT_DOC_REF = doc(db, "kads_reports", "latest_monthly_report");

Chart.register(ChartDataLabels);

let allTasks = [];
let filteredTasks = [];
let tablePage = 1;
const pageSize = 10;

let moduleChart = null;
let priorityDoughnut = null;
let assigneeDonutChart = null;

function setSyncBadge(state, text) {
    const badge = document.getElementById('firebaseSyncStatus');
    if (!badge) return;
    if (state === 'connected') {
        badge.innerHTML = `<span class="w-2 h-2 rounded-full bg-emerald-400"></span> ${text}`;
    } else if (state === 'syncing') {
        badge.innerHTML = `<span class="w-2 h-2 rounded-full bg-amber-300 animate-ping"></span> ${text}`;
    } else {
        badge.innerHTML = `<span class="w-2 h-2 rounded-full bg-rose-400"></span> ${text}`;
    }
}

// Initialize UI & Subscribe to Firestore Real-Time Updates
document.addEventListener('DOMContentLoaded', function () {
    try { if (window.lucide) lucide.createIcons(); } catch (e) {}

    initCharts();

    // Event Listeners
    document.getElementById('excelFileInput').addEventListener('change', handleExcelUpload);
    document.getElementById('monthFilterSelect').addEventListener('change', applyMonthFilter);
    document.getElementById('btnResetMonth').addEventListener('click', resetMonthFilter);

    document.getElementById('tableSearchInput').addEventListener('input', () => { tablePage = 1; renderTable(); });
    document.getElementById('tableViewMode').addEventListener('change', () => { tablePage = 1; renderTable(); });
    document.getElementById('tableAssigneeFilter').addEventListener('change', () => { tablePage = 1; renderTable(); });

    // Listen to Firestore document in real time
    onSnapshot(REPORT_DOC_REF, (docSnap) => {
        if (docSnap.exists()) {
            const data = docSnap.data();
            allTasks = Array.isArray(data.tasks) ? data.tasks : [];

            let updatedInfo = "Synced";
            if (data.updatedAt && data.updatedAt.toDate) {
                const d = data.updatedAt.toDate();
                updatedInfo = `Synced: ${d.toLocaleDateString()} ${d.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })}`;
            }
            setSyncBadge('connected', updatedInfo);

            populateMonthDropdown();
            populateAssigneeDropdown();
            applyMonthFilter();
        } else {
            setSyncBadge('connected', 'Cloud Ready (Empty)');
            document.getElementById('kpiTotalCompareDesc').textContent = 'Please upload an Excel file';
            document.getElementById('filterActiveLabel').textContent = 'Awaiting first upload...';
            document.getElementById('pendingTableBody').innerHTML = `
                <tr>
                    <td colspan="9" class="text-center py-12 text-slate-400 italic">
                        <div class="flex flex-col items-center justify-center gap-2">
                            <i data-lucide="file-spreadsheet" class="w-8 h-8 text-slate-300"></i>
                            <span>No saved cloud data yet. Click <strong>"Upload & Sync Excel"</strong> above to publish your report.</span>
                        </div>
                    </td>
                </tr>`;
            try { if (window.lucide) lucide.createIcons(); } catch (e) {}
        }
    }, (error) => {
        console.error("Firestore error:", error);
        setSyncBadge('error', 'Cloud Permission / Offline');
        document.getElementById('kpiTotalCompareDesc').textContent = 'Check Firestore rules or connection';
    });
});

async function handleExcelUpload(e) {
    const file = e.target.files[0];
    if (!file) return;

    const uploadBtnText = document.getElementById('uploadBtnText');
    uploadBtnText.textContent = 'Uploading...';
    setSyncBadge('syncing', 'Saving to Firebase...');

    const reader = new FileReader();
    reader.onload = async function (evt) {
        try {
            const data = new Uint8Array(evt.target.result);
            const workbook = XLSX.read(data, { type: 'array' });

            const sheetName = workbook.SheetNames.includes('Consolidated Data') ? 'Consolidated Data' : workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: '' });

            if (jsonData && jsonData.length > 0) {
                const parsedTasks = jsonData.map(row => {
                    return {
                        'Task ID': String(row['Task ID'] || ''),
                        'Task Name': String(row['Task Name'] || ''),
                        'Bucket': String(row['Bucket'] || ''),
                        'Goal': String(row['Goal'] || ''),
                        'Status': String(row['Status'] || ''),
                        'Priority': String(row['Priority'] || 'Medium'),
                        'Assigned To': String(row['Assigned To'] || ''),
                        'Created By': String(row['Created By'] || ''),
                        'Created Date': formatDateStr(row['Created Date']),
                        'Due date': formatDateStr(row['Due date']),
                        'Start date': formatDateStr(row['Start date']),
                        'Late': row['Late'] === true || row['Late'] === 'True' || row['Late'] === 'TRUE',
                        'Completed Date': formatDateStr(row['Completed Date']),
                        'Completed By': String(row['Completed By'] || ''),
                        'Labels': String(row['Labels'] || ''),
                        'Notes': String(row['Notes'] || '')
                    };
                });

                // Save to Firebase Firestore
                await setDoc(REPORT_DOC_REF, {
                    fileName: file.name,
                    totalRecords: parsedTasks.length,
                    updatedAt: serverTimestamp(),
                    tasks: parsedTasks
                });

                uploadBtnText.textContent = 'Upload & Sync Excel';
                e.target.value = '';
            } else {
                alert('The selected Excel sheet appears to be empty.');
                uploadBtnText.textContent = 'Upload & Sync Excel';
            }
        } catch (err) {
            console.error(err);
            alert('Error saving to Firebase: ' + err.message + '\n\nTip: Ensure Firestore Database is created in Firebase Console and rules allow read/write.');
            setSyncBadge('error', 'Sync Failed');
            uploadBtnText.textContent = 'Upload & Sync Excel';
        }
    };
    reader.readAsArrayBuffer(file);
}

function formatDateStr(val) {
    if (!val) return '';
    if (typeof val === 'number') {
        const date = new Date(Math.round((val - 25569) * 86400 * 1000));
        return date.toISOString().substring(0, 10);
    }
    return String(val).substring(0, 10);
}

function populateMonthDropdown() {
    const select = document.getElementById('monthFilterSelect');
    const currentVal = select.value;
    select.innerHTML = '<option value="ALL">All Time Records</option>';

    const monthSet = new Set();
    allTasks.forEach(t => {
        const ym = String(t['Created Date'] || '').substring(0, 7);
        if (ym && ym.length === 7) monthSet.add(ym);
    });

    const sortedMonths = Array.from(monthSet).sort().reverse();
    sortedMonths.forEach(ym => {
        const opt = document.createElement('option');
        opt.value = ym;
        const [year, month] = ym.split('-');
        const monthName = new Date(year, month - 1).toLocaleString('default', { month: 'short', year: 'numeric' });
        opt.textContent = `${monthName} (${ym})`;
        select.appendChild(opt);
    });

    if (sortedMonths.includes(currentVal)) {
        select.value = currentVal;
    }
}

function populateAssigneeDropdown() {
    const select = document.getElementById('tableAssigneeFilter');
    const currentVal = select.value;
    select.innerHTML = '<option value="">All Assignees</option>';
    const assignees = new Set();
    allTasks.forEach(t => {
        const names = String(t['Assigned To'] || '').split(';');
        names.forEach(n => {
            const trimmed = n.trim();
            if (trimmed) assignees.add(trimmed);
        });
    });
    Array.from(assignees).sort().forEach(a => {
        const opt = document.createElement('option');
        opt.value = a;
        opt.textContent = a;
        select.appendChild(opt);
    });
    if (assignees.has(currentVal)) {
        select.value = currentVal;
    }
}

function resetMonthFilter() {
    document.getElementById('monthFilterSelect').value = 'ALL';
    applyMonthFilter();
}

function applyMonthFilter() {
    const selectedMonth = document.getElementById('monthFilterSelect').value;

    if (selectedMonth === 'ALL') {
        filteredTasks = [...allTasks];
        document.getElementById('filterActiveLabel').textContent = `Showing all ${allTasks.length} tasks`;
    } else {
        filteredTasks = allTasks.filter(t => String(t['Created Date'] || '').startsWith(selectedMonth));
        document.getElementById('filterActiveLabel').textContent = `Filtered to ${selectedMonth} (${filteredTasks.length} tasks)`;
    }

    updateTopKPIs(selectedMonth);
    updateModuleLineChart(selectedMonth);
    updatePriorityChart();
    updateAssigneeDonutChart();

    tablePage = 1;
    renderTable();
}

function updateTopKPIs(selectedMonth) {
    const total = filteredTasks.length;
    const completed = filteredTasks.filter(t => t['Status'] === 'Completed').length;
    const pending = filteredTasks.filter(t => t['Status'] === 'In progress' || t['Status'] !== 'Completed').length;

    const completeRate = total > 0 ? ((completed / total) * 100).toFixed(1) : 0;
    const pendingRate = total > 0 ? ((pending / total) * 100).toFixed(1) : 0;

    document.getElementById('kpiTotalTask').textContent = total;
    document.getElementById('kpiResolveCount').textContent = completed;
    document.getElementById('kpiResolvePctBadge').textContent = `${completeRate}%`;
    document.getElementById('kpiResolveBar').style.width = `${completeRate}%`;

    document.getElementById('kpiPendingCount').textContent = pending;
    document.getElementById('kpiPendingPctBadge').textContent = `${pendingRate}%`;
    document.getElementById('kpiPendingBar').style.width = `${pendingRate}%`;

    const badge = document.getElementById('totalVsLastMonthBadge');
    const desc = document.getElementById('kpiTotalCompareDesc');

    if (total === 0) {
        badge.textContent = '--';
        desc.textContent = 'No tasks in selected period';
        return;
    }

    if (selectedMonth === 'ALL') {
        badge.className = 'text-[11px] font-semibold px-2 py-0.5 rounded-md bg-blue-50 text-blue-700';
        badge.textContent = '100% volume';
        desc.textContent = 'All time volume';
    } else {
        const [curY, curM] = selectedMonth.split('-').map(Number);
        let prevY = curY;
        let prevM = curM - 1;
        if (prevM === 0) { prevM = 12; prevY -= 1; }
        const prevMonthStr = `${prevY}-${String(prevM).padStart(2, '0')}`;

        const lastMonthCount = allTasks.filter(t => String(t['Created Date'] || '').startsWith(prevMonthStr)).length;

        if (lastMonthCount > 0) {
            const diffPct = (((total - lastMonthCount) / lastMonthCount) * 100).toFixed(1);
            if (diffPct >= 0) {
                badge.className = 'text-[11px] font-bold px-2 py-0.5 rounded-md bg-emerald-50 text-emerald-700';
                badge.textContent = `+${diffPct}% vs last month`;
            } else {
                badge.className = 'text-[11px] font-bold px-2 py-0.5 rounded-md bg-rose-50 text-rose-700';
                badge.textContent = `${diffPct}% vs last month`;
            }
            desc.textContent = `vs ${prevMonthStr} (${lastMonthCount} tasks)`;
        } else {
            badge.className = 'text-[11px] font-semibold px-2 py-0.5 rounded-md bg-slate-100 text-slate-600';
            badge.textContent = 'No prior data';
            desc.textContent = 'Prior month has 0 recorded tasks';
        }
    }
}

function initCharts() {
    const ctx1 = document.getElementById('workloadModuleChart');
    moduleChart = new Chart(ctx1, {
        type: 'line',
        data: { labels: [], datasets: [] },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            layout: { padding: { top: 25, right: 15, left: 10, bottom: 5 } },
            scales: {
                x: { grid: { display: false }, border: { display: false } },
                y: { grid: { display: false }, border: { display: false }, beginAtZero: true, grace: '15%' }
            },
            plugins: {
                legend: { display: false },
                datalabels: {
                    anchor: 'end',
                    align: 'top',
                    offset: 4,
                    color: '#1e293b',
                    font: { weight: 'bold', size: 11 },
                    formatter: (value) => value > 0 ? value : ''
                }
            }
        }
    });

    const ctx2 = document.getElementById('priorityChart');
    priorityDoughnut = new Chart(ctx2, {
        type: 'doughnut',
        data: {
            labels: ['Medium', 'Important', 'Urgent'],
            datasets: [{
                data: [0, 0, 0],
                backgroundColor: ['#94a3b8', '#fb923c', '#f43f5e'],
                borderWidth: 2,
                borderColor: '#ffffff'
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            cutout: '68%',
            plugins: {
                legend: { display: false },
                datalabels: {
                    color: '#ffffff',
                    font: { weight: 'bold', size: 10 },
                    formatter: (value, ctx) => {
                        const total = ctx.dataset.data.reduce((a, b) => a + b, 0);
                        if (value === 0 || total === 0) return '';
                        const pct = ((value / total) * 100).toFixed(0);
                        return pct > 5 ? `${pct}%` : '';
                    }
                }
            }
        }
    });

    const ctx3 = document.getElementById('assigneeDonutChart');
    assigneeDonutChart = new Chart(ctx3, {
        type: 'doughnut',
        data: {
            labels: [],
            datasets: [{
                data: [],
                backgroundColor: ['#F37021', '#0284c7', '#10b981', '#a855f7', '#64748b'],
                borderWidth: 2,
                borderColor: '#ffffff'
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            cutout: '62%',
            plugins: {
                legend: { display: false },
                datalabels: {
                    color: '#ffffff',
                    font: { weight: 'bold', size: 10 },
                    formatter: (value) => value > 0 ? value : ''
                }
            }
        }
    });
}

function updateModuleLineChart(selectedMonth) {
    const mainModules = ['KADS-Service', 'KADS-Sale', 'Reset MFA', 'K-POS', 'KADS-Sparepart', 'SKC-D', 'Other'];
    const counts = [0, 0, 0, 0, 0, 0, 0];

    filteredTasks.forEach(task => {
        const label = task['Labels'] || 'Other';
        let matchedIdx = mainModules.findIndex(m => label.includes(m));
        if (matchedIdx === -1) matchedIdx = 6;
        counts[matchedIdx]++;
    });

    const activeIndices = mainModules.map((m, i) => (counts[i] > 0 ? i : -1)).filter(i => i !== -1);
    const displayLabels = activeIndices.map(i => mainModules[i]);
    const displayCounts = activeIndices.map(i => counts[i]);

    moduleChart.data = {
        labels: displayLabels,
        datasets: [{
            label: 'Total Cases',
            data: displayCounts,
            borderColor: '#F37021',
            backgroundColor: 'rgba(243, 112, 33, 0.12)',
            pointBackgroundColor: '#F37021',
            pointBorderColor: '#ffffff',
            pointBorderWidth: 2,
            pointRadius: 5,
            pointHoverRadius: 7,
            tension: 0.3,
            borderWidth: 2.5,
            fill: true
        }]
    };
    moduleChart.update();

    document.getElementById('moduleChartSubtitle').textContent = selectedMonth === 'ALL'
        ? 'Total cases per main module (All Time)'
        : `Total cases per main module (${selectedMonth})`;
}

function updatePriorityChart() {
    const total = filteredTasks.length || 1;
    const mediumCount = filteredTasks.filter(t => t['Priority'] === 'Medium' || !t['Priority']).length;
    const importantCount = filteredTasks.filter(t => t['Priority'] === 'Important').length;
    const urgentCount = filteredTasks.filter(t => t['Priority'] === 'Urgent').length;

    const mPct = filteredTasks.length > 0 ? ((mediumCount / total) * 100).toFixed(1) : 0;
    const iPct = filteredTasks.length > 0 ? ((importantCount / total) * 100).toFixed(1) : 0;
    const uPct = filteredTasks.length > 0 ? ((urgentCount / total) * 100).toFixed(1) : 0;

    document.getElementById('priorityMediumPct').textContent = `${mPct}%`;
    document.getElementById('priorityImportantPct').textContent = `${iPct}%`;
    document.getElementById('priorityUrgentPct').textContent = `${uPct}%`;

    priorityDoughnut.data.datasets[0].data = [mediumCount, importantCount, urgentCount];
    priorityDoughnut.update();
}

function updateAssigneeDonutChart() {
    const personCounts = {};

    filteredTasks.forEach(task => {
        const assignedField = String(task['Assigned To'] || '').trim();
        if (!assignedField) return;

        const names = assignedField.split(';').map(n => n.trim()).filter(n => n);
        names.forEach(name => {
            personCounts[name] = (personCounts[name] || 0) + 1;
        });
    });

    const sortedEntries = Object.entries(personCounts).sort((a, b) => b[1] - a[1]);
    const displayLabels = sortedEntries.map(e => e[0].split(';')[0]);
    const displayCounts = sortedEntries.map(e => e[1]);

    assigneeDonutChart.data = {
        labels: displayLabels,
        datasets: [{
            data: displayCounts,
            backgroundColor: ['#F37021', '#0284c7', '#10b981', '#a855f7', '#64748b'],
            borderWidth: 2,
            borderColor: '#ffffff'
        }]
    };
    assigneeDonutChart.update();

    const legendContainer = document.getElementById('assigneeLegendSummary');
    legendContainer.innerHTML = '';
    if (sortedEntries.length === 0) {
        legendContainer.innerHTML = '<span class="text-slate-400 text-xs font-normal">No data loaded</span>';
        return;
    }

    const palette = ['#F37021', '#0284c7', '#10b981', '#a855f7', '#64748b'];
    sortedEntries.forEach((e, idx) => {
        const item = document.createElement('div');
        item.className = 'flex items-center gap-1';
        item.innerHTML = `<span class="w-2.5 h-2.5 rounded-full" style="background-color: ${palette[idx] || '#64748b'}"></span> <span>${e[0].split(' ')[0]}: <b>${e[1]}</b></span>`;
        legendContainer.appendChild(item);
    });
}

function getPriorityBadge(priority) {
    if (priority === 'Urgent') return '<span class="px-2 py-0.5 font-bold text-[10px] rounded-full bg-rose-100 text-rose-700">Urgent</span>';
    if (priority === 'Important') return '<span class="px-2 py-0.5 font-bold text-[10px] rounded-full bg-orange-100 text-orange-700">Important</span>';
    return '<span class="px-2 py-0.5 font-medium text-[10px] rounded-full bg-slate-100 text-slate-600">Medium</span>';
}

function getStatusBadge(status) {
    if (status === 'Completed') return '<span class="inline-flex items-center gap-1 text-emerald-700 font-semibold bg-emerald-50 px-2 py-0.5 rounded-full border border-emerald-200"><span class="w-1.5 h-1.5 rounded-full bg-emerald-500"></span> Resolved</span>';
    return '<span class="inline-flex items-center gap-1 text-rose-700 font-bold bg-rose-50 px-2 py-0.5 rounded-full border border-rose-200"><span class="w-1.5 h-1.5 rounded-full bg-rose-500 animate-pulse"></span> Pending</span>';
}

function escapeHtml(text) {
    if (!text) return '';
    return String(text).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

function renderTable() {
    const searchVal = (document.getElementById('tableSearchInput').value || '').toLowerCase().trim();
    const viewMode = document.getElementById('tableViewMode').value;
    const assigneeFilter = document.getElementById('tableAssigneeFilter').value;

    const tableData = filteredTasks.filter(task => {
        const taskName = String(task['Task Name'] || '').toLowerCase();
        const taskNotes = String(task['Notes'] || '').toLowerCase();
        const matchSearch = !searchVal || taskName.includes(searchVal) || taskNotes.includes(searchVal);
        const matchAssignee = !assigneeFilter || String(task['Assigned To'] || '').includes(assigneeFilter);

        let matchStatus = true;
        if (viewMode === 'PendingOnly') {
            matchStatus = (task['Status'] !== 'Completed');
        } else if (viewMode === 'Completed') {
            matchStatus = (task['Status'] === 'Completed');
        }

        return matchSearch && matchAssignee && matchStatus;
    });

    const tbody = document.getElementById('pendingTableBody');
    const total = tableData.length;
    const totalPages = Math.ceil(total / pageSize) || 1;

    if (tablePage > totalPages) tablePage = totalPages;
    const start = (tablePage - 1) * pageSize;
    const end = Math.min(start + pageSize, total);
    const pageRecords = tableData.slice(start, end);

    tbody.innerHTML = '';

    if (pageRecords.length === 0) {
        tbody.innerHTML = '<tr><td colspan="9" class="text-center py-10 text-slate-400 italic">No tasks found matching current filter criteria.</td></tr>';
    } else {
        pageRecords.forEach(task => {
            const tr = document.createElement('tr');
            tr.className = 'hover:bg-slate-50 transition';

            const taskNameEsc = escapeHtml(task['Task Name']);
            const labelsEsc = escapeHtml(task['Labels'] || 'None');
            const assignedEsc = escapeHtml(task['Assigned To']);
            const notesEsc = escapeHtml(task['Notes'] || '-');
            const createdDate = task['Created Date'] ? String(task['Created Date']).substring(0, 10) : '-';
            const dueDate = task['Due date'] ? String(task['Due date']).substring(0, 10) : '-';
            const compDate = task['Completed Date'] ? String(task['Completed Date']).substring(0, 10) : '-';

            tr.innerHTML = `
                <td class="py-2.5 px-4 font-medium text-slate-900 max-w-xs">
                    <div class="truncate" title="${taskNameEsc}">${taskNameEsc}</div>
                </td>
                <td class="py-2.5 px-4 text-slate-500 whitespace-nowrap">${createdDate}</td>
                <td class="py-2.5 px-4">
                    <span class="px-2 py-0.5 rounded bg-slate-100 text-slate-700 font-semibold text-[10px]">${labelsEsc}</span>
                </td>
                <td class="py-2.5 px-4">${getStatusBadge(task['Status'])}</td>
                <td class="py-2.5 px-4">${getPriorityBadge(task['Priority'])}</td>
                <td class="py-2.5 px-4 text-slate-700">${assignedEsc}</td>
                <td class="py-2.5 px-4 text-slate-500">${dueDate}</td>
                <td class="py-2.5 px-4 text-slate-500">${compDate}</td>
                <td class="py-2.5 px-4 text-slate-400 max-w-[200px]">
                    <div class="truncate" title="${notesEsc}">${notesEsc}</div>
                </td>
            `;
            tbody.appendChild(tr);
        });
    }

    document.getElementById('tablePaginationInfo').textContent = `Showing ${total === 0 ? 0 : start + 1} to ${end} of ${total} tasks`;

    const paginationContainer = document.getElementById('tablePaginationControls');
    paginationContainer.innerHTML = `
        <button id="btnTblPrev" ${tablePage === 1 ? 'disabled' : ''} class="px-2.5 py-1 rounded border border-slate-200 text-slate-600 hover:bg-slate-100 disabled:opacity-40 disabled:cursor-not-allowed text-xs">Prev</button>
        <span class="px-2 font-semibold text-slate-700 text-xs">Page ${tablePage} / ${totalPages}</span>
        <button id="btnTblNext" ${tablePage === totalPages ? 'disabled' : ''} class="px-2.5 py-1 rounded border border-slate-200 text-slate-600 hover:bg-slate-100 disabled:opacity-40 disabled:cursor-not-allowed text-xs">Next</button>
    `;

    document.getElementById('btnTblPrev').onclick = () => { if (tablePage > 1) { tablePage--; renderTable(); } };
    document.getElementById('btnTblNext').onclick = () => { if (tablePage < totalPages) { tablePage++; renderTable(); } };
}