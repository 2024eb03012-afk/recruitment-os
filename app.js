// ===== Configuration =====
const CONFIG = {
    webhookUrl: 'https://n8n.smallgrp.com/webhook-test/3f956aa9-6f69-436a-952a-9aa80df55740',
    // Google Sheet URL - Sheet 3 specific
    googleSheetUrl: 'https://docs.google.com/spreadsheets/d/1DAVbKlf0bI3Dkmt8YQlAawoG7Oeez4px6vHhWy7Ts8w/gviz/tq?tqx=out:csv&sheet=Sheet3',
    // Auto-refresh interval in milliseconds (30 seconds)
    refreshInterval: 30000,
    emailWebhookUrl: 'https://n8n.zenquill.tech/webhook-test/f21f355b-e854-48c6-a58f-2de26c242998'
};

// Column mapping from Google Sheet headers to display
const COLUMN_MAPPING = {
    'Company name': 'companyName',
    'Job type': 'jobType',
    'City': 'city',
    'JD': 'jd',
    'Company job url': 'companyJobUrl',
    'Salary': 'salary',
    'Company descriptions': 'companyDescription',
    'Title': 'title',
    'Match score analysis': 'matchScore',
    'Website': 'website',
    'Decision maker email': 'decisionMakerEmail',
    'Outreach email text': 'outreachEmailText'
};

// ===== DOM Elements =====
const elements = {
    form: document.getElementById('scrapingForm'),
    submitBtn: document.getElementById('submitBtn'),
    formMessage: document.getElementById('formMessage'),
    linkedinKeywordsGroup: document.getElementById('linkedinKeywordsGroup'),
    linkedinKeywords: document.getElementById('linkedinKeywords'),
    platformLinkedInPost: document.getElementById('platformLinkedInPost'),
    analyticsGrid: document.getElementById('analyticsGrid'),
    analyticsError: document.getElementById('analyticsError'),
    tableBody: document.getElementById('tableBody'),
    tableCount: document.getElementById('tableCount'),
    emptyState: document.getElementById('emptyState'),
    tableError: document.getElementById('tableError'),
    refreshBtn: document.getElementById('refreshBtn'),
    textModal: document.getElementById('textModal'),
    modalTitle: document.getElementById('modalTitle'),
    modalBodyText: document.getElementById('modalBodyText'),
    closeModal: document.getElementById('closeModal'),
    modalMeta: document.getElementById('modalMeta'),
    editBtn: document.getElementById('editBtn'),
    modalEditText: document.getElementById('modalEditText'),
    modalFooter: document.getElementById('modalFooter'),
    saveBtn: document.getElementById('saveBtn'),
    cancelBtn: document.getElementById('cancelBtn')
};

// ===== Global State =====
let sheetData = [];
let autoRefreshTimer = null;
let localEdits = {}; // Persistent local overrides during session
let currentModalContext = {
    dataIndex: -1,
    columnKey: '',
    isEditing: false
};
let chartInstances = {}; // Track Chart.js instances

// ===== Initialization =====
document.addEventListener('DOMContentLoaded', () => {
    initForm();
    loadSheetData();
    setupEventListeners();
    startAutoRefresh();
});

// ===== Event Listeners =====
function setupEventListeners() {
    elements.refreshBtn.addEventListener('click', manualRefresh);
    elements.platformLinkedInPost.addEventListener('change', toggleLinkedInKeywords);

    // Global click handler for expandable text trigger
    elements.tableBody.addEventListener('click', (e) => {
        const expandable = e.target.closest('.expandable-text');
        if (expandable) {
            const row = expandable.closest('tr');
            const cell = expandable.closest('td');
            const companyName = row ? row.cells[0].textContent : 'Details';
            const fullText = expandable.querySelector('.text-content').textContent;

            // Identify column index to determine if it's the outreach email
            const columnIndex = Array.from(row.cells).indexOf(cell);
            const columnKeys = ['companyName', 'jobType', 'city', 'jd', 'companyJobUrl', 'salary', 'companyDescription', 'title', 'matchScore', 'website', 'decisionMakerEmail', 'outreachEmailText'];
            const columnKey = columnKeys[columnIndex];

            // Find the original index in sheetData
            // Since we reversed the data for the table, we need to find the match
            const dataIndex = sheetData.findIndex(item => item.companyName === companyName && (item.jd === fullText || item.companyDescription === fullText || item.outreachEmailText === fullText));

            openModal(companyName, fullText, {
                dataIndex,
                columnKey,
                columnLabel: row.closest('table').querySelectorAll('th')[columnIndex].textContent
            });
        }

        // Handle Send Email button click
        const sendBtn = e.target.closest('.btn-send-email');
        if (sendBtn) {
            e.stopPropagation(); // Prevent expandable text click if any
            const row = sendBtn.closest('tr');
            handleSendEmail(row);
        }
    });

    // Modal events
    elements.editBtn.addEventListener('click', () => toggleEditMode(true));
    elements.cancelBtn.addEventListener('click', () => toggleEditMode(false));
    elements.saveBtn.addEventListener('click', saveChanges);
    elements.closeModal.addEventListener('click', closeModal);

    elements.textModal.addEventListener('click', (e) => {
        if (e.target === elements.textModal) closeModal();
    });

    // Close on Escape key
    document.addEventListener('keydown', (e) => {
        if (e.key === 'Escape' && elements.textModal.classList.contains('active')) {
            closeModal();
        }
    });
}

function toggleLinkedInKeywords() {
    const isChecked = elements.platformLinkedInPost.checked;
    elements.linkedinKeywordsGroup.style.display = isChecked ? 'block' : 'none';
    elements.linkedinKeywords.required = isChecked;

    if (isChecked) {
        elements.linkedinKeywordsGroup.style.animation = 'slideIn 0.3s ease';
    }
}

// ===== Auto Refresh =====
function startAutoRefresh() {
    // Clear any existing timer
    if (autoRefreshTimer) {
        clearInterval(autoRefreshTimer);
    }

    // Set up auto-refresh every 30 seconds
    autoRefreshTimer = setInterval(() => {
        console.log('Auto-refreshing data...');
        loadSheetData();
    }, CONFIG.refreshInterval);
}

function manualRefresh() {
    elements.refreshBtn.style.transform = 'rotate(360deg)';
    loadSheetData();

    setTimeout(() => {
        elements.refreshBtn.style.transform = '';
    }, 500);
}

// ===== Form Handling =====
function initForm() {
    elements.form.addEventListener('submit', handleFormSubmit);
    elements.form.addEventListener('reset', () => {
        hideMessage();
        elements.linkedinKeywordsGroup.style.display = 'none';
        elements.linkedinKeywords.required = false;
    });
}

async function handleFormSubmit(e) {
    e.preventDefault();

    const formData = collectFormData();

    if (!validateForm(formData)) return;

    setLoadingState(true);
    hideMessage();

    try {
        const response = await fetch(CONFIG.webhookUrl, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(formData)
        });

        if (!response.ok) throw new Error(`HTTP error! status: ${response.status}`);

        showMessage('success', 'Scraping request submitted successfully. Data will sync automatically from Google Sheet.');
        elements.form.reset();
        elements.linkedinKeywordsGroup.style.display = 'none';

        // Refresh data after successful submission
        setTimeout(() => {
            loadSheetData();
        }, 3000);

    } catch (error) {
        console.error('Form submission error:', error);
        showMessage('error', 'Failed to submit request. Please check your connection and try again.');
    } finally {
        setLoadingState(false);
    }
}

function collectFormData() {
    const formData = new FormData(elements.form);

    return {
        jobTitle: formData.get('jobTitle'),
        jobTypes: formData.getAll('jobType'),
        location: formData.get('location'),
        country: formData.get('country'),
        numJobs: parseInt(formData.get('numJobs')),
        platforms: formData.getAll('platform'),
        linkedinKeywords: formData.get('linkedinKeywords') || null,
        companySize: formData.getAll('companySize'),
        salaryRange: {
            min: formData.get('salaryMin') ? parseInt(formData.get('salaryMin')) : null,
            max: formData.get('salaryMax') ? parseInt(formData.get('salaryMax')) : null
        },
        timestamp: new Date().toISOString()
    };
}

function validateForm(data) {
    if (!data.jobTitle.trim()) {
        showMessage('error', 'Please enter a job title.');
        return false;
    }

    if (!data.numJobs || data.numJobs < 1) {
        showMessage('error', 'Please enter a valid number of jobs to scrape.');
        return false;
    }

    if (data.platforms.includes('linkedin_post') && !data.linkedinKeywords?.trim()) {
        showMessage('error', 'Please enter job searching keywords for LinkedIn Post.');
        return false;
    }

    return true;
}

function setLoadingState(loading) {
    const btnText = elements.submitBtn.querySelector('.btn-text');
    const btnLoader = elements.submitBtn.querySelector('.btn-loader');

    elements.submitBtn.disabled = loading;
    btnText.style.display = loading ? 'none' : 'inline-flex';
    btnLoader.style.display = loading ? 'inline-flex' : 'none';
}

function showMessage(type, message) {
    elements.formMessage.className = `form-message ${type}`;
    elements.formMessage.textContent = message;
    elements.formMessage.style.display = 'block';
}

function hideMessage() {
    elements.formMessage.style.display = 'none';
}

// ===== Google Sheet Data Loading =====
async function loadSheetData() {
    showTableSkeleton();
    showAnalyticsSkeleton();
    elements.emptyState.style.display = 'none';
    elements.tableError.style.display = 'none';
    elements.analyticsError.style.display = 'none';

    try {
        const data = await fetchSheetData();

        // Merge with local edits
        sheetData = data.map(item => {
            const editKey = `${item.companyName}|${item.title}`;
            if (localEdits[editKey]) {
                return { ...item, ...localEdits[editKey] };
            }
            return item;
        });

        // Render analytics and charts
        renderAnalytics(sheetData);
        renderCharts(sheetData);

        // Render table with all data (newest first)
        renderTable([...sheetData].reverse());

    } catch (error) {
        console.error('Sheet data load error:', error);
        elements.tableError.style.display = 'block';
        elements.analyticsError.style.display = 'block';
        elements.tableBody.innerHTML = '';
        elements.analyticsGrid.innerHTML = '';
    }
}

async function fetchSheetData() {
    const response = await fetch(CONFIG.googleSheetUrl);

    if (!response.ok) {
        throw new Error(`Failed to fetch sheet data: ${response.status}`);
    }

    const csvText = await response.text();
    return parseCSV(csvText);
}

function parseCSV(csvText) {
    const rows = [];
    let currentRow = [];
    let currentField = '';
    let inQuotes = false;

    // Remove any potential UTF-8 BOM
    const content = csvText.replace(/^\uFEFF/, '');

    for (let i = 0; i < content.length; i++) {
        const char = content[i];
        const nextChar = content[i + 1];

        if (char === '"') {
            if (inQuotes && nextChar === '"') {
                currentField += '"';
                i++;
            } else {
                inQuotes = !inQuotes;
            }
        } else if (char === ',' && !inQuotes) {
            currentRow.push(currentField);
            currentField = '';
        } else if ((char === '\r' || char === '\n') && !inQuotes) {
            if (currentRow.length > 0 || currentField !== '') {
                currentRow.push(currentField);
                rows.push(currentRow);
                currentRow = [];
                currentField = '';
            }
            if (char === '\r' && nextChar === '\n') {
                i++;
            }
        } else {
            currentField += char;
        }
    }

    // Add last row if exists
    if (currentRow.length > 0 || currentField !== '') {
        currentRow.push(currentField);
        rows.push(currentRow);
    }

    if (rows.length < 2) return [];

    const headers = rows[0].map(header => header.trim().replace(/\s+/g, ' '));
    const data = [];

    for (let i = 1; i < rows.length; i++) {
        const values = rows[i];
        const row = {};

        headers.forEach((header, index) => {
            const mappedKey = COLUMN_MAPPING[header] || header;
            // Clean up the value (remove extra quotes if any and trim)
            let value = values[index] ? values[index].trim() : '';

            // Handle some common CSV artifacts like escaped newlines
            value = value.replace(/\\n/g, '\n').replace(/\\r/g, '\r');

            row[mappedKey] = value;
        });

        if (Object.values(row).some(val => val !== '')) {
            data.push(row);
        }
    }

    return data;
}

// ===== Analytics Rendering =====
function showAnalyticsSkeleton() {
    const skeletons = Array(6).fill(`
        <div class="stat-card glass-card">
            <div class="stat-content">
                <span class="stat-label"><span class="skeleton-loader" style="width: 100px;"></span></span>
                <span class="stat-value"><span class="skeleton-loader"></span></span>
            </div>
        </div>
    `).join('');
    elements.analyticsGrid.innerHTML = skeletons;
}

function renderAnalytics(data) {
    if (!data || data.length === 0) {
        elements.analyticsGrid.innerHTML = '<p class="empty-msg">No data available for analytics.</p>';
        return;
    }

    const metrics = calculateMetrics(data);

    elements.analyticsGrid.innerHTML = `
        <div class="stat-card glass-card">
            <div class="stat-icon total"></div>
            <div class="stat-content">
                <span class="stat-label">Total Jobs Scraped</span>
                <span class="stat-value">${formatNumber(metrics.totalJobs)}</span>
            </div>
        </div>
        <div class="stat-card glass-card">
            <div class="stat-icon companies"></div>
            <div class="stat-content">
                <span class="stat-label">Unique Companies</span>
                <span class="stat-value">${formatNumber(metrics.uniqueCompanies)}</span>
            </div>
        </div>
        <div class="stat-card glass-card">
            <div class="stat-icon types"></div>
            <div class="stat-content">
                <span class="stat-label">Job Types</span>
                <span class="stat-value">${metrics.topJobType}</span>
            </div>
        </div>
        <div class="stat-card glass-card">
            <div class="stat-icon cities"></div>
            <div class="stat-content">
                <span class="stat-label">Cities Covered</span>
                <span class="stat-value">${formatNumber(metrics.uniqueCities)}</span>
            </div>
        </div>
        <div class="stat-card glass-card">
            <div class="stat-icon score"></div>
            <div class="stat-content">
                <span class="stat-label">Avg Match Score</span>
                <span class="stat-value">${metrics.avgMatchScore}%</span>
            </div>
        </div>
        <div class="stat-card glass-card">
            <div class="stat-icon leads"></div>
            <div class="stat-content">
                <span class="stat-label">Outreach Ready</span>
                <span class="stat-value">${formatNumber(metrics.outreachReady)}</span>
            </div>
        </div>
    `;
}

function calculateMetrics(data) {
    const uniqueCompanies = new Set(data.map(item => item.companyName).filter(Boolean));
    const uniqueCities = new Set(data.map(item => item.city).filter(Boolean));

    // Job Types
    const jobTypeCounts = data.reduce((acc, item) => {
        const type = item.jobType || 'Other';
        acc[type] = (acc[type] || 0) + 1;
        return acc;
    }, {});
    const topJobType = Object.entries(jobTypeCounts)
        .sort((a, b) => b[1] - a[1])[0]?.[0] || 'N/A';

    // Match Score
    const scores = data.map(item => parseFloat(item.matchScore)).filter(s => !isNaN(s));
    const avgMatchScore = scores.length > 0
        ? Math.round(scores.reduce((a, b) => a + b, 0) / scores.length)
        : 0;

    // Outreach Ready
    const outreachReady = data.filter(item => item.decisionMakerEmail && item.decisionMakerEmail.trim() !== '').length;

    return {
        totalJobs: data.length,
        uniqueCompanies: uniqueCompanies.size,
        topJobType,
        uniqueCities: uniqueCities.size,
        avgMatchScore,
        outreachReady
    };
}

// ===== Charts Rendering =====
function renderCharts(data) {
    if (!data || data.length === 0) return;

    // Destroy existing charts
    Object.values(chartInstances).forEach(chart => chart.destroy());

    // 1. Jobs by Company (Top 10)
    const companyCounts = data.reduce((acc, item) => {
        acc[item.companyName] = (acc[item.companyName] || 0) + 1;
        return acc;
    }, {});
    const topCompanies = Object.entries(companyCounts)
        .sort((a, b) => b[1] - a[1])
        .slice(0, 10);

    createBarChart('companyChart',
        topCompanies.map(c => c[0]),
        topCompanies.map(c => c[1]),
        '#6366f1'
    );

    // 2. Jobs by Job Type
    const jobTypeCounts = data.reduce((acc, item) => {
        const type = item.jobType || 'Other';
        acc[type] = (acc[type] || 0) + 1;
        return acc;
    }, {});

    createBarChart('jobTypeChart',
        Object.keys(jobTypeCounts),
        Object.values(jobTypeCounts),
        '#f59e0b'
    );

    // 3. Jobs by City (Top 10)
    const cityCounts = data.reduce((acc, item) => {
        let city = (item.city || 'Remote/Unknown').trim();
        // Capitalize each word for uniform grouping
        city = city.split(' ').map(w => w.charAt(0).toUpperCase() + w.slice(1).toLowerCase()).join(' ');
        acc[city] = (acc[city] || 0) + 1;
        return acc;
    }, {});
    const topCities = Object.entries(cityCounts)
        .sort((a, b) => b[1] - a[1])
        .slice(0, 10);

    createBarChart('cityChart',
        topCities.map(c => c[0]),
        topCities.map(c => c[1]),
        '#10b981'
    );


    // 4. Match Score Distribution
    const scoreBuckets = { '0-60': 0, '61-75': 0, '76-85': 0, '86-100': 0 };
    data.forEach(item => {
        const score = parseFloat(item.matchScore);
        if (isNaN(score)) return;
        if (score <= 60) scoreBuckets['0-60']++;
        else if (score <= 75) scoreBuckets['61-75']++;
        else if (score <= 85) scoreBuckets['76-85']++;
        else scoreBuckets['86-100']++;
    });

    createBarChart('matchScoreChart',
        Object.keys(scoreBuckets),
        Object.values(scoreBuckets),
        '#8b5cf6'
    );

    // 5. Outreach Coverage
    const coverage = { 'With Email': 0, 'Without Email': 0 };
    data.forEach(item => {
        if (item.decisionMakerEmail && item.decisionMakerEmail.trim() !== '') coverage['With Email']++;
        else coverage['Without Email']++;
    });

    createBarChart('outreachChart',
        Object.keys(coverage),
        Object.values(coverage),
        '#ef4444'
    );

    // 6. Role Category Split
    const categories = {
        'SEO': 0,
        'Software Engineer': 0,
        'Backend': 0,
        'Founding Engineer': 0,
        'Other': 0
    };
    data.forEach(item => {
        const title = (item.title || '').toLowerCase();
        let categorized = false;

        if (title.includes('seo')) { categories['SEO']++; categorized = true; }
        else if (title.includes('software engineer') || title.includes('swe')) { categories['Software Engineer']++; categorized = true; }
        else if (title.includes('backend')) { categories['Backend']++; categorized = true; }
        else if (title.includes('founding engineer')) { categories['Founding Engineer']++; categorized = true; }

        if (!categorized) categories['Other']++;
    });

    createBarChart('roleChart',
        Object.keys(categories),
        Object.values(categories),
        '#06b6d4'
    );
}

function createBarChart(canvasId, labels, data, color) {
    const ctx = document.getElementById(canvasId).getContext('2d');

    chartInstances[canvasId] = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: labels,
            datasets: [{
                label: 'Count',
                data: data,
                backgroundColor: color,
                borderRadius: 6,
                borderSkipped: false,
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { display: false },
                tooltip: {
                    backgroundColor: '#1e293b',
                    padding: 12,
                    titleFont: { size: 14, weight: 'bold' },
                    bodyFont: { size: 13 },
                    displayColors: false
                }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    grid: { color: '#f1f5f9' },
                    ticks: { stepSize: 1, color: '#64748b' }
                },
                x: {
                    grid: { display: false },
                    ticks: { color: '#64748b' }
                }
            }
        }
    });
}

function formatNumber(num) {
    return new Intl.NumberFormat().format(num);
}

// ===== Table Rendering =====
function showTableSkeleton() {
    const skeletonRows = Array(5).fill().map(() => `
        <tr class="skeleton-row">
            ${Array(12).fill('<td><span class="skeleton-loader"></span></td>').join('')}
        </tr>
    `).join('');

    elements.tableBody.innerHTML = skeletonRows;
}

function renderTable(data) {
    elements.tableError.style.display = 'none';

    if (!data || data.length === 0) {
        elements.tableBody.innerHTML = '';
        elements.emptyState.style.display = 'block';
        elements.tableCount.textContent = '0 records';
        return;
    }

    elements.emptyState.style.display = 'none';
    elements.tableCount.textContent = `${data.length} records`;

    elements.tableBody.innerHTML = data.map(row => `
        <tr>
            <td>${escapeHtml(row.companyName || '')}</td>
            <td>${escapeHtml(row.jobType || '')}</td>
            <td>${escapeHtml(row.city || '')}</td>
            <td class="wrap">${renderExpandableCell(row.jd || '', 100)}</td>
            <td>${formatUrl(row.companyJobUrl || '')}</td>
            <td>${escapeHtml(row.salary || '')}</td>
            <td class="wrap">${renderExpandableCell(row.companyDescription || '', 80)}</td>
            <td>${escapeHtml(row.title || '')}</td>
            <td>${escapeHtml(row.matchScore || '')}</td>
            <td>${formatUrl(row.website || '')}</td>
            <td>${escapeHtml(row.decisionMakerEmail || '')}</td>
            <td class="wrap outreach-cell">
                ${renderExpandableCell(row.outreachEmailText || '', 100)}
                <button class="btn-send-email" title="Send Email">Send</button>
            </td>
        </tr>
    `).join('');
}

function renderExpandableCell(text, maxLength) {
    if (!text) return '';
    const escapedText = escapeHtml(text);
    if (text.length <= maxLength) return escapedText;

    return `
        <div class="expandable-text" title="Click to view full text">
            <div class="text-content" style="display:none;">${escapedText}</div>
            <div class="truncated-text">${escapedText.substring(0, maxLength)}...</div>
            <div class="expand-hint">Click to expand</div>
        </div>
    `;
}

// ===== Modal Helpers =====
function openModal(title, text, context = {}) {
    currentModalContext = {
        dataIndex: context.dataIndex,
        columnKey: context.columnKey,
        isEditing: false
    };

    elements.modalTitle.textContent = title;
    elements.modalMeta.textContent = context.columnLabel || '';
    elements.modalBodyText.textContent = text;
    elements.modalEditText.value = text;

    // Only show edit button for outreach email text
    elements.editBtn.style.display = context.columnKey === 'outreachEmailText' ? 'flex' : 'none';

    // Reset edit mode
    toggleEditMode(false);

    elements.textModal.classList.add('active');
    document.body.style.overflow = 'hidden';
}

function toggleEditMode(editing) {
    currentModalContext.isEditing = editing;
    elements.modalBodyText.style.display = editing ? 'none' : 'block';
    elements.modalEditText.style.display = editing ? 'block' : 'none';
    elements.modalFooter.style.display = editing ? 'block' : 'none';
    elements.editBtn.style.display = (editing || currentModalContext.columnKey !== 'outreachEmailText') ? 'none' : 'flex';

    if (editing) {
        elements.modalEditText.focus();
    }
}

function closeModal() {
    if (currentModalContext.isEditing) {
        if (!confirm('You have unsaved changes. Are you sure you want to close?')) return;
    }
    elements.textModal.classList.remove('active');
    document.body.style.overflow = '';
    toggleEditMode(false);
}

async function saveChanges() {
    const newText = elements.modalEditText.value;
    const { dataIndex, columnKey } = currentModalContext;

    if (dataIndex !== -1 && columnKey) {
        // Update local state and tracking
        const item = sheetData[dataIndex];
        const editKey = `${item.companyName}|${item.title}`;

        if (!localEdits[editKey]) localEdits[editKey] = {};
        localEdits[editKey][columnKey] = newText;
        item[columnKey] = newText;

        // Success feedback
        elements.modalBodyText.textContent = newText;
        toggleEditMode(false);

        // Re-render table to reflect changes
        renderTable([...sheetData].reverse());

        console.log(`Saved changes for ${columnKey} at index ${dataIndex}`);
        // Note: In a production app, we would send a request to the backend/Google Script here
    }
}

// ===== Utility Functions =====
function escapeHtml(text) {
    if (!text) return '';
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

function truncateText(text, maxLength) {
    if (!text) return '';
    if (text.length <= maxLength) return escapeHtml(text);
    return escapeHtml(text.substring(0, maxLength)) + '...';
}

function formatUrl(url) {
    if (!url) return '';
    const displayUrl = url.length > 30 ? url.substring(0, 30) + '...' : url;
    return `<a href="${escapeHtml(url)}" target="_blank" rel="noopener noreferrer">${escapeHtml(displayUrl)}</a>`;
}

async function handleSendEmail(row) {
    const titleCell = row.cells[7]; // Title is at index 7
    const emailCell = row.cells[10]; // Decision maker email is at index 10
    const outreachCell = row.cells[11]; // Outreach text is at index 11

    const jobTitle = titleCell ? titleCell.textContent.trim() : '';
    const emailAddress = emailCell ? emailCell.textContent.trim() : '';
    const emailBody = outreachCell ? (outreachCell.querySelector('.text-content')?.textContent || '') : '';

    if (!emailBody) {
        alert('Cannot send email: No outreach text found.');
        return;
    }

    const sendBtn = outreachCell.querySelector('.btn-send-email');
    if (!sendBtn) return;

    const originalText = sendBtn.textContent;

    try {
        sendBtn.disabled = true;
        sendBtn.textContent = 'Sending...';

        const response = await fetch(CONFIG.emailWebhookUrl, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                title: jobTitle,
                email: emailAddress,
                body: emailBody,
                timestamp: new Date().toISOString()
            })
        });

        if (!response.ok) throw new Error('Webhook failed');

        sendBtn.textContent = 'Sent';
        sendBtn.classList.add('sent');

        setTimeout(() => {
            sendBtn.textContent = originalText;
            sendBtn.disabled = false;
            sendBtn.classList.remove('sent');
        }, 3000);

    } catch (error) {
        console.error('Send email error:', error);
        alert('Failed to send email. Check console for details.');
        sendBtn.textContent = 'Failed';
        sendBtn.disabled = false;
    }
}
