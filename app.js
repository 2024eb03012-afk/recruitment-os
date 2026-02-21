// ===== Configuration =====
const CONFIG = {
    webhookUrl: 'https://n8n.smallgrp.com/webhook-test/3be05281-781a-407e-8691-311c1876a1f2',
    // Google Sheet URLs
    allJobsSheetUrl: 'https://docs.google.com/spreadsheets/d/1DAVbKlf0bI3Dkmt8YQlAawoG7Oeez4px6vHhWy7Ts8w/gviz/tq?tqx=out:csv&gid=1341392940',
    currentJobsSheetUrl: 'https://docs.google.com/spreadsheets/d/1DAVbKlf0bI3Dkmt8YQlAawoG7Oeez4px6vHhWy7Ts8w/gviz/tq?tqx=out:csv&gid=635880039',
    // Auto-refresh interval in milliseconds (15 seconds)
    refreshInterval: 15000,
    // Auto-refresh duration after search (5 minutes)
    autoRefreshDuration: 5 * 60 * 1000,
    emailWebhookUrl: 'https://n8n.smallgrp.com/webhook-test/8f452cec-9856-40ce-a791-e81b57bef0e3'
};

// Column mapping from Google Sheet headers to internal keys
const COLUMN_MAPPING = {
    'id': 'id',
    'platform': 'platform',
    'company_name': 'companyName',
    'title': 'title',
    'job_type': 'jobType',
    'location': 'location',
    'work_model': 'workModel',
    'jd': 'jd',
    'company_job_url': 'companyJobUrl',
    'apply_url': 'applyUrl',
    'salary': 'salary',
    'company_description': 'companyDescription',
    'website': 'website',
    'decision_maker_email': 'decisionMakerEmail',
    'match_score_analysis': 'matchScore',
    'outreach_email_text': 'outreachEmailText'
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
    selectedCount: document.getElementById('selectedCount'),
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
    cancelBtn: document.getElementById('cancelBtn'),
    tableSearch: document.getElementById('tableSearch'),
    selectAll: document.getElementById('selectAll'),
    quickActions: document.getElementById('quickActions'),
    quickActionsCount: document.getElementById('quickActionsCount'),
    bulkEmailBtn: document.getElementById('bulkEmailBtn'),
    clearSelectionBtn: document.getElementById('clearSelectionBtn'),
    sheetToggle: document.getElementById('sheetToggle'),
    autoRefreshIndicator: document.getElementById('autoRefreshIndicator')
};

// ===== Global State =====
let sheetData = [];
let filteredData = [];
let autoRefreshTimer = null;
let autoRefreshEndTime = null; // When auto-refresh should stop
let localEdits = {}; // Persistent local overrides during session
let selectedRows = new Set(); // Track selected row indices
let currentSheet = 'current'; // 'current' or 'all'
let currentModalContext = {
    dataIndex: -1,
    columnKey: '',
    isEditing: false
};
let chartInstances = {}; // Track Chart.js instances
let isRefreshing = false; // Prevent concurrent refreshes

// ===== Initialization =====
document.addEventListener('DOMContentLoaded', () => {
    initForm();
    loadSheetData();
    setupEventListeners();
    // Don't start auto-refresh on page load - only after search
});

// ===== Event Listeners =====
function setupEventListeners() {
    elements.refreshBtn.addEventListener('click', manualRefresh);
    elements.platformLinkedInPost.addEventListener('change', toggleLinkedInKeywords);

    // Sheet toggle
    if (elements.sheetToggle) {
        elements.sheetToggle.addEventListener('click', (e) => {
            const btn = e.target.closest('.toggle-btn');
            if (btn && !btn.classList.contains('active')) {
                elements.sheetToggle.querySelectorAll('.toggle-btn').forEach(b => b.classList.remove('active'));
                btn.classList.add('active');
                currentSheet = btn.dataset.sheet;
                loadSheetData();
            }
        });
    }

    // Table search
    elements.tableSearch.addEventListener('input', handleTableSearch);

    // Select all checkbox
    elements.selectAll.addEventListener('change', handleSelectAll);

    // Quick actions
    elements.bulkEmailBtn.addEventListener('click', handleBulkEmail);
    elements.clearSelectionBtn.addEventListener('click', clearSelection);

    // Global click handler for view buttons, expandable text, and row checkboxes
    elements.tableBody.addEventListener('click', (e) => {
        // Handle row checkbox
        const rowCheckbox = e.target.closest('.row-checkbox input');
        if (rowCheckbox) {
            const row = rowCheckbox.closest('tr');
            const rowIndex = parseInt(row.dataset.index);
            handleRowSelect(rowIndex, rowCheckbox.checked);
            return;
        }

        // Handle View button clicks
        const viewBtn = e.target.closest('.btn-view');
        if (viewBtn) {
            e.stopPropagation();
            const row = viewBtn.closest('tr');
            const originalIndex = parseInt(row.dataset.index);
            const columnKey = viewBtn.dataset.column;
            const item = sheetData[originalIndex];
            if (item) {
                const columnLabels = {
                    'jd': 'Job Description',
                    'companyDescription': 'Company Description',
                    'outreachEmailText': 'Outreach Email'
                };
                openModal(item.companyName || 'Details', item[columnKey] || '', {
                    dataIndex: originalIndex,
                    columnKey,
                    columnLabel: columnLabels[columnKey] || columnKey
                });
            }
            return;
        }

        // Handle Send Email button click
        const sendBtn = e.target.closest('.btn-send-email');
        if (sendBtn) {
            e.stopPropagation();
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

    // Date range buttons
    document.querySelectorAll('.date-btn').forEach(btn => {
        btn.addEventListener('click', (e) => {
            document.querySelectorAll('.date-btn').forEach(b => b.classList.remove('active'));
            e.target.classList.add('active');
            // In a real app, this would filter the data
        });
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

// ===== Table Search =====
function handleTableSearch(e) {
    const query = e.target.value.toLowerCase().trim();

    if (!query) {
        filteredData = [...sheetData].reverse();
    } else {
        filteredData = [...sheetData].reverse().filter(row => {
            return Object.values(row).some(val =>
                val && val.toString().toLowerCase().includes(query)
            );
        });
    }

    renderTable(filteredData);
}

// ===== Row Selection =====
function handleRowSelect(index, isSelected) {
    if (isSelected) {
        selectedRows.add(index);
    } else {
        selectedRows.delete(index);
    }
    updateSelectionUI();
}

function handleSelectAll(e) {
    const isChecked = e.target.checked;
    const rows = elements.tableBody.querySelectorAll('tr');

    selectedRows.clear();

    rows.forEach(row => {
        const checkbox = row.querySelector('.row-checkbox input');
        if (checkbox) {
            checkbox.checked = isChecked;
            if (isChecked) {
                selectedRows.add(parseInt(row.dataset.index));
            }
        }
    });

    updateSelectionUI();
}

function clearSelection() {
    selectedRows.clear();
    elements.selectAll.checked = false;
    elements.tableBody.querySelectorAll('.row-checkbox input').forEach(cb => cb.checked = false);
    updateSelectionUI();
}

function updateSelectionUI() {
    const count = selectedRows.size;

    if (count > 0) {
        elements.selectedCount.textContent = `${count} selected`;
        elements.selectedCount.style.display = 'inline';
        elements.quickActions.style.display = 'flex';
        elements.quickActionsCount.textContent = count;
    } else {
        elements.selectedCount.style.display = 'none';
        elements.quickActions.style.display = 'none';
    }
}

// ===== Bulk Actions =====
async function handleBulkEmail() {
    if (selectedRows.size === 0) return;

    const btn = elements.bulkEmailBtn;
    const originalText = btn.innerHTML;
    btn.disabled = true;
    btn.innerHTML = 'Sending...';

    let sent = 0;
    let failed = 0;

    for (const index of selectedRows) {
        const row = sheetData[index];
        if (row && row.decisionMakerEmail && row.outreachEmailText) {
            try {
                await fetch(CONFIG.emailWebhookUrl, {
                    method: 'POST',
                    mode: 'no-cors',
                    headers: { 'Content-Type': 'text/plain' },
                    body: JSON.stringify({
                        title: row.title || '',
                        email: row.decisionMakerEmail,
                        body: row.outreachEmailText,
                        timestamp: new Date().toISOString()
                    })
                });
                sent++;
            } catch (error) {
                failed++;
            }
        }
    }

    btn.innerHTML = `Sent ${sent} emails`;
    setTimeout(() => {
        btn.innerHTML = originalText;
        btn.disabled = false;
        clearSelection();
    }, 2000);
}

// ===== Auto Refresh =====
function startAutoRefresh() {
    // Clear any existing timer
    stopAutoRefresh();

    // Set end time for auto-refresh (5 minutes from now)
    autoRefreshEndTime = Date.now() + CONFIG.autoRefreshDuration;
    updateAutoRefreshIndicator();

    // Set up auto-refresh every 15 seconds
    autoRefreshTimer = setInterval(() => {
        if (Date.now() >= autoRefreshEndTime) {
            stopAutoRefresh();
            return;
        }
        console.log('Auto-refreshing data...');
        loadSheetDataSmart();
        updateAutoRefreshIndicator();
    }, CONFIG.refreshInterval);
}

function stopAutoRefresh() {
    if (autoRefreshTimer) {
        clearInterval(autoRefreshTimer);
        autoRefreshTimer = null;
    }
    autoRefreshEndTime = null;
    updateAutoRefreshIndicator();
}

function updateAutoRefreshIndicator() {
    if (!elements.autoRefreshIndicator) return;

    if (autoRefreshEndTime && Date.now() < autoRefreshEndTime) {
        const remainingMs = autoRefreshEndTime - Date.now();
        const remainingMin = Math.ceil(remainingMs / 60000);
        elements.autoRefreshIndicator.innerHTML = `
            <span class="auto-refresh-dot"></span>
            Auto-refresh: ${remainingMin}m left
        `;
        elements.autoRefreshIndicator.style.display = 'flex';
    } else {
        elements.autoRefreshIndicator.style.display = 'none';
    }
}

function manualRefresh() {
    elements.refreshBtn.classList.add('spinning');
    loadSheetData();

    setTimeout(() => {
        elements.refreshBtn.classList.remove('spinning');
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
        console.log('Sending payload:', formData);

        const response = await fetch(CONFIG.webhookUrl, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(formData)
        });

        if (!response.ok) {
            if (response.status === 500) {
                // Specific hint for n8n 500 errors
                throw new Error('n8n Workflow Error (500). Check n8n Executions log.');
            }
            throw new Error(`Server error: ${response.status} ${response.statusText}`);
        }

        // Keep form data visible (don't reset)
        showMessage('success', 'Search request submitted. Results will appear automatically for the next 5 minutes.');

        // Start timed auto-refresh (15s intervals for 5 minutes)
        startAutoRefresh();

        // Initial refresh after 3 seconds
        setTimeout(() => {
            loadSheetData();
        }, 3000);

    } catch (error) {
        console.error('Form submission error:', error);

        let displayMsg = 'Failed to submit request. Please check your connection.';
        if (error.message.includes('n8n') || error.message.includes('Server error')) {
            displayMsg = error.message;
        }

        showMessage('error', displayMsg);
    } finally {
        setLoadingState(false);
    }
}

function collectFormData() {
    const formData = new FormData(elements.form);
    const selectedPlatforms = formData.getAll('platform');

    return {
        jobTitle: formData.get('jobTitle'),
        jobTypes: formData.getAll('jobType'),
        location: formData.get('location'),
        numJobs: parseInt(formData.get('numJobs')),
        platforms: {
            linkedin: selectedPlatforms.includes('linkedin'),
            indeed: selectedPlatforms.includes('indeed'),
            glassdoor: selectedPlatforms.includes('glassdoor'),
            linkedinPost: selectedPlatforms.includes('linkedin_post')
        },
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

    if (data.platforms.linkedinPost && !data.linkedinKeywords?.trim()) {
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

        // Render table with all data (newest first - reverse order)
        renderTable([...sheetData].reverse());

    } catch (error) {
        console.error('Sheet data load error:', error);
        elements.tableError.style.display = 'block';
        elements.analyticsError.style.display = 'block';
        elements.tableBody.innerHTML = '';
        elements.analyticsGrid.innerHTML = '';
    }
}

// Smart refresh - only adds new rows without jarring reload
async function loadSheetDataSmart() {
    if (isRefreshing) return;
    isRefreshing = true;

    try {
        const data = await fetchSheetData();

        // Create a Set of existing IDs for quick lookup
        const existingIds = new Set(sheetData.map(item => item.id || `${item.companyName}|${item.title}`));

        // Find new items
        const newItems = data.filter(item => {
            const itemId = item.id || `${item.companyName}|${item.title}`;
            return !existingIds.has(itemId);
        });

        if (newItems.length > 0) {
            console.log(`Found ${newItems.length} new items`);

            // Merge with local edits
            const processedNewItems = newItems.map(item => {
                const editKey = `${item.companyName}|${item.title}`;
                if (localEdits[editKey]) {
                    return { ...item, ...localEdits[editKey] };
                }
                return item;
            });

            // Add new items to sheetData
            sheetData = [...sheetData, ...processedNewItems];

            // Update analytics and charts
            renderAnalytics(sheetData);
            renderCharts(sheetData);

            // Add new rows to the top of the table (since we display reversed)
            prependTableRows(processedNewItems.reverse());

            // Update count
            elements.tableCount.textContent = `${sheetData.length} records`;
        }
    } catch (error) {
        console.error('Smart refresh error:', error);
    } finally {
        isRefreshing = false;
    }
}

// Prepend new rows to the table without full re-render
function prependTableRows(newItems) {
    if (!newItems || newItems.length === 0) return;

    const newRowsHtml = newItems.map(row => {
        const originalIndex = sheetData.findIndex(item =>
            item.companyName === row.companyName && item.title === row.title
        );
        return generateTableRow(row, originalIndex);
    }).join('');

    // Create a temporary container to parse HTML
    const temp = document.createElement('tbody');
    temp.innerHTML = newRowsHtml;

    // Prepend each new row with animation
    const existingFirstRow = elements.tableBody.firstChild;
    Array.from(temp.children).reverse().forEach(newRow => {
        newRow.classList.add('new-row');
        elements.tableBody.insertBefore(newRow, existingFirstRow);
    });

    // Remove animation class after animation completes
    setTimeout(() => {
        elements.tableBody.querySelectorAll('.new-row').forEach(row => {
            row.classList.remove('new-row');
        });
    }, 1000);
}

async function fetchSheetData() {
    const sheetUrl = currentSheet === 'all'
        ? CONFIG.allJobsSheetUrl
        : CONFIG.currentJobsSheetUrl;

    const response = await fetch(sheetUrl);

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

            // Clean City column to remove raw JSON if present
            if (mappedKey === 'city' && value.trim().startsWith('{')) {
                try {
                    const cityObj = JSON.parse(value);
                    const parts = [];
                    // Extract structured location data
                    if (cityObj.city && cityObj.city !== 'null') parts.push(cityObj.city);
                    if (cityObj.state && cityObj.state !== 'null') parts.push(cityObj.state);
                    if (cityObj.country && cityObj.country !== 'null') parts.push(cityObj.country);

                    if (parts.length > 0) {
                        value = parts.join(', ');
                    } else if (cityObj.formattedAddressShort) {
                        value = cityObj.formattedAddressShort;
                    }
                } catch (e) {
                    // diverse formats sometimes occur, fallback to original if parse fails
                }
            }

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

    // Generate random trends for demo (in production, compare with previous period)
    const trends = {
        totalJobs: { value: 12, positive: true },
        uniqueCompanies: { value: 8, positive: true },
        topJobType: null,
        uniqueCities: { value: 3, positive: true },
        avgMatchScore: { value: 5, positive: true },
        outreachReady: { value: 15, positive: true }
    };

    elements.analyticsGrid.innerHTML = `
        <div class="stat-card glass-card clickable" data-filter="all">
            <div class="stat-icon total"></div>
            <div class="stat-content">
                <span class="stat-label">Total Jobs Found</span>
                <span class="stat-value">${formatNumber(metrics.totalJobs)}</span>
                ${renderTrend(trends.totalJobs)}
            </div>
        </div>
        <div class="stat-card glass-card clickable" data-filter="company">
            <div class="stat-icon companies"></div>
            <div class="stat-content">
                <span class="stat-label">Unique Companies</span>
                <span class="stat-value">${formatNumber(metrics.uniqueCompanies)}</span>
                ${renderTrend(trends.uniqueCompanies)}
            </div>
        </div>
        <div class="stat-card glass-card clickable" data-filter="jobType">
            <div class="stat-icon types"></div>
            <div class="stat-content">
                <span class="stat-label">Top Job Type</span>
                <span class="stat-value">${metrics.topJobType}</span>
            </div>
        </div>
        <div class="stat-card glass-card clickable" data-filter="location">
            <div class="stat-icon cities"></div>
            <div class="stat-content">
                <span class="stat-label">Locations</span>
                <span class="stat-value">${formatNumber(metrics.uniqueLocations)}</span>
                ${renderTrend(trends.uniqueCities)}
            </div>
        </div>
        <div class="stat-card glass-card clickable" data-filter="score">
            <div class="stat-icon score"></div>
            <div class="stat-content">
                <span class="stat-label">Avg Match Score</span>
                <span class="stat-value">${metrics.avgMatchScore}%</span>
                ${renderTrend(trends.avgMatchScore)}
            </div>
        </div>
        <div class="stat-card glass-card highlight clickable" data-filter="outreach">
            <div class="stat-icon leads"></div>
            <div class="stat-content">
                <span class="stat-label">Ready for Outreach</span>
                <span class="stat-value">${formatNumber(metrics.outreachReady)}</span>
                ${renderTrend(trends.outreachReady)}
            </div>
        </div>
    `;

    // Add click handlers for stat cards
    document.querySelectorAll('.stat-card.clickable').forEach(card => {
        card.addEventListener('click', () => {
            const filter = card.dataset.filter;
            if (filter === 'outreach') {
                // Filter to show only rows with emails
                elements.tableSearch.value = '';
                filteredData = [...sheetData].reverse().filter(row => row.decisionMakerEmail && row.decisionMakerEmail.trim());
                renderTable(filteredData);
            }
        });
    });
}

function renderTrend(trend) {
    if (!trend) return '';
    const icon = trend.positive
        ? '<svg width="12" height="12" viewBox="0 0 12 12" fill="none"><path d="M6 2v8M3 5l3-3 3 3" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"/></svg>'
        : '<svg width="12" height="12" viewBox="0 0 12 12" fill="none"><path d="M6 10V2M3 7l3 3 3-3" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"/></svg>';
    return `<span class="stat-trend ${trend.positive ? 'positive' : 'negative'}">${icon} ${trend.value}%</span>`;
}

function calculateMetrics(data) {
    const uniqueCompanies = new Set(data.map(item => item.companyName).filter(Boolean));
    const uniqueLocations = new Set(data.map(item => item.location).filter(Boolean));

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
        uniqueLocations: uniqueLocations.size,
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
        '#c96442'
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
        '#c9a227'
    );

    // 3. Jobs by Location (Top 10)
    const locationCounts = data.reduce((acc, item) => {
        let location = (item.location || 'Remote/Unknown').trim();
        // Capitalize each word for uniform grouping
        location = location.split(' ').map(w => w.charAt(0).toUpperCase() + w.slice(1).toLowerCase()).join(' ');
        acc[location] = (acc[location] || 0) + 1;
        return acc;
    }, {});
    const topLocations = Object.entries(locationCounts)
        .sort((a, b) => b[1] - a[1])
        .slice(0, 10);

    createBarChart('cityChart',
        topLocations.map(c => c[0]),
        topLocations.map(c => c[1]),
        '#5a8f7b'
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
        '#9b8bb4'
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
        '#c45c4a'
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
        '#6b9080'
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
                    backgroundColor: '#2d2a26',
                    padding: 10,
                    titleFont: { size: 13, weight: '600' },
                    bodyFont: { size: 12 },
                    displayColors: false,
                    cornerRadius: 6
                }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    grid: { color: '#f3f1ee' },
                    ticks: { stepSize: 1, color: '#8a837a', font: { size: 11 } }
                },
                x: {
                    grid: { display: false },
                    ticks: { color: '#8a837a', font: { size: 11 } }
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
            <td class="checkbox-col"><span class="skeleton-loader" style="width: 20px; height: 20px;"></span></td>
            ${Array(13).fill('<td><span class="skeleton-loader"></span></td>').join('')}
        </tr>
    `).join('');

    elements.tableBody.innerHTML = skeletonRows;
}

function renderTable(data) {
    elements.tableError.style.display = 'none';
    selectedRows.clear();
    updateSelectionUI();
    elements.selectAll.checked = false;

    if (!data || data.length === 0) {
        elements.tableBody.innerHTML = '';
        elements.emptyState.style.display = 'block';
        elements.tableCount.textContent = '0 records';
        return;
    }

    elements.emptyState.style.display = 'none';
    elements.tableCount.textContent = `${data.length} records`;

    elements.tableBody.innerHTML = data.map((row, index) => {
        // Find the original index in sheetData
        const originalIndex = sheetData.findIndex(item =>
            item.companyName === row.companyName && item.title === row.title
        );
        return generateTableRow(row, originalIndex);
    }).join('');
}

function generateTableRow(row, originalIndex) {
    return `
        <tr data-index="${originalIndex}">
            <td class="checkbox-col">
                <label class="table-checkbox row-checkbox">
                    <input type="checkbox">
                    <span class="checkmark"></span>
                </label>
            </td>
            <td><strong>${escapeHtml(row.companyName || '')}</strong></td>
            <td><strong>${escapeHtml(row.title || '')}</strong></td>
            <td>${renderPlatformBadge(row.platform)}</td>
            <td><span class="job-type-badge">${escapeHtml(row.jobType || '')}</span></td>
            <td>${escapeHtml(row.location || '')}</td>
            <td>${escapeHtml(row.workModel || '')}</td>
            <td>${renderViewButton(row.jd, 'jd', 'JD')}</td>
            <td>${renderUrlButton(row.companyJobUrl, 'Job URL')}</td>
            <td>${escapeHtml(row.salary || '')}</td>
            <td>${renderViewButton(row.companyDescription, 'companyDescription', 'Info')}</td>
            <td>${renderMatchScore(row.matchScore)}</td>
            <td>${escapeHtml(row.decisionMakerEmail || '')}</td>
            <td class="outreach-cell">
                ${renderViewButton(row.outreachEmailText, 'outreachEmailText', 'Email')}
                ${row.outreachEmailText ? `
                <button class="btn-send-email" title="Send Email">
                    <svg width="12" height="12" viewBox="0 0 12 12" fill="none" xmlns="http://www.w3.org/2000/svg">
                        <path d="M1 3L6 6.5L11 3" stroke="currentColor" stroke-width="1.5" stroke-linecap="round"/>
                        <rect x="1" y="2" width="10" height="8" rx="1" stroke="currentColor" stroke-width="1.5"/>
                    </svg>
                    Send
                </button>` : ''}
            </td>
        </tr>
    `;
}

function renderPlatformBadge(platform) {
    if (!platform) return '';
    const platformLower = platform.toLowerCase();
    let colorClass = 'platform-default';
    if (platformLower.includes('linkedin')) colorClass = 'platform-linkedin';
    else if (platformLower.includes('indeed')) colorClass = 'platform-indeed';
    else if (platformLower.includes('glassdoor')) colorClass = 'platform-glassdoor';
    return `<span class="platform-badge ${colorClass}">${escapeHtml(platform)}</span>`;
}

function renderViewButton(text, columnKey, label) {
    if (!text || !text.trim()) return '<span class="text-muted">-</span>';
    return `<button class="btn-view" data-column="${columnKey}">${label}</button>`;
}

function renderUrlButton(url, label) {
    if (!url || !url.trim()) return '<span class="text-muted">-</span>';
    return `<a href="${escapeHtml(url)}" target="_blank" rel="noopener noreferrer" class="btn-url">${label}</a>`;
}

function renderMatchScore(score) {
    if (!score) return '';
    const numScore = parseFloat(score);
    if (isNaN(numScore)) return escapeHtml(score);

    let colorClass = 'score-low';
    if (numScore >= 85) colorClass = 'score-high';
    else if (numScore >= 70) colorClass = 'score-medium';

    return `<span class="match-score-badge ${colorClass}">${numScore}%</span>`;
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
    const originalIndex = parseInt(row.dataset.index);
    const item = sheetData[originalIndex];

    if (!item) {
        alert('Cannot send email: Row data not found.');
        return;
    }

    const jobTitle = item.title || '';
    const emailAddress = item.decisionMakerEmail || '';
    const emailBody = item.outreachEmailText || '';

    if (!emailBody) {
        alert('Cannot send email: No outreach text found.');
        return;
    }

    if (!emailAddress) {
        alert('Cannot send email: No email address found.');
        return;
    }

    const sendBtn = row.querySelector('.btn-send-email');
    if (!sendBtn) return;

    const originalHTML = sendBtn.innerHTML;

    try {
        sendBtn.disabled = true;
        sendBtn.innerHTML = 'Sending...';

        await fetch(CONFIG.emailWebhookUrl, {
            method: 'POST',
            mode: 'no-cors',
            headers: { 'Content-Type': 'text/plain' },
            body: JSON.stringify({
                title: jobTitle,
                email: emailAddress,
                body: emailBody,
                timestamp: new Date().toISOString()
            })
        });

        sendBtn.innerHTML = 'Sent!';
        sendBtn.classList.add('sent');

        setTimeout(() => {
            sendBtn.innerHTML = originalHTML;
            sendBtn.disabled = false;
            sendBtn.classList.remove('sent');
        }, 3000);

    } catch (error) {
        console.error('Send email error:', error);
        alert('Failed to send email. Check console for details.');
        sendBtn.innerHTML = 'Failed';
        sendBtn.disabled = false;
    }
}
