/**
 * Hiring Signal Engine — Raw Data Sync from Google Sheet
 * Fetches CSV data and displays it exactly as-is in a 16-column table.
 */

const HiringSignalEngine = (() => {
    // ── Config ──────────────────────────────────────────────
    const SHEET_URL = 'https://docs.google.com/spreadsheets/d/1NnHxHQjYxvQv-I1OiSGJWzPcXfqrmliG47hjHm1KiVs/gviz/tq?tqx=out:csv';
    const REFRESH_INTERVAL = 5 * 60 * 1000; // 5 minutes
    const WEBHOOK_URL = 'https://n8n.smallgrp.com/webhook-test/17c802e4-a19b-4528-9722-ed4eb486e433';

    // Maps CSV header names → internal keys (in column order)
    const COLUMN_MAP = [
        { csv: 'company_name', key: 'companyName' },
        { csv: 'company_website', key: 'companyWebsite' },
        { csv: 'funding_amount', key: 'fundingAmount' },
        { csv: 'funding_date', key: 'fundingDate' },
        { csv: 'company_size', key: 'companySize' },
        { csv: 'industry', key: 'industry' },
        { csv: 'company_description', key: 'companyDescription' },
        { csv: 'target_decision_maker', key: 'targetDecisionMaker' },
        { csv: 'contact_email', key: 'contactEmail' },
        { csv: 'url_company_linkedin', key: 'companyLinkedin' },
        { csv: 'comapany_recent_news', key: 'companyRecentNews' },
        { csv: 'outreach_message', key: 'outreachMessage' },
        { csv: 'predicted_hiring_score', key: 'hiringScore' },
        { csv: 'hiring_impact_level', key: 'impactLevel' },
        { csv: 'hiring_impact_type', key: 'impactType' },
        { csv: 'senior_employee_linkedin_profile', key: 'seniorLinkedin' },
    ];

    // ── DOM refs ────────────────────────────────────────────
    const el = {
        tableBody: () => document.getElementById('hsTableBody'),
        recordCount: () => document.getElementById('hsRecordCount'),
        emptyState: () => document.getElementById('hsEmptyState'),
        errorState: () => document.getElementById('hsError'),
        refreshBtn: () => document.getElementById('hsRefreshBtn'),
        modal: () => document.getElementById('hsTextModal'),
        modalTitle: () => document.getElementById('hsTextModalTitle'),
        modalBody: () => document.getElementById('hsTextModalBody'),
        modalFooter: () => document.getElementById('hsTextModalFooter'),
        modalClose: () => document.getElementById('hsTextModalClose'),
    };

    let refreshTimer = null;
    let initialized = false;

    // ── CSV Parser ──────────────────────────────────────────
    function parseCSV(text) {
        const rows = [];
        let current = '';
        let inQuotes = false;
        const lines = [];

        for (let i = 0; i < text.length; i++) {
            const ch = text[i];
            if (ch === '"') {
                inQuotes = !inQuotes;
                current += ch;
            } else if (ch === '\n' && !inQuotes) {
                lines.push(current);
                current = '';
            } else {
                current += ch;
            }
        }
        if (current.trim()) lines.push(current);

        if (lines.length === 0) return [];

        const headers = parseCSVLine(lines[0]);

        for (let i = 1; i < lines.length; i++) {
            if (!lines[i].trim()) continue;
            const values = parseCSVLine(lines[i]);
            const row = {};
            headers.forEach((h, idx) => {
                const key = h.trim().toLowerCase().replace(/[^a-z0-9]+/g, '_').replace(/^_|_$/g, '');
                row[key] = (values[idx] || '').trim();
            });
            rows.push(row);
        }

        return rows;
    }

    function parseCSVLine(line) {
        const result = [];
        let current = '';
        let inQuotes = false;

        for (let i = 0; i < line.length; i++) {
            const ch = line[i];
            if (ch === '"') {
                if (inQuotes && line[i + 1] === '"') {
                    current += '"';
                    i++;
                } else {
                    inQuotes = !inQuotes;
                }
            } else if (ch === ',' && !inQuotes) {
                result.push(current);
                current = '';
            } else {
                current += ch;
            }
        }
        result.push(current);
        return result;
    }

    // ── Data Mapping ────────────────────────────────────────
    function mapRow(raw) {
        const mapped = {};
        for (const col of COLUMN_MAP) {
            mapped[col.key] = raw[col.csv] || '';
        }
        return mapped;
    }

    // ── Employee Parsing ────────────────────────────────────
    function parseEmployees(text) {
        if (!text) return [];
        const employees = [];
        const regex = /\d+\.\s*First name\s*:\s*(.+?)\s*Linkedin url\s*:\s*(https?:\/\/[^\s]+)\s*Current Title\s*:\s*(.+?)(?=\s*\d+\.\s*First name|$)/gis;

        let match;
        while ((match = regex.exec(text)) !== null) {
            employees.push({
                name: match[1].trim(),
                linkedin: match[2].trim(),
                title: match[3].trim()
            });
        }
        return employees;
    }

    function renderEmployeeCard(emp) {
        return `
            <div class="hs-employee-card">
                <div class="hs-emp-name">${escapeHtml(emp.name)}</div>
                <div class="hs-emp-title-badge">${escapeHtml(emp.title)}</div>
                <a href="${escapeHtml(emp.linkedin)}" target="_blank" rel="noopener" class="hs-emp-linkedin-btn">
                    View LinkedIn
                </a>
            </div>
        `;
    }

    // ── Helpers ─────────────────────────────────────────────
    function escapeHtml(str) {
        if (!str) return '-';
        return String(str)
            .replace(/&/g, "&amp;")
            .replace(/</g, "&lt;")
            .replace(/>/g, "&gt;")
            .replace(/"/g, "&quot;")
            .replace(/'/g, "&#039;");
    }

    function isUrl(str) {
        return str && (str.startsWith('http://') || str.startsWith('https://'));
    }

    function renderViewButton(content, title, label, email = '') {
        if (!content || content === '-') return '<span class="text-muted">-</span>';

        const emailAttr = email ? `data-contact-email="${escapeHtml(email)}"` : '';

        return `<button class="btn-view" 
            data-full-text="${encodeURIComponent(content)}" 
            data-modal-title="${title}"
            ${emailAttr}>${label}</button>`;
    }

    function renderUrlButton(url, label) {
        if (!url || !isUrl(url)) return '<span class="text-muted">-</span>';
        return `<a href="${escapeHtml(url)}" target="_blank" rel="noopener" class="btn-url">${label}</a>`;
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

    function renderBadge(text) {
        if (!text || text === '-') return '';
        return `<span class="job-type-badge">${escapeHtml(text)}</span>`;
    }

    // ── Modal Logic ─────────────────────────────────────────
    function openTextModal(title, content, email, isHtml = false) {
        const modal = el.modal();
        if (!modal) return;
        el.modalTitle().textContent = title;

        const isEditable = title === 'Outreach Message';
        const bodyContent = isEditable
            ? `<textarea class="hs-modal-textarea" id="hsOutreachText">${escapeHtml(content)}</textarea>`
            : `<div class="hs-modal-text">${isHtml ? content : escapeHtml(content)}</div>`;

        el.modalBody().innerHTML = bodyContent;

        const footer = el.modalFooter();
        if (footer) {
            if (isEditable && email) {
                footer.style.display = 'flex';
                footer.innerHTML = `<button class="hs-btn-send" data-email="${escapeHtml(email)}">
                    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                        <line x1="22" y1="2" x2="11" y2="13"></line>
                        <polygon points="22 2 15 22 11 13 2 9 22 2"></polygon>
                    </svg>
                    Send Email
                </button>`;
            } else {
                footer.style.display = 'none';
                footer.innerHTML = '';
            }
        }

        modal.classList.add('active');
    }

    function closeTextModal() {
        const modal = el.modal();
        if (modal) modal.classList.remove('active');
    }

    function handleSendEmail(btn) {
        const email = btn.getAttribute('data-email');
        const textarea = document.getElementById('hsOutreachText');

        let content = '';
        // If textarea exists (modal open), use its value. Otherwise use data-full-text from button context if stored?
        // Actually, handleSendEmail is called from modal footer usually.
        // BUT we also have a "Send" button in the table row now.

        if (textarea) {
            content = textarea.value;
        } else {
            // Direct send from table row
            // We need the content. It might be stored on the button or we need to look it up.
            // For simplicity, let's say the row button opens the modal by default OR sends directly if we have content.
            // But we don't have the content easily accessible here without passing it.
            // Let's make the row "Send" button open the mailto directly with the raw content from data attribute.
            const rawContent = btn.getAttribute('data-full-text');
            if (rawContent) content = decodeURIComponent(rawContent);
        }

        if (!email) return;

        let subject = "Outreach";
        let body = content;

        // Try to parse JSON for subject/body
        try {
            if (content.trim().startsWith('{')) {
                const json = JSON.parse(content);
                if (json.subject) subject = json.subject;
                if (json.body) body = json.body;
                else if (json.message) body = json.message;
            }
        } catch (e) {
            // Not JSON, fall back to plain text body
        }

        // Fire webhook with email, message, and company name
        const companyName = btn.getAttribute('data-company-name') || '';
        sendToWebhook({ email, message: body, subject, companyName });

        // Visual feedback
        btn.textContent = '✓ Sent';
        btn.disabled = true;
        setTimeout(() => {
            btn.textContent = 'Send';
            btn.disabled = false;
        }, 3000);
    }

    // ── Webhook ─────────────────────────────────────────────
    function sendToWebhook(payload) {
        fetch(WEBHOOK_URL, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(payload)
        })
            .then(res => console.log('Webhook sent:', res.status))
            .catch(err => console.warn('Webhook error:', err));
    }

    // ── Render ──────────────────────────────────────────────
    function renderCell(value, key, row) { // Added row parameter
        if (!value || value === '-') return '<td class="hs-cell">-</td>';

        // Employee Column: Compact View
        if (key === 'seniorLinkedin') {
            const employees = parseEmployees(value);
            if (employees.length > 0) {
                // Structured view with cards
                return `<td class="hs-cell">
                    <button class="btn-view" 
                        data-type="employees"
                        data-full-text="${encodeURIComponent(value)}" 
                        data-modal-title="Senior Team (${employees.length})">
                        View Team (${employees.length})
                    </button>
                </td>`;
            } else {
                // Fallback: show raw text in modal
                return `<td class="hs-cell">
                    <button class="btn-view" 
                        data-type="raw"
                        data-full-text="${encodeURIComponent(value)}" 
                        data-modal-title="Senior LinkedIn Profiles">
                        View Profiles
                    </button>
                </td>`;
            }
        }

        // URL columns
        if (key === 'companyWebsite') return `<td class="hs-cell">${renderUrlButton(value, 'Website')}</td>`;
        if (key === 'companyLinkedin') return `<td class="hs-cell">${renderUrlButton(value, 'LinkedIn')}</td>`;
        if (key === 'contactEmail') return `<td class="hs-cell"><a href="mailto:${escapeHtml(value)}" class="hs-link">${escapeHtml(value)}</a></td>`;

        // Interactive Columns
        if (key === 'companyDescription') return `<td class="hs-cell">${renderViewButton(value, 'Company Description', 'Info')}</td>`;
        if (key === 'companyRecentNews') return `<td class="hs-cell">${renderViewButton(value, 'Recent News', 'News')}</td>`;

        // Data Columns
        if (key === 'hiringScore') return `<td class="hs-cell">${renderMatchScore(value)}</td>`;
        if (key === 'impactLevel') {
            if (!value || value === '-') return '<td class="hs-cell">-</td>';

            // Always make it a clickable button to show full details
            // Extract short label from start (e.g. "Very High" from "Very High: explanation...")
            let label = value;
            const colonIndex = value.indexOf(':');
            if (colonIndex > 0 && colonIndex < 30) {
                label = value.substring(0, colonIndex).trim();
            } else if (value.length > 20) {
                label = value.substring(0, 20).trim() + '…';
            }
            return `<td class="hs-cell">${renderViewButton(value, 'Impact Details', label)}</td>`;
        }

        // Outreach Message (Special Handling)
        if (key === 'outreachMessage') {
            const email = row.contactEmail;
            const companyName = row.companyName || '';
            const viewBtn = renderViewButton(value, 'Outreach Message', 'Email', email);

            let sendBtn = '';
            if (email) {
                sendBtn = `<button class="btn-send-email" 
                    data-email="${escapeHtml(email)}" 
                    data-full-text="${encodeURIComponent(value)}"
                    data-company-name="${escapeHtml(companyName)}"
                    title="Send Email">
                    <svg width="12" height="12" viewBox="0 0 12 12" fill="none" xmlns="http://www.w3.org/2000/svg">
                        <path d="M1 3L6 6.5L11 3" stroke="currentColor" stroke-width="1.5" stroke-linecap="round"/>
                        <rect x="1" y="2" width="10" height="8" rx="1" stroke="currentColor" stroke-width="1.5"/>
                    </svg>
                    Send
                </button>`;
            }

            return `<td class="hs-cell">
                <div class="outreach-cell">
                    ${viewBtn}
                    ${sendBtn}
                </div>
             </td>`;
        }

        return `<td class="hs-cell">${escapeHtml(value)}</td>`;
    }

    function renderTable(data) {
        const tbody = el.tableBody();
        const empty = el.emptyState();
        const count = el.recordCount();
        const error = el.errorState();

        error.style.display = 'none';

        if (!data || data.length === 0) {
            tbody.innerHTML = '';
            empty.style.display = 'block';
            count.textContent = '0 signals';
            updateStats([]);
            return;
        }

        empty.style.display = 'none';
        count.textContent = `${data.length} signals`;

        tbody.innerHTML = data.map(row => {
            const cells = COLUMN_MAP.map(col => renderCell(row[col.key], col.key, row)).join('');
            return `<tr>${cells}</tr>`;
        }).join('');

        updateStats(data);
    }

    let _allData = [];
    let _activeFilter = 'all';

    function updateStats(data) {
        _allData = data;

        const newsCount = data.filter(r => r.companyRecentNews && r.companyRecentNews !== '-' && r.companyRecentNews.trim() !== '').length;
        const fundedCount = data.filter(r => r.fundingAmount && r.fundingAmount !== '-' && r.fundingAmount.trim() !== '').length;
        const jobChangeCount = data.filter(r =>
            (r.impactType && r.impactType.toLowerCase().includes('job')) ||
            (r.seniorLinkedin && r.seniorLinkedin !== '-' && r.seniorLinkedin.trim() !== '')
        ).length;
        const total = data.length;

        const pct = (n) => total > 0 ? `↑ ${Math.round(n / total * 100)}%` : '—';

        const set = (id, val) => { const e = document.getElementById(id); if (e) e.textContent = val; };
        set('hsStatNews', newsCount);
        set('hsStatFunded', fundedCount);
        set('hsStatJobChange', jobChangeCount);
        set('hsStatTotal', total);

        // Trend badges
        set('hsStatNewsTrend', pct(newsCount));
        set('hsStatFundedTrend', pct(fundedCount));
        set('hsStatJobChangeTrend', pct(jobChangeCount));

        // Remove skeleton loaders
        document.querySelectorAll('.hs-stat-card.loading').forEach(c => c.classList.remove('loading'));

        // Wire up click-to-filter (only on first call)
        document.querySelectorAll('.hs-stat-card[data-filter]').forEach(card => {
            card.onclick = () => {
                const filter = card.getAttribute('data-filter');
                _activeFilter = filter;
                document.querySelectorAll('.hs-stat-card').forEach(c => c.classList.remove('active'));
                card.classList.add('active');
                applyFilter(filter);
            };
        });
    }

    function applyFilter(filter) {
        if (!_allData.length) return;
        let filtered = _allData;
        if (filter === 'news') {
            filtered = _allData.filter(r => r.companyRecentNews && r.companyRecentNews !== '-' && r.companyRecentNews.trim() !== '');
        } else if (filter === 'funded') {
            filtered = _allData.filter(r => r.fundingAmount && r.fundingAmount !== '-' && r.fundingAmount.trim() !== '');
        } else if (filter === 'jobchange') {
            filtered = _allData.filter(r =>
                (r.impactType && r.impactType.toLowerCase().includes('job')) ||
                (r.seniorLinkedin && r.seniorLinkedin !== '-' && r.seniorLinkedin.trim() !== '')
            );
        }

        const tbody = el.tableBody();
        const count = el.recordCount();
        if (!tbody) return;
        tbody.innerHTML = filtered.map(row => {
            const cells = COLUMN_MAP.map(col => renderCell(row[col.key], col.key, row)).join('');
            return `<tr>${cells}</tr>`;
        }).join('');
        count.textContent = filter === 'all' ? `${filtered.length} signals` : `${filtered.length} of ${_allData.length} signals`;
    }


    // ── Skeleton loader ─────────────────────────────────────
    function showSkeleton() {
        const tbody = el.tableBody();
        const skeletonRows = Array(5).fill(null).map(() => {
            const cells = COLUMN_MAP.map(() =>
                `<td class="hs-cell"><div class="skeleton-line" style="width:${60 + Math.random() * 40}%;height:14px;border-radius:4px;background:rgba(0,0,0,0.06);animation:pulse 1.5s ease-in-out infinite"></div></td>`
            ).join('');
            return `<tr>${cells}</tr>`;
        }).join('');
        tbody.innerHTML = skeletonRows;
    }

    // ── Fetch & Load ────────────────────────────────────────
    async function loadData() {
        showSkeleton();
        el.emptyState().style.display = 'none';
        el.errorState().style.display = 'none';

        try {
            const res = await fetch(SHEET_URL);
            if (!res.ok) throw new Error(`HTTP ${res.status}`);
            const csv = await res.text();
            const rawRows = parseCSV(csv);
            const data = rawRows.map(mapRow).reverse(); // newest row first
            renderTable(data);
        } catch (err) {
            console.error('Hiring Signals fetch error:', err);
            el.tableBody().innerHTML = '';
            el.errorState().style.display = 'block';
        }
    }

    // ── Init ────────────────────────────────────────────────
    function init() {
        if (initialized) return;
        initialized = true;

        // Refresh button
        const btn = el.refreshBtn();
        if (btn) {
            btn.addEventListener('click', () => {
                btn.classList.add('spin');
                loadData().finally(() => setTimeout(() => btn.classList.remove('spin'), 600));
            });
        }

        // Modal Close Button
        const closeBtn = el.modalClose();
        if (closeBtn) {
            closeBtn.addEventListener('click', closeTextModal);
        }

        // Modal Outside Click
        const modal = el.modal();
        if (modal) {
            modal.addEventListener('click', (e) => {
                if (e.target === modal) closeTextModal();
            });
        }

        // Global Event Delegation for Buttons
        const tbody = el.tableBody();
        if (tbody) {
            tbody.addEventListener('click', (e) => {
                // Handle View / Email Button
                const viewBtn = e.target.closest('.btn-view');
                if (viewBtn) {
                    const encoded = viewBtn.getAttribute('data-full-text');
                    const title = viewBtn.getAttribute('data-modal-title');
                    const email = viewBtn.getAttribute('data-contact-email');
                    const type = viewBtn.getAttribute('data-type');

                    console.log('Signal View Clicked:', { title, encodedLength: encoded ? encoded.length : 0 });

                    try {
                        let content = '';
                        if (encoded) {
                            try {
                                content = decodeURIComponent(encoded);
                            } catch (e) {
                                console.warn('Failed to decode content, using raw:', e);
                                content = encoded;
                            }
                        } else {
                            content = 'No details available.';
                        }

                        console.log('Decoded Content Preview:', content.substring(0, 50) + '...');

                        if (type === 'employees') {
                            const employees = parseEmployees(content);
                            // Generate HTML for employee cards
                            const html = `<div class="hs-employee-list">${employees.map(renderEmployeeCard).join('')}</div>`;
                            openTextModal(title, html, null, true); // true = isHtml
                        } else {
                            openTextModal(title, content, email);
                        }
                    } catch (e) {
                        console.error('Error handling modal content:', e);
                        openTextModal('Error', 'Failed to load content.', null);
                    }
                    return;
                }

                // Handle Send Button (Direct)
                const sendBtn = e.target.closest('.btn-send-email');
                if (sendBtn) {
                    handleSendEmail(sendBtn);
                    return;
                }
            });
        }

        // Modal Footer delegation (Send Email)
        const footer = el.modalFooter();
        if (footer) {
            footer.addEventListener('click', (e) => {
                const btn = e.target.closest('.hs-btn-send');
                if (btn) {
                    handleSendEmail(btn); // Uses textarea content
                }
            });
        }

        // Initial load
        loadData();

        // Auto-refresh
        refreshTimer = setInterval(loadData, REFRESH_INTERVAL);
    }

    // Public API
    return { init, loadData };
})();
