/**
 * Job Change Detect Form Logic
 * Handles dynamic inputs, validation, and webhook submission.
 */

document.addEventListener('DOMContentLoaded', () => {
    const WEBHOOK_URL = 'https://n8n.smallgrp.com/webhook-test/2abc030e-272f-4f7f-a6e6-1dcbec6f7088';

    // DOM Elements
    const form = document.getElementById('jobChangeForm');
    const submitBtn = document.getElementById('jobChangeSubmitBtn');

    // Sections
    const urlContainer = document.getElementById('urlInputContainer');
    const emailContainer = document.getElementById('emailInputContainer');

    // Buttons
    const addUrlBtn = document.getElementById('addUrlBtn');
    const addEmailBtn = document.getElementById('addEmailBtn');
    const bulkPasteBtn = document.getElementById('bulkPasteBtn');

    // Modal
    const bulkModal = document.getElementById('bulkPasteModal');
    const closeBulkModal = document.getElementById('closeBulkPaste');
    const cancelBulk = document.getElementById('cancelBulkPaste');
    const confirmBulk = document.getElementById('confirmBulkPaste');
    const bulkTextArea = document.getElementById('bulkPasteArea');
    const urlCountSpan = document.getElementById('urlCount');

    // Select
    const signalSelect = document.getElementById('signalSelect');
    const signalDropdown = document.getElementById('signalDropdown');
    const signalTags = document.getElementById('signalTags');
    const signalInput = document.getElementById('signalInput');

    // State
    let selectedSignals = ['job_change'];
    const MAX_URLS = 1000;

    // --- 1. Dynamic Inputs Handling ---

    function createInputRow(type, value = '') {
        const div = document.createElement('div');
        div.className = 'input-row fade-in';

        const input = document.createElement('input');
        input.type = type === 'email' ? 'email' : 'text';
        input.className = `form-input ${type === 'email' ? 'email-input' : 'linkedin-url-input'}`;
        input.placeholder = type === 'email' ? 'recipient@example.com' : 'https://www.linkedin.com/in/username/';
        input.value = value;
        input.required = true;

        const removeBtn = document.createElement('button');
        removeBtn.type = 'button';
        removeBtn.className = 'btn-icon-remove';
        removeBtn.innerHTML = '×';
        removeBtn.title = 'Remove';

        removeBtn.addEventListener('click', () => {
            if (div.parentElement.children.length > 1) {
                div.remove();
                validateForm();
            }
        });

        input.addEventListener('input', validateForm);
        input.addEventListener('blur', () => {
            if (type === 'url') validateLinkedInUrl(input);
            else validateEmail(input);
        });

        div.appendChild(input);
        div.appendChild(removeBtn);
        return div;
    }

    function updateRemoveButtons(container) {
        const buttons = container.querySelectorAll('.btn-icon-remove');
        const disabled = buttons.length === 1;
        buttons.forEach(btn => btn.disabled = disabled);
    }

    addUrlBtn.addEventListener('click', () => {
        urlContainer.appendChild(createInputRow('url'));
        updateRemoveButtons(urlContainer);
        validateForm();
    });

    addEmailBtn.addEventListener('click', () => {
        emailContainer.appendChild(createInputRow('email'));
        updateRemoveButtons(emailContainer);
        validateForm();
    });

    // Observer to update remove buttons when children change
    [urlContainer, emailContainer].forEach(container => {
        const observer = new MutationObserver(() => updateRemoveButtons(container));
        observer.observe(container, { childList: true });
    });

    // --- 2. Validation Logic ---

    function validateLinkedInUrl(input) {
        const regex = /^https:\/\/www\.linkedin\.com\/in\/[\w-]+\/?$/;
        const isValid = regex.test(input.value.trim());
        setValidity(input, isValid);
        return isValid;
    }

    function validateEmail(input) {
        const regex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
        const isValid = regex.test(input.value.trim());
        setValidity(input, isValid);
        return isValid;
    }

    function setValidity(input, isValid) {
        if (!isValid && input.value.trim() !== '') {
            input.classList.add('invalid');
        } else {
            input.classList.remove('invalid');
        }
    }

    function validateForm() {
        // Debounce slightly
        setTimeout(() => {
            const urls = Array.from(document.querySelectorAll('.linkedin-url-input'));
            const emails = Array.from(document.querySelectorAll('.email-input'));

            const urlsValid = urls.every(i => i.value.trim() !== '' && !i.classList.contains('invalid'));
            const emailsValid = emails.every(i => i.value.trim() !== '' && !i.classList.contains('invalid'));
            const signalsValid = selectedSignals.length > 0;

            submitBtn.disabled = !(urlsValid && emailsValid && signalsValid);
        }, 50);
    }

    // --- 3. Multi-select Dropdown ---

    signalSelect.addEventListener('click', (e) => {
        if (e.target.classList.contains('remove-tag')) return;
        signalDropdown.classList.toggle('active');
    });

    // Close dropdown when clicking outside
    document.addEventListener('click', (e) => {
        if (!signalSelect.contains(e.target) && !signalDropdown.contains(e.target)) {
            signalDropdown.classList.remove('active');
        }
    });

    signalDropdown.addEventListener('click', (e) => {
        const option = e.target.closest('.dropdown-option');
        if (option) {
            const value = option.dataset.value;
            toggleSignal(value);
        }
    });

    function toggleSignal(value) {
        const index = selectedSignals.indexOf(value);
        if (index === -1) {
            selectedSignals.push(value);
        } else {
            selectedSignals.splice(index, 1);
        }
        renderSignals();
        validateForm();
    }

    function renderSignals() {
        // Update tags
        signalTags.innerHTML = '';
        selectedSignals.forEach(val => {
            const tag = document.createElement('span');
            tag.className = 'tag';
            tag.innerHTML = `${val} <span class="remove-tag">×</span>`;
            tag.querySelector('.remove-tag').onclick = (e) => {
                e.stopPropagation();
                toggleSignal(val);
            };
            signalTags.appendChild(tag);
        });

        // Update dropdown classes
        document.querySelectorAll('.dropdown-option').forEach(opt => {
            if (selectedSignals.includes(opt.dataset.value)) {
                opt.classList.add('selected');
            } else {
                opt.classList.remove('selected');
            }
        });

        signalInput.value = selectedSignals.join(',');
    }

    // --- 4. Bulk Paste Modal ---

    bulkPasteBtn.addEventListener('click', () => bulkModal.classList.add('active'));

    [closeBulkModal, cancelBulk].forEach(btn => {
        btn.addEventListener('click', () => {
            bulkModal.classList.remove('active');
            bulkTextArea.value = '';
            urlCountSpan.innerText = '0';
        });
    });

    bulkTextArea.addEventListener('input', () => {
        const text = bulkTextArea.value;
        // Match approximate URLs
        const matches = text.match(/https:\/\/www\.linkedin\.com\/in\/[\w-]+\/?/g) || [];
        urlCountSpan.innerText = matches.length;
    });

    confirmBulk.addEventListener('click', () => {
        const text = bulkTextArea.value;
        const rawLines = text.split(/[\n,]/).map(s => s.trim()).filter(s => s !== '');

        let validUrls = [];
        const regex = /^https:\/\/www\.linkedin\.com\/in\/[\w-]+\/?$/;

        rawLines.forEach(line => {
            if (regex.test(line)) {
                validUrls.push(line);
            }
        });

        // Dedup
        validUrls = [...new Set(validUrls)];

        if (validUrls.length > 0) {
            // Remove existing empty inputs
            const existing = Array.from(urlContainer.querySelectorAll('.linkedin-url-input'));
            existing.forEach(input => {
                if (input.value.trim() === '') input.parentElement.remove();
            });

            // Add new
            validUrls.forEach(url => {
                urlContainer.appendChild(createInputRow('url', url));
            });

            updateRemoveButtons(urlContainer);
            validateForm();
            bulkModal.classList.remove('active');
            bulkTextArea.value = '';
        }
    });

    // --- 5. Submission ---

    form.addEventListener('submit', async (e) => {
        e.preventDefault();

        const btnText = submitBtn.querySelector('.btn-text');
        const loader = submitBtn.querySelector('.btn-loader');

        // Prepare Data
        const profile_urls = Array.from(document.querySelectorAll('.linkedin-url-input'))
            .map(i => i.value.trim())
            .filter(v => v);

        const email_recipients = Array.from(document.querySelectorAll('.email-input'))
            .map(i => i.value.trim())
            .filter(v => v);

        // Dedup lists
        const uniqueUrls = [...new Set(profile_urls)];
        const uniqueEmails = [...new Set(email_recipients)];

        const payload = {
            email_recipients: uniqueEmails,
            profile_urls: uniqueUrls,
            signal_configurations: selectedSignals
        };

        // UI Loading
        submitBtn.disabled = true;
        btnText.style.display = 'none';
        loader.style.display = 'inline-block';

        try {
            const response = await fetch(WEBHOOK_URL, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify(payload)
            });

            if (response.ok) {
                const data = await response.json();

                // Assuming the response is an array as per user request
                const result = Array.isArray(data) ? data[0] : data;

                if (result) {
                    showToast('Tracking started successfully', 'success');
                    displayResults(result);
                } else {
                    showToast('Tracking started, but no data returned', 'warning');
                }
            } else {
                showToast('Failed to activate tracking. Please try again.', 'error');
            }

        } catch (error) {
            console.error('Submission error:', error);
            showToast('Network error. Check console for details.', 'error');
        } finally {
            submitBtn.disabled = false;
            btnText.style.display = 'inline';
            loader.style.display = 'none';
        }
    });

    // --- 6. Result Display ---

    const resultContainer = document.getElementById('jobChangeResult');
    const resetFormBtn = document.getElementById('resetFormBtn');

    // Metrics Elements
    const resTimestamp = document.getElementById('resultTimestamp');
    const resProfiles = document.getElementById('resProfiles');
    const resChanges = document.getElementById('resChanges');
    const resRunId = document.getElementById('resRunId');
    const resErrorContainer = document.getElementById('resErrorContainer');
    const resErrorList = document.getElementById('resErrorList');

    function displayResults(data) {
        // Hide Form
        form.style.display = 'none';
        document.querySelector('.job-change-header').style.display = 'none';

        // Show Result
        resultContainer.style.display = 'block';

        // Populate Data and Handle Logic
        if (data.timestamp) {
            const date = new Date(data.timestamp);
            resTimestamp.innerText = date.toLocaleString();
        }

        const changes = data.changes_detected !== undefined ? data.changes_detected : 0;
        const profiles = data.profiles_processed !== undefined ? data.profiles_processed : 0;

        // Populate standard metrics
        resProfiles.innerText = profiles;
        resChanges.innerText = changes;
        resRunId.innerText = data.run_id || '-';

        // Conditional Display
        const headerTitle = resultContainer.querySelector('.result-header h4');
        const metricsGrid = resultContainer.querySelector('.metrics-grid');

        // Clear any existing custom message
        const existingMsg = resultContainer.querySelector('.no-change-message');
        if (existingMsg) existingMsg.remove();

        if (changes > 0) {
            // Show standard metrics
            headerTitle.innerHTML = 'Changes Detected Successfully';
            // Ensure icon is checkmark (default via CSS)
            metricsGrid.style.display = 'grid';
        } else {
            // Show "No changes detected" message
            headerTitle.innerHTML = 'Tracking Started';
            metricsGrid.style.display = 'none';

            const msgDiv = document.createElement('div');
            msgDiv.className = 'no-change-message';
            msgDiv.innerHTML = `
                <div class="info-icon">i</div>
                <p>No immediate job changes detected.</p>
                <p class="sub-text">You will be notified on your email ID if a job change occurs.</p>
            `;
            // Insert after header
            resultContainer.querySelector('.result-header').after(msgDiv);
        }

        // Handle Errors
        if (data.errors && data.errors.length > 0) {
            resErrorContainer.style.display = 'block';
            resErrorList.innerHTML = data.errors.map(err => `<li>${err}</li>`).join('');
        } else {
            resErrorContainer.style.display = 'none';
        }
    }

    resetFormBtn.addEventListener('click', () => {
        // Reset View
        resultContainer.style.display = 'none';
        form.style.display = 'flex';
        document.querySelector('.job-change-header').style.display = 'flex';

        // Reset Form
        form.reset();
        urlContainer.innerHTML = '';
        emailContainer.innerHTML = '';
        urlContainer.appendChild(createInputRow('url'));
        emailContainer.appendChild(createInputRow('email'));
        selectedSignals = ['job_change'];
        renderSignals();
        validateForm();
    });

    // Toast Notification Helper
    function showToast(message, type = 'success') {
        const toast = document.createElement('div');
        toast.className = `form-message ${type}`;
        toast.style.position = 'fixed';
        toast.style.bottom = '24px';
        toast.style.right = '24px';
        toast.style.zIndex = '9999';
        toast.style.background = type === 'success' ? '#fff' : '#fff';
        toast.style.color = type === 'success' ? 'var(--success)' : 'var(--error)';
        toast.style.border = `1px solid ${type === 'success' ? 'var(--success)' : 'var(--error)'}`;
        toast.style.boxShadow = '0 4px 12px rgba(0,0,0,0.1)';
        toast.innerText = message;

        document.body.appendChild(toast);

        setTimeout(() => {
            toast.style.opacity = '0';
            setTimeout(() => toast.remove(), 300);
        }, 3000);
    }
});
