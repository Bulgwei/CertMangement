const caSelect = document.getElementById('caSelect');
const refreshCaBtn = document.getElementById('refreshCaBtn');
const caStatus = document.getElementById('caStatus');
const templateSelect = document.getElementById('templateSelect');
const csrFile = document.getElementById('csrFile');
const dumpCsrBtn = document.getElementById('dumpCsrBtn');
const submitRequestBtn = document.getElementById('submitRequestBtn');
const csrPreview = document.getElementById('csrPreview');
const requesterName = document.getElementById('requesterName');
const timeWindow = document.getElementById('timeWindow');
const loadIssuedBtn = document.getElementById('loadIssuedBtn');
const outputFolder = document.getElementById('outputFolder');
const issuedTable = document.getElementById('issuedTable').querySelector('tbody');
const loadPendingBtn = document.getElementById('loadPendingBtn');
const pendingTable = document.getElementById('pendingTable').querySelector('tbody');
const activityLog = document.getElementById('activityLog');

let currentRequestFilePath = '';

async function log(message) {
    const timestamp = new Date().toLocaleString();
    activityLog.textContent += `[${timestamp}] ${message}\n`;
    activityLog.scrollTop = activityLog.scrollHeight;
}

async function loadCAList() {
    caStatus.textContent = 'Loading...';
    try {
        const response = await fetch('/api/ca/list');
        const cas = await response.json();
        caSelect.innerHTML = '';
        cas.forEach(ca => {
            const option = document.createElement('option');
            option.value = ca.configString;
            option.textContent = `${ca.caName} (${ca.dnsHostName})`;
            caSelect.appendChild(option);
        });
        if (cas.length > 0) {
            await updateCaStatus();
            await loadTemplates();
            log('CA list loaded.');
        } else {
            caStatus.textContent = 'No enrollment CAs found.';
            templateSelect.innerHTML = '<option value="">No CA available</option>';
        }
    } catch (error) {
        caStatus.textContent = 'Could not load CA list.';
        log('Error loading CA list: ' + error);
    }
}

async function updateCaStatus() {
    const config = caSelect.value;
    if (!config) return;

    const response = await fetch(`/api/ca/${encodeURIComponent(config)}/status`);
    const status = await response.json();
    const running = status.isResponding && status.hasAdmin;
    caStatus.textContent = running ? 'CA online and accessible' : 'CA not accessible';
    if (!running) {
        log('CA status indicates problems.');
    }
}

async function loadTemplates() {
    const config = caSelect.value;
    if (!config) return;

    try {
        const response = await fetch(`/api/ca/${encodeURIComponent(config)}/templates`);
        const templates = await response.json();
        templateSelect.innerHTML = templates.length
            ? templates.map(t => `<option value="${t}">${t}</option>`).join('')
            : '<option value="">No templates returned</option>';
        log('Template list loaded.');
    } catch (error) {
        templateSelect.innerHTML = '<option value="">Failed to load templates</option>';
        log('Error loading templates: ' + error);
    }
}

async function inspectCsr() {
    if (!csrFile.files.length) {
        alert('Please select a CSR file first.');
        return;
    }

    const formData = new FormData();
    formData.append('csr', csrFile.files[0]);

    try {
        const response = await fetch('/api/request/dump', {
            method: 'POST',
            body: formData
        });
        if (!response.ok) {
            const text = await response.text();
            throw new Error(text);
        }

        const payload = await response.json();
        currentRequestFilePath = payload.requestFilePath;
        const dump = payload.dump;
        csrPreview.textContent = `Subject CN: ${dump.subjectCommonName || 'N/A'}\n` +
            `Organization: ${dump.subjectOrganization || 'N/A'}\n` +
            `SAN: ${dump.sanEntries.length ? dump.sanEntries.join(', ') : 'None'}\n` +
            `Key Length: ${dump.keyLength || 'N/A'}\n` +
            `Key Algorithm: ${dump.keyAlgorithm || 'N/A'}\n` +
            `Signature Algorithm: ${dump.signatureAlgorithm || 'N/A'}\n` +
            `Signature Valid: ${dump.signatureMatchesPublicKey ? 'Yes' : 'No'}\n` +
            `Enhanced Key Usage: ${dump.enhancedKeyUsage.length ? dump.enhancedKeyUsage.join('; ') : 'N/A'}`;
        log('CSR inspected successfully.');
    } catch (error) {
        log('CSR inspection failed: ' + error);
        csrPreview.textContent = 'Unable to inspect CSR. See activity log.';
    }
}

async function submitRequest() {
    if (!currentRequestFilePath) {
        alert('Upload and inspect a CSR file first.');
        return;
    }

    const template = templateSelect.value;
    if (!template) {
        alert('Select a certificate template.');
        return;
    }

    const targetCa = caSelect.value;
    const body = {
        targetCaConfig: targetCa,
        templateName: template,
        requestFilePath: currentRequestFilePath
    };

    try {
        const response = await fetch('/api/request/submit', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(body)
        });
        const payload = await response.json();
        if (!response.ok) {
            throw new Error(payload?.message || response.statusText);
        }
        log('Submit result: ' + payload.message);
    } catch (error) {
        log('Submit request failed: ' + error);
    }
}

async function loadIssuedCertificates() {
    const targetCa = caSelect.value;
    const name = requesterName.value.trim();
    const days = parseInt(timeWindow.value, 10) || 7;
    const params = new URLSearchParams({
        targetCaConfig: targetCa,
        timeWindow: days.toString()
    });
    if (name) {
        params.append('requesterName', name);
    }

    try {
        const response = await fetch(`/api/certs?${params.toString()}`);
        const items = await response.json();
        issuedTable.innerHTML = items.map(item =>
            `<tr>
                <td>${item.requestId}</td>
                <td>${item.requesterName}</td>
                <td>${item.commonName}</td>
                <td>${item.san || ''}</td>
                <td><button data-request-id="${item.requestId}" class="retrieveBtn">Retrieve</button></td>
            </tr>`
        ).join('');
        bindIssuedActions();
        log('Issued certificate list loaded.');
    } catch (error) {
        log('Loading issued certificates failed: ' + error);
    }
}

async function bindIssuedActions() {
    document.querySelectorAll('.retrieveBtn').forEach(button => {
        button.onclick = async () => {
            const requestId = button.dataset.requestId;
            const folder = outputFolder.value.trim();
            if (!folder) {
                alert('Please provide an output folder.');
                return;
            }

            try {
                const response = await fetch('/api/cert/retrieve', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ targetCaConfig: caSelect.value, requestId, folderPath: folder })
                });
                const payload = await response.json();
                log(`Retrieve ${requestId}: ${payload.message}`);
            } catch (error) {
                log('Certificate retrieval failed: ' + error);
            }
        };
    });
}

async function loadPendingRequests() {
    try {
        const response = await fetch(`/api/pending?targetCaConfig=${encodeURIComponent(caSelect.value)}`);
        const items = await response.json();
        pendingTable.innerHTML = items.map(item =>
            `<tr>
                <td>${item.requestId}</td>
                <td>${item.requesterName}</td>
                <td>${item.commonName}</td>
                <td>${item.san || ''}</td>
                <td>
                    <button data-id="${item.requestId}" data-action="issue">Issue</button>
                    <button data-id="${item.requestId}" data-action="deny">Deny</button>
                </td>
            </tr>`
        ).join('');
        bindPendingActions();
        log('Pending request list loaded.');
    } catch (error) {
        log('Loading pending requests failed: ' + error);
    }
}

async function bindPendingActions() {
    document.querySelectorAll('#pendingTable button').forEach(button => {
        button.onclick = async () => {
            const requestId = button.dataset.id;
            const action = button.dataset.action;
            if (!requestId || !action) return;

            try {
                const response = await fetch(`/api/pending/${encodeURIComponent(requestId)}/${action}?targetCaConfig=${encodeURIComponent(caSelect.value)}`, {
                    method: 'POST'
                });
                const payload = await response.json();
                log(`${action} ${requestId}: ${payload.message}`);
                await loadPendingRequests();
            } catch (error) {
                log(`${action} request failed: ` + error);
            }
        };
    });
}

function switchTab(tabName) {
    // Hide all tab contents
    document.querySelectorAll('.tab-content').forEach(content => {
        content.classList.remove('active');
    });
    // Remove active class from all tab buttons
    document.querySelectorAll('.tab-button').forEach(button => {
        button.classList.remove('active');
    });
    // Show the selected tab content
    document.getElementById(tabName + '-tab').classList.add('active');
    // Add active class to the clicked tab button
    document.querySelector(`[data-tab="${tabName}"]`).classList.add('active');
}

refreshCaBtn.addEventListener('click', async () => {
    await loadCAList();
});
caSelect.addEventListener('change', async () => {
    await updateCaStatus();
    await loadTemplates();
});
dumpCsrBtn.addEventListener('click', inspectCsr);
submitRequestBtn.addEventListener('click', submitRequest);
loadIssuedBtn.addEventListener('click', loadIssuedCertificates);
loadPendingBtn.addEventListener('click', loadPendingRequests);

// Tab switching event listeners
document.querySelectorAll('.tab-button').forEach(button => {
    button.addEventListener('click', () => {
        const tabName = button.dataset.tab;
        switchTab(tabName);
    });
});

window.addEventListener('DOMContentLoaded', async () => {
    await loadCAList();
});
