// Admissions Database Frontend Application Logic

// Constants
const LOCAL_STORAGE_KEY = 'eub_admissions_records';
const THEME_STORAGE_KEY = 'eub_admissions_theme';
const AUTH_STORAGE_KEY = 'eub_admissions_m365_account';
const LOGIN_SCOPES = ['openid', 'profile', 'email'];

const AUTH_CONFIG = window.AUTH_CONFIG || {
    clientId: '',
    tenantId: 'common',
    redirectUri: window.location.origin + window.location.pathname,
    postLogoutRedirectUri: window.location.origin + window.location.pathname,
    cacheLocation: 'localStorage'
};

const msalConfig = {
    auth: {
        clientId: AUTH_CONFIG.clientId,
        authority: `https://login.microsoftonline.com/${AUTH_CONFIG.tenantId || 'common'}`,
        redirectUri: AUTH_CONFIG.redirectUri,
        postLogoutRedirectUri: AUTH_CONFIG.postLogoutRedirectUri,
        navigateToLoginRequestUrl: false
    },
    cache: {
        cacheLocation: AUTH_CONFIG.cacheLocation || 'localStorage',
        storeAuthStateInCookie: false
    }
};

let msalInstance = null;

// Application State
let state = {
    records: [],
    filteredRecords: [],
    currentPage: 1,
    pageSize: 15,
    sortColumn: 'App Serial No.',
    sortOrder: 'desc', // desc to show newest Serial Numbers first
    currentRecordId: null, // Holds the Serial No of editing record, null for new record
    theme: 'light-theme'
};

// Section mapping from metadata groups to HTML element IDs
const SECTION_MAPPING = {
    "Applicant, Source and Parent / Guardian Data": "grid-sec-applicant",
    "Contact Verification": "grid-sec-verification",
    "Offer Readiness": "grid-sec-readiness",
    "Offer Decision and Offer Issue": "grid-sec-decision",
    "Conversion and Engagement": "grid-sec-engagement",
    "Funding Processing": "grid-sec-funding",
    "Registration": "grid-sec-registration",
    "Admin / Audit": "grid-sec-audit"
};

// Initialization
document.addEventListener('DOMContentLoaded', () => {
    initializeAuthFlow();
});

function initializeAuthFlow() {
    loadTheme();

    if (!window.msal || !AUTH_CONFIG.clientId) {
        showAuthSetupMessage();
        return;
    }

    msalInstance = new msal.PublicClientApplication(msalConfig);
    handleAuthRedirect();
    bindAuthButtons();
}

async function handleAuthRedirect() {
    try {
        const response = await msalInstance.handleRedirectPromise();
        if (response && response.account) {
            msalInstance.setActiveAccount(response.account);
            finalizeSignIn(response.account);
            return;
        }

        const activeAccount = msalInstance.getActiveAccount() || msalInstance.getAllAccounts()[0];
        if (activeAccount) {
            msalInstance.setActiveAccount(activeAccount);
            finalizeSignIn(activeAccount);
            return;
        }

        showAuthScreen();
    } catch (error) {
        console.error('Microsoft sign-in initialization failed', error);
        showAuthError('Microsoft sign-in could not be initialized. Check the tenant and client ID configuration.');
    }
}

function bindAuthButtons() {
    const loginButton = document.getElementById('microsoftLoginBtn');
    if (loginButton) {
        loginButton.addEventListener('click', signInWithMicrosoft);
    }

    const signOutButton = document.getElementById('signOutBtn');
    if (signOutButton) {
        signOutButton.addEventListener('click', signOutWithMicrosoft);
    }
}

function showAuthScreen() {
    document.getElementById('authScreen').classList.add('visible');
    document.getElementById('appShell').classList.add('app-hidden');
    document.getElementById('appShell').setAttribute('aria-hidden', 'true');
    document.body.classList.add('auth-only');
}

function showAuthSetupMessage() {
    showAuthScreen();
    const status = document.getElementById('authStatusText');
    if (status) {
        status.textContent = 'Microsoft sign-in is not configured yet. Add your Entra app settings in auth-config.js.';
    }
    const loginButton = document.getElementById('microsoftLoginBtn');
    if (loginButton) {
        loginButton.disabled = true;
    }
}

function showAuthError(message) {
    showAuthScreen();
    const status = document.getElementById('authStatusText');
    if (status) {
        status.textContent = message;
    }
}

function finalizeSignIn(account) {
    document.getElementById('authScreen').classList.remove('visible');
    document.getElementById('appShell').classList.remove('app-hidden');
    document.getElementById('appShell').setAttribute('aria-hidden', 'false');
    document.body.classList.remove('auth-only');

    const status = document.getElementById('authStatusText');
    if (status) {
        status.textContent = `Signed in as ${account.name || account.username || 'Microsoft user'}.`;
    }

    localStorage.setItem(AUTH_STORAGE_KEY, JSON.stringify({
        name: account.name || '',
        username: account.username || '',
        homeAccountId: account.homeAccountId || ''
    }));

    loadRecords();
    initFilterDropdowns();
    renderFormFields();
    setupEventListeners();
    updateUI();
}

async function signInWithMicrosoft() {
    if (!msalInstance) return;

    const status = document.getElementById('authStatusText');
    if (status) {
        status.textContent = 'Opening Microsoft sign-in...';
    }

    try {
        await msalInstance.loginRedirect({
            scopes: LOGIN_SCOPES,
            prompt: 'select_account'
        });
    } catch (error) {
        console.error('Microsoft sign-in failed', error);
        showAuthError('Microsoft sign-in failed. Verify the Entra app registration and redirect URI.');
    }
}

async function signOutWithMicrosoft() {
    if (!msalInstance) return;

    try {
        localStorage.removeItem(AUTH_STORAGE_KEY);
        await msalInstance.logoutRedirect({
            account: msalInstance.getActiveAccount() || msalInstance.getAllAccounts()[0]
        });
    } catch (error) {
        console.error('Microsoft sign-out failed', error);
        showAuthError('Sign-out failed. Please refresh and try again.');
    }
}

// Load theme from localStorage
function loadTheme() {
    const savedTheme = localStorage.getItem(THEME_STORAGE_KEY) || 'light-theme';
    state.theme = savedTheme;
    document.body.className = savedTheme;
    if (!localStorage.getItem(THEME_STORAGE_KEY)) {
        localStorage.setItem(THEME_STORAGE_KEY, savedTheme);
    }
}

// Toggle Theme between dark and light
function toggleTheme() {
    if (state.theme === 'dark-theme') {
        state.theme = 'light-theme';
    } else {
        state.theme = 'dark-theme';
    }
    document.body.className = state.theme;
    localStorage.setItem(THEME_STORAGE_KEY, state.theme);
}

// Load records from local storage or fallback to spreadsheet data
function loadRecords() {
    const stored = localStorage.getItem(LOCAL_STORAGE_KEY);
    if (stored) {
        try {
            state.records = JSON.parse(stored);
        } catch (e) {
            console.error("Error parsing stored records, resetting to seed data", e);
            state.records = [...ADMISSIONS_DATA.records];
        }
    } else {
        state.records = [...ADMISSIONS_DATA.records];
    }
    state.filteredRecords = [...state.records];
}

// Save current records state to local storage
function saveRecords() {
    localStorage.setItem(LOCAL_STORAGE_KEY, JSON.stringify(state.records));
}

// Initialize dynamic filters based on sheet contents
function initFilterDropdowns() {
    // 1. Pipeline Stages
    const stages = ADMISSIONS_DATA.dropdowns['Current Pipeline Stage'] || [];
    const filterStage = document.getElementById('filterStage');
    stages.forEach(stg => {
        const opt = document.createElement('option');
        opt.value = stg;
        opt.textContent = stg;
        filterStage.appendChild(opt);
    });

    // 2. Advisors (Unique values from data)
    const advisors = [...new Set(state.records.map(r => r['Advisor / Admissions Owner']).filter(Boolean))].sort();
    const filterAdvisor = document.getElementById('filterAdvisor');
    advisors.forEach(adv => {
        const opt = document.createElement('option');
        opt.value = adv;
        opt.textContent = adv;
        filterAdvisor.appendChild(opt);
    });

    // 3. Offer Statuses
    const offerStatuses = ADMISSIONS_DATA.dropdowns['Offer Status'] || [];
    const filterOffer = document.getElementById('filterOffer');
    offerStatuses.forEach(st => {
        const opt = document.createElement('option');
        opt.value = st;
        opt.textContent = st;
        filterOffer.appendChild(opt);
    });
}

// Generate the 88 edit fields dynamically in their respective tabs inside the drawer
function renderFormFields() {
    ADMISSIONS_DATA.columns.forEach(col => {
        const sectionId = SECTION_MAPPING[col.group];
        if (!sectionId) return;

        const container = document.getElementById(sectionId);
        if (!container) return;

        const formGroup = document.createElement('div');
        formGroup.className = 'form-group';
        
        // Full width fields for notes and long text
        if (col.name.toLowerCase().includes('note') || col.name.toLowerCase().includes('comment') || col.name.toLowerCase().includes('outstanding')) {
            formGroup.classList.add('full-width');
        }

        const label = document.createElement('label');
        label.setAttribute('for', `fld_${col.name}`);
        label.textContent = col.name;
        formGroup.appendChild(label);

        let inputEl;

        // Custom field configurations based on type
        if (col.type === 'select') {
            inputEl = document.createElement('select');
            inputEl.id = `fld_${col.name}`;
            inputEl.name = col.name;
            
            // Add empty option
            const emptyOpt = document.createElement('option');
            emptyOpt.value = "";
            emptyOpt.textContent = "-- Select --";
            inputEl.appendChild(emptyOpt);

            const options = ADMISSIONS_DATA.dropdowns[col.name] || [];
            options.forEach(optVal => {
                const opt = document.createElement('option');
                opt.value = optVal;
                opt.textContent = optVal;
                inputEl.appendChild(opt);
            });
        } else if (col.name.toLowerCase().includes('note') || col.name.toLowerCase().includes('comment') || col.name.toLowerCase().includes('grades')) {
            inputEl = document.createElement('textarea');
            inputEl.id = `fld_${col.name}`;
            inputEl.name = col.name;
            inputEl.rows = 2;
        } else {
            inputEl = document.createElement('input');
            inputEl.id = `fld_${col.name}`;
            inputEl.name = col.name;
            
            if (col.type === 'date') {
                inputEl.type = 'date';
            } else if (col.type === 'number') {
                inputEl.type = 'number';
            } else {
                inputEl.type = 'text';
            }
        }

        // Configure read-only status for system / audit fields
        const readOnlyFields = [
            'App Serial No.', 'Created', 'Created By', 'Modified', 'Modified By', 
            'Current Pipeline Stage Date', 'Registration Handover Date', 'Registration Date', 'Parent / Guardian Consent Date'
        ];
        if (readOnlyFields.includes(col.name)) {
            inputEl.setAttribute('readonly', true);
            if (inputEl.tagName === 'SELECT') {
                inputEl.setAttribute('disabled', true);
            }
        }

        formGroup.appendChild(inputEl);
        container.appendChild(formGroup);
    });
}

// Update the entire UI elements (KPIs, Charts, Table rows)
function updateUI() {
    calculateKPIs();
    renderCharts();
    filterAndSortRecords();
    renderTable();
}

// Compute dashboard statistics dynamically
function calculateKPIs() {
    const totalCount = state.records.length;
    document.getElementById('kpiTotalCount').textContent = totalCount;

    // Contact Verification Gate: Email & Mobile both Verified
    const verifiedCount = state.records.filter(r => 
        r['Mobile Verification Status'] === 'Verified' && 
        r['Email Verification Status'] === 'Verified'
    ).length;
    document.getElementById('kpiVerifiedCount').textContent = verifiedCount;
    const verifiedPercent = totalCount > 0 ? Math.round((verifiedCount / totalCount) * 100) : 0;
    document.getElementById('kpiVerifiedPercent').textContent = `${verifiedPercent}% Fully Verified`;

    // Offers Issued: Conditional or Unconditional
    const offersCount = state.records.filter(r => 
        r['Offer Status'] === 'Conditional' || 
        r['Offer Status'] === 'Unconditional'
    ).length;
    document.getElementById('kpiOffersCount').textContent = offersCount;
    const offersPercent = totalCount > 0 ? Math.round((offersCount / totalCount) * 100) : 0;
    document.getElementById('kpiOffersPercent').textContent = `${offersPercent}% of applicants`;

    // Registered Students
    const registeredCount = state.records.filter(r => r['Registration Status'] === 'Registered').length;
    document.getElementById('kpiRegisteredCount').textContent = registeredCount;
    const registeredPercent = totalCount > 0 ? Math.round((registeredCount / totalCount) * 100) : 0;
    document.getElementById('kpiRegistrationPercent').textContent = `${registeredPercent}% Conversion Rate`;
}

// Render dynamic visual SVG charts
function renderCharts() {
    // 1. Pipeline Stages volume chart
    const stageContainer = document.getElementById('pipelineStageChart');
    stageContainer.innerHTML = '';
    
    // Group records by stage
    const stages = ADMISSIONS_DATA.dropdowns['Current Pipeline Stage'] || [];
    const stageCounts = {};
    stages.forEach(stg => stageCounts[stg] = 0);
    state.records.forEach(r => {
        const stage = r['Current Pipeline Stage'];
        if (stages.includes(stage)) {
            stageCounts[stage]++;
        } else if (stage) {
            stageCounts[stage] = (stageCounts[stage] || 0) + 1;
        }
    });

    const maxCount = Math.max(...Object.values(stageCounts), 1);
    
    // Mapping color themes for the visual charts
    const stageColors = {
        'Applicant Data Incomplete': 'var(--stage-applicant)',
        'Contact Verification': 'var(--stage-verification)',
        'Contact Verification Dormant': 'var(--stage-verification)',
        'Offer Readiness': 'var(--stage-readiness)',
        'Offer In Process': 'var(--stage-decision)',
        'Conditional Offer Issued': 'var(--stage-decision)',
        'Unconditional Offer Issued': 'var(--stage-decision)',
        'Conversion': 'var(--stage-engagement)',
        'Registration In Process': 'var(--stage-registration)',
        'Registered': 'var(--stage-registration)',
        'Withdrawn': 'var(--stage-audit)',
        'Rejected': 'var(--stage-audit)',
        'Duplicate / Invalid': 'var(--stage-audit)'
    };

    Object.keys(stageCounts).forEach(stage => {
        const count = stageCounts[stage];
        const pct = (count / maxCount) * 80 + 10; // Scaled between 10% and 90% for look

        const col = document.createElement('div');
        col.className = 'chart-bar-col';
        
        const bar = document.createElement('div');
        bar.className = 'chart-bar-fill';
        bar.style.height = `${pct}%`;
        bar.style.backgroundColor = stageColors[stage] || 'var(--color-primary)';
        
        const tooltip = document.createElement('span');
        tooltip.className = 'chart-bar-tooltip';
        tooltip.textContent = `${stage}: ${count}`;
        bar.appendChild(tooltip);

        const label = document.createElement('span');
        label.className = 'chart-bar-label';
        label.textContent = stage;
        label.title = stage;

        col.appendChild(bar);
        col.appendChild(label);
        stageContainer.appendChild(col);
    });

    // 2. Advisor allocation chart
    const advisorContainer = document.getElementById('advisorWorkloadChart');
    advisorContainer.innerHTML = '';

    // Group records by Advisor
    const advisors = [...new Set(state.records.map(r => r['Advisor / Admissions Owner']).filter(Boolean))];
    const advisorCounts = {};
    advisors.forEach(adv => advisorCounts[adv] = 0);
    state.records.forEach(r => {
        const adv = r['Advisor / Admissions Owner'];
        if (adv) advisorCounts[adv]++;
    });

    const sortedAdvisors = Object.entries(advisorCounts).sort((a, b) => b[1] - a[1]);
    const maxAdvisorCount = Math.max(...Object.values(advisorCounts), 1);

    sortedAdvisors.forEach(([adv, count]) => {
        const pct = (count / maxAdvisorCount) * 100;
        
        const row = document.createElement('div');
        row.style.width = '100%';
        row.style.marginBottom = '0.5rem';
        
        const info = document.createElement('div');
        info.style.display = 'flex';
        info.style.justifyContent = 'space-between';
        info.style.fontSize = '0.75rem';
        info.style.marginBottom = '2px';
        info.style.fontWeight = '500';
        info.innerHTML = `<span>${adv}</span><span style="color:var(--text-secondary)">${count} Records</span>`;
        
        const progressContainer = document.createElement('div');
        progressContainer.style.width = '100%';
        progressContainer.style.height = '8px';
        progressContainer.style.backgroundColor = 'var(--border-color)';
        progressContainer.style.borderRadius = '4px';
        progressContainer.style.overflow = 'hidden';
        
        const bar = document.createElement('div');
        bar.style.width = `${pct}%`;
        bar.style.height = '100%';
        bar.style.background = 'linear-gradient(90deg, var(--color-primary), var(--stage-verification))';
        bar.style.transition = 'width 0.6s ease';
        
        progressContainer.appendChild(bar);
        row.appendChild(info);
        row.appendChild(progressContainer);
        advisorContainer.appendChild(row);
    });
}

// Perform client-side search, filtering and sorting
function filterAndSortRecords() {
    const searchVal = document.getElementById('searchInput').value.toLowerCase().trim();
    const stageVal = document.getElementById('filterStage').value;
    const advisorVal = document.getElementById('filterAdvisor').value;
    const offerVal = document.getElementById('filterOffer').value;
    const verificationVal = document.getElementById('filterVerification').value;

    state.filteredRecords = state.records.filter(r => {
        // 1. Text Search
        const name = `${r['First Name'] || ''} ${r['Last Name'] || ''}`.toLowerCase();
        const serial = (r['App Serial No.'] || '').toLowerCase();
        const nationality = (r['Nationality'] || '').toLowerCase();
        const email = (r['Email Address'] || '').toLowerCase();
        const advisor = (r['Advisor / Admissions Owner'] || '').toLowerCase();
        const searchMatches = !searchVal || 
            name.includes(searchVal) || 
            serial.includes(searchVal) || 
            nationality.includes(searchVal) || 
            email.includes(searchVal) || 
            advisor.includes(searchVal);

        // 2. Exact Filters
        const stageMatches = !stageVal || r['Current Pipeline Stage'] === stageVal;
        const advisorMatches = !advisorVal || r['Advisor / Admissions Owner'] === advisorVal;
        const offerMatches = !offerVal || r['Offer Status'] === offerVal;
        
        // 3. Verification blocker logic (Blocked = Verification Blocker has a value, or mobile/email is not verified and current stage is verification/offer)
        let verificationMatches = true;
        if (verificationVal === 'Blocked') {
            verificationMatches = r['Verification Blocker'] || 
                (r['Mobile Verification Status'] !== 'Verified' || r['Email Verification Status'] !== 'Verified');
        } else if (verificationVal === 'Unblocked') {
            verificationMatches = !r['Verification Blocker'] && 
                r['Mobile Verification Status'] === 'Verified' && 
                r['Email Verification Status'] === 'Verified';
        }

        return searchMatches && stageMatches && advisorMatches && offerMatches && verificationMatches;
    });

    // Sort records
    const colName = state.sortColumn;
    const isAsc = state.sortOrder === 'asc';

    state.filteredRecords.sort((a, b) => {
        let valA = a[colName] || '';
        let valB = b[colName] || '';

        // If numeric type
        if (!isNaN(valA) && !isNaN(valB) && valA !== '' && valB !== '') {
            return isAsc ? Number(valA) - Number(valB) : Number(valB) - Number(valA);
        }

        valA = String(valA).toLowerCase();
        valB = String(valB).toLowerCase();

        if (valA < valB) return isAsc ? -1 : 1;
        if (valA > valB) return isAsc ? 1 : -1;
        return 0;
    });
}

// Render dynamic rows of the grid table
function renderTable() {
    const tableBody = document.getElementById('tableBody');
    tableBody.innerHTML = '';

    const totalRecords = state.filteredRecords.length;
    const totalPages = Math.max(Math.ceil(totalRecords / state.pageSize), 1);
    
    // Boundary check for current page
    if (state.currentPage > totalPages) {
        state.currentPage = totalPages;
    }
    if (state.currentPage < 1) {
        state.currentPage = 1;
    }

    // Pagination slice
    const startIndex = (state.currentPage - 1) * state.pageSize;
    const endIndex = Math.min(startIndex + state.pageSize, totalRecords);
    const pageRecords = state.filteredRecords.slice(startIndex, endIndex);

    // Update Pagination footer display
    document.getElementById('resultCountText').textContent = `Showing ${totalRecords > 0 ? startIndex + 1 : 0} to ${endIndex} of ${totalRecords} records`;
    document.getElementById('pageNumberDisplay').textContent = `Page ${state.currentPage} of ${totalPages}`;
    document.getElementById('prevPageBtn').disabled = state.currentPage === 1;
    document.getElementById('nextPageBtn').disabled = state.currentPage === totalPages;

    if (pageRecords.length === 0) {
        const row = document.createElement('tr');
        row.innerHTML = `<td colspan="11" style="text-align: center; padding: 2rem; color: var(--text-muted);">No records found matching filters.</td>`;
        tableBody.appendChild(row);
        return;
    }

    pageRecords.forEach(r => {
        const tr = document.createElement('tr');
        tr.dataset.id = r['App Serial No.'];

        // Color coding current stage badge
        let stageBadgeClass = 'badge-stage-audit';
        const stg = r['Current Pipeline Stage'] || '';
        if (stg.includes('Applicant')) stageBadgeClass = 'badge-stage-applicant';
        else if (stg.includes('Verification')) stageBadgeClass = 'badge-stage-verification';
        else if (stg.includes('Readiness')) stageBadgeClass = 'badge-stage-readiness';
        else if (stg.includes('Offer')) stageBadgeClass = 'badge-stage-decision';
        else if (stg.includes('Conversion')) stageBadgeClass = 'badge-stage-engagement';
        else if (stg.includes('Funding')) stageBadgeClass = 'badge-stage-funding';
        else if (stg.includes('Registration')) stageBadgeClass = 'badge-stage-registration';

        // Verification Badge builder
        const mobVer = r['Mobile Verification Status'] === 'Verified' ? 
            `<span class="badge badge-success"><i class="fa-solid fa-circle-check"></i> Verified</span>` : 
            `<span class="badge badge-muted">Unverified</span>`;
            
        const emVer = r['Email Verification Status'] === 'Verified' ? 
            `<span class="badge badge-success"><i class="fa-solid fa-circle-check"></i> Verified</span>` : 
            `<span class="badge badge-muted">Unverified</span>`;

        // Offer Status styling
        let offerBadge = `<span class="badge badge-muted">${r['Offer Status'] || 'None'}</span>`;
        if (r['Offer Status'] === 'Conditional') offerBadge = `<span class="badge badge-warning">Conditional</span>`;
        else if (r['Offer Status'] === 'Unconditional') offerBadge = `<span class="badge badge-success">Unconditional</span>`;
        else if (r['Offer Status'] === 'Rejected') offerBadge = `<span class="badge badge-danger">Rejected</span>`;

        // Registration badge
        let regBadge = `<span class="badge badge-muted">${r['Registration Status'] || 'Pending'}</span>`;
        if (r['Registration Status'] === 'Registered') regBadge = `<span class="badge badge-success"><i class="fa-solid fa-square-check"></i> Registered</span>`;
        else if (r['Registration Status'] === 'Registration In Process') regBadge = `<span class="badge badge-info">In Process</span>`;

        // Populate table row
        tr.innerHTML = `
            <td style="font-weight: 600; color: var(--color-primary)">${r['App Serial No.'] || ''}</td>
            <td style="font-weight: 500">${r['First Name'] || ''} ${r['Last Name'] || ''}</td>
            <td>${r['Application Date'] || ''}</td>
            <td>${mobVer}</td>
            <td>${emVer}</td>
            <td>${r['Programme Name'] || ''}</td>
            <td><span class="badge badge-info">${r['Academic Status'] || 'Student'}</span></td>
            <td>${offerBadge}</td>
            <td>${r['Ready for Offer Review Date'] || ''}</td>
            <td>${regBadge}</td>
            <td><span class="badge badge-stage ${stageBadgeClass}">${stg || 'Created'}</span></td>
        `;

        tr.addEventListener('click', () => openRecord(r['App Serial No.']));
        tableBody.appendChild(tr);
    });
}

// Open specific record details in drawer and populate form values
function openRecord(serialNo) {
    const record = state.records.find(r => r['App Serial No.'] === serialNo);
    if (!record) return;

    state.currentRecordId = serialNo;
    document.getElementById('drawerSerialBadge').textContent = serialNo;
    document.getElementById('drawerTitle').textContent = `${record['First Name'] || ''} ${record['Last Name'] || ''}`;

    // Populating inputs
    ADMISSIONS_DATA.columns.forEach(col => {
        const el = document.getElementById(`fld_${col.name}`);
        if (el) {
            el.value = record[col.name] || "";
        }
    });

    // Select default tab
    switchTab('stage-sec-applicant');
    
    // Evaluate verification alerts
    validateBusinessRules();
    
    // Show delete button
    document.getElementById('deleteRecordBtn').style.display = 'block';

    // Open centered modal
    document.getElementById('detailDrawer').classList.add('open');
    document.body.classList.add('modal-open');
}

// Initialize blank form for adding a new record
function openNewRecord() {
    state.currentRecordId = null;
    document.getElementById('drawerSerialBadge').textContent = "NEW";
    document.getElementById('drawerTitle').textContent = "New Applicant Profile";

    // Clear form inputs
    ADMISSIONS_DATA.columns.forEach(col => {
        const el = document.getElementById(`fld_${col.name}`);
        if (el) {
            el.value = "";
        }
    });

    // Populate default and audit fields
    const today = new Date().toISOString().split('T')[0];
    const nowTime = new Date().toISOString().replace('T', ' ').substring(0, 19);
    
    // Auto-generate App Serial Number based on largest numerical suffix
    const existingIds = state.records.map(r => r['App Serial No.']).filter(s => s && s.startsWith('AP26'));
    let maxSuffix = 100;
    existingIds.forEach(id => {
        const num = parseInt(id.replace('AP26', ''), 10);
        if (!isNaN(num) && num > maxSuffix) {
            maxSuffix = num;
        }
    });
    const newSerial = `AP26${maxSuffix + 1}`;
    
    document.getElementById('fld_App Serial No.').value = newSerial;
    document.getElementById('fld_Application Date').value = today;
    document.getElementById('fld_Current Pipeline Stage').value = "Applicant Data Incomplete";
    document.getElementById('fld_Current Pipeline Stage Date').value = today;
    document.getElementById('fld_Created').value = nowTime;
    document.getElementById('fld_Created By').value = "admissions.system@eub.edu.bh";
    document.getElementById('fld_Modified').value = nowTime;
    document.getElementById('fld_Modified By').value = "admissions.system@eub.edu.bh";
    document.getElementById('fld_Record Status').value = "Active";

    // Switch to first tab
    switchTab('stage-sec-applicant');
    
    // Clear validation banner
    document.getElementById('validationBanner').classList.add('hidden');
    
    // Hide delete button for new profiles
    document.getElementById('deleteRecordBtn').style.display = 'none';

    // Open centered modal
    document.getElementById('detailDrawer').classList.add('open');
    document.body.classList.add('modal-open');
}

// Switch tabs within the editing drawer
function switchTab(targetSectionId) {
    // Tab headers
    document.querySelectorAll('.tab-btn').forEach(btn => {
        if (btn.getAttribute('data-target') === targetSectionId) {
            btn.classList.add('active');
        } else {
            btn.classList.remove('active');
        }
    });

    // Content panels
    document.querySelectorAll('.form-stage-section').forEach(sec => {
        if (sec.id === targetSectionId) {
            sec.classList.add('active');
        } else {
            sec.classList.remove('active');
        }
    });
}

// Validate pipeline business logic constraints and alert users via UI banners
function validateBusinessRules() {
    const banner = document.getElementById('validationBanner');
    const title = document.getElementById('bannerTitle');
    const msg = document.getElementById('bannerMessage');
    
    const emailVer = document.getElementById('fld_Email Verification Status').value;
    const mobVer = document.getElementById('fld_Mobile Verification Status').value;
    const offerStatus = document.getElementById('fld_Offer Status').value;

    const regStatus = document.getElementById('fld_Registration Status').value;
    const studentId = document.getElementById('fld_Student ID').value;
    const feeStatus = document.getElementById('fld_Fees / Deposit Status').value;

    // Rule 1: Offer issue blocked by lack of email or mobile contact verification
    if ((offerStatus === 'Conditional' || offerStatus === 'Unconditional') && 
        (emailVer !== 'Verified' || mobVer !== 'Verified')) {
        banner.className = 'warning-banner';
        title.innerHTML = `<i class="fa-solid fa-ban"></i> OFFER BLOCKED`;
        msg.textContent = "Critical Block: Offer status cannot be Conditional or Unconditional because email and mobile verifications are not fully Verified.";
        return true;
    }

    // Rule 2: Registered student requirements
    if (regStatus === 'Registered' && (!studentId.trim() || feeStatus !== 'Paid' && feeStatus !== 'Waived')) {
        banner.className = 'warning-banner';
        title.innerHTML = `<i class="fa-solid fa-triangle-exclamation"></i> REGISTRATION WARNING`;
        msg.textContent = "Compliance Warning: Registered status requires fee/deposit clearance (Paid/Waived) and an issued Student ID.";
        return true;
    }

    // Fallback: Clear warning
    banner.className = 'warning-banner hidden';
    return false;
}

// Save or Create record after processing edits and timestamps
function saveRecordChanges() {
    const serialNo = document.getElementById('fld_App Serial No.').value;
    if (!serialNo) {
        alert("App Serial No. is required.");
        return;
    }

    // Check for hard blocks
    const emailVer = document.getElementById('fld_Email Verification Status').value;
    const mobVer = document.getElementById('fld_Mobile Verification Status').value;
    const offerStatus = document.getElementById('fld_Offer Status').value;

    if ((offerStatus === 'Conditional' || offerStatus === 'Unconditional') && 
        (emailVer !== 'Verified' || mobVer !== 'Verified')) {
        alert("Action Denied: You cannot issue an offer while applicant email or mobile details are unverified.");
        return;
    }

    // Read form values
    const formVals = {};
    ADMISSIONS_DATA.columns.forEach(col => {
        const el = document.getElementById(`fld_${col.name}`);
        if (el) {
            formVals[col.name] = el.value;
        }
    });

    const nowTime = new Date().toISOString().replace('T', ' ').substring(0, 19);
    const today = new Date().toISOString().split('T')[0];

    // Pipeline Stage Date Auto-update trigger
    if (state.currentRecordId) {
        // Edit Mode
        const idx = state.records.findIndex(r => r['App Serial No.'] === state.currentRecordId);
        if (idx !== -1) {
            const oldRecord = state.records[idx];
            
            // Check if Current Pipeline Stage changed
            if (oldRecord['Current Pipeline Stage'] !== formVals['Current Pipeline Stage']) {
                formVals['Current Pipeline Stage Date'] = today;
            }
            
            // Check if Registration Status changed to Registered
            if (oldRecord['Registration Status'] !== 'Registered' && formVals['Registration Status'] === 'Registered') {
                formVals['Registration Date'] = today;
            }

            formVals['Modified'] = nowTime;
            formVals['Modified By'] = "admissions.advisor@eub.edu.bh";
            formVals['Created'] = oldRecord['Created'];
            formVals['Created By'] = oldRecord['Created By'];

            state.records[idx] = formVals;
        }
    } else {
        // Create Mode
        formVals['Created'] = nowTime;
        formVals['Created By'] = "admissions.advisor@eub.edu.bh";
        formVals['Modified'] = nowTime;
        formVals['Modified By'] = "admissions.advisor@eub.edu.bh";
        
        state.records.push(formVals);
    }

    saveRecords();
    closeDrawer();
    updateUI();
}

// Delete Record confirmation
function deleteRecord() {
    if (!state.currentRecordId) return;

    if (confirm(`Are you sure you want to permanently delete record ${state.currentRecordId}? This action cannot be undone.`)) {
        state.records = state.records.filter(r => r['App Serial No.'] !== state.currentRecordId);
        saveRecords();
        closeDrawer();
        updateUI();
    }
}

// Close detailing drawer
function closeDrawer() {
    document.getElementById('detailDrawer').classList.remove('open');
    document.body.classList.remove('modal-open');
    state.currentRecordId = null;
}

// Wire events
function setupEventListeners() {
    // Text search
    document.getElementById('searchInput').addEventListener('input', () => {
        state.currentPage = 1;
        updateUI();
    });

    // Filters
    const filterSelectors = ['filterStage', 'filterAdvisor', 'filterOffer', 'filterVerification'];
    filterSelectors.forEach(id => {
        document.getElementById(id).addEventListener('change', () => {
            state.currentPage = 1;
            updateUI();
        });
    });

    // Clear Filters
    document.getElementById('clearFiltersBtn').addEventListener('click', () => {
        document.getElementById('searchInput').value = "";
        document.getElementById('filterStage').value = "";
        document.getElementById('filterAdvisor').value = "";
        document.getElementById('filterOffer').value = "";
        document.getElementById('filterVerification').value = "";
        state.currentPage = 1;
        updateUI();
    });

    // Pagination
    document.getElementById('prevPageBtn').addEventListener('click', () => {
        if (state.currentPage > 1) {
            state.currentPage--;
            renderTable();
        }
    });

    document.getElementById('nextPageBtn').addEventListener('click', () => {
        const totalPages = Math.ceil(state.filteredRecords.length / state.pageSize);
        if (state.currentPage < totalPages) {
            state.currentPage++;
            renderTable();
        }
    });

    // Sorting headers
    document.querySelectorAll('.field-headers-row th').forEach(th => {
        th.addEventListener('click', () => {
            const colName = th.getAttribute('data-sort');
            if (!colName) return;

            if (state.sortColumn === colName) {
                state.sortOrder = state.sortOrder === 'asc' ? 'desc' : 'asc';
            } else {
                state.sortColumn = colName;
                state.sortOrder = 'asc';
            }

            // Update arrow classes in headers
            document.querySelectorAll('.field-headers-row th i').forEach(icon => {
                icon.className = 'fa-solid fa-sort';
            });
            const subIcon = th.querySelector('i');
            subIcon.className = state.sortOrder === 'asc' ? 'fa-solid fa-sort-up' : 'fa-solid fa-sort-down';

            state.currentPage = 1;
            updateUI();
        });
    });

    // Theme Toggle
    document.getElementById('themeToggleBtn').addEventListener('click', toggleTheme);

    // Export JSON
    document.getElementById('exportJsonBtn').addEventListener('click', exportDatabaseToJson);

    // Reset Data
    document.getElementById('resetDataBtn').addEventListener('click', () => {
        if (confirm("Are you sure you want to reset the admissions database to its original Excel state? All your custom modifications will be lost.")) {
            localStorage.removeItem(LOCAL_STORAGE_KEY);
            loadRecords();
            updateUI();
        }
    });

    // New record profile
    document.getElementById('addApplicantBtn').addEventListener('click', openNewRecord);

    // Form section toggling inside drawer
    document.querySelectorAll('.tab-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            const target = btn.getAttribute('data-target');
            switchTab(target);
        });
    });

    // Inputs value monitoring for warning banner
    const monitoredFields = [
        'fld_Email Verification Status', 
        'fld_Mobile Verification Status', 
        'fld_Offer Status', 
        'fld_Registration Status', 
        'fld_Student ID', 
        'fld_Fees / Deposit Status'
    ];
    monitoredFields.forEach(id => {
        const el = document.getElementById(id);
        if (el) {
            el.addEventListener('change', validateBusinessRules);
        }
    });

    // Drawer closing bindings
    document.getElementById('closeDrawerBtn').addEventListener('click', closeDrawer);
    document.getElementById('cancelDrawerBtn').addEventListener('click', closeDrawer);
    document.getElementById('drawerBackdrop').addEventListener('click', closeDrawer);

    // Form save action
    document.getElementById('saveRecordBtn').addEventListener('click', saveRecordChanges);
    
    // Delete action
    document.getElementById('deleteRecordBtn').addEventListener('click', deleteRecord);
}

// Download file of records database as JSON
function exportDatabaseToJson() {
    const dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify(state.records, null, 2));
    const dlAnchorElem = document.createElement('a');
    dlAnchorElem.setAttribute("href", dataStr);
    
    const today = new Date().toISOString().split('T')[0];
    dlAnchorElem.setAttribute("download", `eub_admissions_backup_${today}.json`);
    dlAnchorElem.click();
}
