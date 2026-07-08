// Admissions Database Frontend Application Logic
//
// Audit flow summary:
// 1. Microsoft MSAL authenticates users before the dashboard is shown.
// 2. Records load from browser localStorage, or from ADMISSIONS_DATA as seed data.
// 3. Dashboard controls update in-memory state, then re-render KPIs, charts and rows.
// 4. Record edits are validated, timestamped, saved back to localStorage and re-rendered.

// Constants
const LOCAL_STORAGE_KEY = 'eub_admissions_records';
const SEED_VARIANT_STORAGE_KEY = 'eub_admissions_seed_variant';
const THEME_STORAGE_KEY = 'eub_admissions_theme';
const AUTH_STORAGE_KEY = 'eub_admissions_m365_account';
const LOGIN_SCOPES = ['openid', 'profile', 'email'];
const THEME_CLASSES = ['light-theme', 'dark-theme'];
const CURRENT_SEED_VARIANT = 'v3';

const NAME_POOLS = {
    Female: ['Alya', 'Nora', 'Mariam', 'Sara', 'Zaina', 'Hala', 'Reem', 'Dana', 'Leen', 'Yara', 'Salma', 'Tala', 'Rana', 'Farah', 'Lina', 'Jana', 'Raghad', 'Maha'],
    Male: ['Omar', 'Yousef', 'Zayd', 'Hasan', 'Kareem', 'Adel', 'Tariq', 'Rayyan', 'Faris', 'Hassan', 'Nasser', 'Jad', 'Samer', 'Bilal', 'Rami', 'Ibrahim', 'Saeed', 'Mazen'],
    Neutral: ['Noor', 'Jordan', 'Ari', 'Mika', 'Sam', 'Taj', 'Robin', 'Nour', 'Aman', 'Sky']
};

const LAST_NAMES = ['Al Khalifa', 'Rahman', 'Al Mazrouei', 'Haddad', 'Qureshi', 'Al Mutairi', 'Farouq', 'Darwish', 'Najjar', 'Al Sabahi', 'Hamdan', 'Sharif', 'Mansoor', 'Abdulla', 'Kanaan', 'Bashir', 'Al Nuaimi', 'Saad', 'Hussain', 'Rizvi'];
const SCHOOL_NAMES = ['Al Hekma College', 'Beacon International School', 'Crescent Sixth Form', 'New Dawn Academy', 'Bahrain Scholars School', 'Al Manar Secondary School', 'Peninsula Learning Centre', 'West Bay International Academy', 'Knowledge Bridge School', 'Al Rawabi Girls School', 'Future Path Institute', 'Ibn Rushd Secondary School', 'Capital Community School', 'Al Salam Private School', 'North Gate Academy', 'Horizon Science College'];
const GRADE_SUMMARIES = ['Overall 68% with stronger humanities scores', 'Overall 74% with a B in mathematics', 'Predicted ABB at A Level', 'IB predicted 29 points with HL English 5', 'American diploma GPA 3.1 / 4.0', 'Overall 81% with distinction in business', 'BTEC profile DMM with solid coursework', 'High school average 77% with improved final term', 'Overall 84% and strong interview notes', 'Predicted 72% equivalent with good attendance', 'Foundation average 3.3 GPA equivalent', 'Overall 69% with resit pending in one subject'];
const NOTE_VARIATIONS = ['Needs a second follow-up after the initial enquiry call.', 'Applicant asked for a family cost breakdown before progressing.', 'Interested in campus life and student support services.', 'Requested evening outreach because daytime calls are missed.', 'Counsellor noted strong motivation but slower document turnaround.', 'Family prefers updates by email before any phone follow-up.', 'Applicant is comparing offers from two regional universities.', 'Guardian requested a clearer explanation of progression routes.', 'Good engagement so far; waiting on one remaining attachment.', 'Student asked for examples of internship and placement outcomes.', 'Prefers short WhatsApp updates rather than long call summaries.', 'Needs help understanding the difference between entry pathways.'];
const RELATIONSHIP_NAMES = {
    Father: ['Khalid', 'Omar', 'Hussain', 'Tariq', 'Nabil', 'Sami', 'Faisal', 'Majid'],
    Mother: ['Aisha', 'Layla', 'Huda', 'Mona', 'Rania', 'Samira', 'Nadia', 'Iman'],
    Guardian: ['Amal', 'Rashid', 'Muna', 'Adnan', 'Salwa', 'Bassam', 'Hanan', 'Wael']
};

const STAGE_WEIGHTS = {
    'Applicant Data Incomplete': 1.3,
    'Contact Verification': 1.25,
    'Contact Verification Dormant': 0.75,
    'Offer Readiness': 1.05,
    'Offer In Process': 1.1,
    'Conditional Offer Issued': 0.95,
    'Unconditional Offer Issued': 0.9,
    'Conversion': 0.9,
    'Registration In Process': 0.85,
    Registered: 0.8,
    Withdrawn: 0.45,
    Rejected: 0.4,
    'Duplicate / Invalid': 0.35
};

const READ_ONLY_FIELDS = [
    'App Serial No.',
    'Created',
    'Created By',
    'Modified',
    'Modified By',
    'Current Pipeline Stage Date',
    'Registration Handover Date',
    'Registration Date',
    'Parent / Guardian Consent Date'
];

const FILTER_CONTROL_IDS = ['searchInput', 'filterStage', 'filterAdvisor', 'filterOffer', 'filterVerification'];
const SYSTEM_USER_EMAIL = 'admissions.system@eub.edu.bh';
const ADVISOR_USER_EMAIL = 'admissions.advisor@eub.edu.bh';

const AUTH_CONFIG = window.AUTH_CONFIG || {
    clientId: '',
    tenantId: 'common',
    redirectUri: new URL('index.html', window.location.href).toString(),
    postLogoutRedirectUri: new URL('index.html', window.location.href).toString(),
    cacheLocation: 'localStorage',
    appPageUrl: new URL('admissions_db.html', window.location.href).toString(),
    loginPageUrl: new URL('index.html', window.location.href).toString()
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
    theme: 'light-theme',
    dashboardFilter: null
};

function byId(id) {
    return document.getElementById(id);
}

function setText(id, value) {
    const el = byId(id);
    if (el) {
        el.textContent = value;
    }
}

function setValue(id, value) {
    const el = byId(id);
    if (el) {
        el.value = value;
    }
}

function currentDate() {
    return new Date().toISOString().split('T')[0];
}

function currentTimestamp() {
    return new Date().toISOString().replace('T', ' ').substring(0, 19);
}

function createOption(value, label = value) {
    const option = document.createElement('option');
    option.value = value;
    option.textContent = label;
    return option;
}

function addSelectOptions(selectId, values) {
    const select = byId(selectId);
    if (!select) {
        return;
    }

    values.filter(Boolean).forEach(value => {
        select.appendChild(createOption(value));
    });
}

function resetControls(controlIds = FILTER_CONTROL_IDS) {
    controlIds.forEach(id => setValue(id, ''));
}

function isActivationKey(event) {
    return event.key === 'Enter' || event.key === ' ';
}

function makeClickableFilterTarget(el, filter, title) {
    if (!el) {
        return;
    }

    el.classList.add('dashboard-clickable');
    el.setAttribute('role', 'button');
    el.setAttribute('tabindex', '0');
    el.setAttribute('title', title);

    el.addEventListener('click', () => setDashboardFilter(filter));
    el.addEventListener('keydown', (event) => {
        if (isActivationKey(event)) {
            event.preventDefault();
            setDashboardFilter(filter);
        }
    });
}

function isFullyVerified(record) {
    return record['Mobile Verification Status'] === 'Verified' && record['Email Verification Status'] === 'Verified';
}

function isOfferIssued(record) {
    return record['Offer Status'] === 'Conditional' || record['Offer Status'] === 'Unconditional';
}

function isRegistrationComplete(record) {
    return record['Registration Status'] === 'Registered';
}

function isOfferBlocked(emailStatus, mobileStatus, offerStatus) {
    return (offerStatus === 'Conditional' || offerStatus === 'Unconditional') &&
        (emailStatus !== 'Verified' || mobileStatus !== 'Verified');
}

function readFormValues() {
    return ADMISSIONS_DATA.columns.reduce((values, col) => {
        const el = byId(`fld_${col.name}`);
        if (el) {
            values[col.name] = el.value;
        }
        return values;
    }, {});
}

function fillFormValues(record) {
    ADMISSIONS_DATA.columns.forEach(col => {
        setValue(`fld_${col.name}`, record[col.name] || '');
    });
}

function setDrawerOpen(isOpen) {
    byId('detailDrawer')?.classList.toggle('open', isOpen);
    document.body.classList.toggle('modal-open', isOpen);
}

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
    initializePage();
});

function initializePage() {
    loadTheme();

    if (isLoginPage()) {
        initializeLoginPage();
        return;
    }

    initializeAdmissionsPage();
}

function isLoginPage() {
    return document.body?.dataset.page === 'login' || !!byId('microsoftLoginBtn');
}

function getAppPageUrl() {
    return resolvePageUrl(AUTH_CONFIG.appPageUrl, 'admissions_db.html');
}

function getLoginPageUrl() {
    return resolvePageUrl(AUTH_CONFIG.loginPageUrl, 'index.html');
}

function resolvePageUrl(configuredUrl, fallbackPage) {
    const fallbackUrl = new URL(fallbackPage, window.location.href).toString();
    if (!configuredUrl) {
        return fallbackUrl;
    }

    try {
        const parsed = new URL(configuredUrl, window.location.href);

        // Guard against accidentally using github.com repo links instead of github.io site links.
        if (parsed.hostname === 'github.com') {
            return fallbackUrl;
        }

        return parsed.toString();
    } catch (error) {
        console.warn('Invalid configured page URL, using fallback.', error);
        return fallbackUrl;
    }
}

function initializeLoginPage() {
    bindAuthButtons();

    if (!window.msal || !AUTH_CONFIG.clientId) {
        showAuthSetupMessage();
        return;
    }

    msalInstance = new msal.PublicClientApplication(msalConfig);
    handleLoginPageRedirect();
}

async function handleLoginPageRedirect() {
    try {
        const response = await msalInstance.handleRedirectPromise();
        if (response && response.account) {
            msalInstance.setActiveAccount(response.account);
            persistAccount(response.account);
            window.location.replace(getAppPageUrl());
            return;
        }

        const activeAccount = msalInstance.getActiveAccount() || msalInstance.getAllAccounts()[0];
        if (activeAccount) {
            msalInstance.setActiveAccount(activeAccount);
            persistAccount(activeAccount);
            window.location.replace(getAppPageUrl());
            return;
        }

        showAuthScreen();
    } catch (error) {
        console.error('Microsoft sign-in initialization failed', error);
        showAuthError('Microsoft sign-in could not be initialized. Check the tenant and client ID configuration.');
    }
}

async function initializeAdmissionsPage() {
    if (!window.msal || !AUTH_CONFIG.clientId) {
        window.location.replace(getLoginPageUrl());
        return;
    }

    msalInstance = new msal.PublicClientApplication(msalConfig);

    try {
        const response = await msalInstance.handleRedirectPromise();
        if (response && response.account) {
            msalInstance.setActiveAccount(response.account);
        }
    } catch (error) {
        console.error('Microsoft redirect handling failed on app page', error);
    }

    const activeAccount = msalInstance.getActiveAccount() || msalInstance.getAllAccounts()[0];
    if (!activeAccount) {
        window.location.replace(getLoginPageUrl());
        return;
    }

    msalInstance.setActiveAccount(activeAccount);
    persistAccount(activeAccount);
    document.body.classList.remove('auth-only');

    loadRecords();
    initFilterDropdowns();
    renderFormFields();
    bindAuthButtons();
    setupEventListeners();
    updateUI();
}

function bindAuthButtons() {
    const loginButton = byId('microsoftLoginBtn');
    if (loginButton) {
        loginButton.addEventListener('click', signInWithMicrosoft);
    }

    const signOutButton = byId('signOutBtn');
    if (signOutButton) {
        signOutButton.addEventListener('click', signOutWithMicrosoft);
    }
}

function showAuthScreen() {
    const authScreen = byId('authScreen');
    if (authScreen) {
        authScreen.classList.add('visible');
    }

    const appShell = byId('appShell');
    if (appShell) {
        appShell.classList.add('app-hidden');
        appShell.setAttribute('aria-hidden', 'true');
    }

    document.body.classList.add('auth-only');
}

function showAuthSetupMessage() {
    showAuthScreen();
    setText('authStatusText', 'Microsoft sign-in is not configured yet. Add your Entra app settings in auth-config.js.');

    const loginButton = byId('microsoftLoginBtn');
    if (loginButton) {
        loginButton.disabled = true;
    }
}

function showAuthError(message) {
    showAuthScreen();
    setText('authStatusText', message);
}

function persistAccount(account) {
    setText('authStatusText', `Signed in as ${account.name || account.username || 'Microsoft user'}.`);

    localStorage.setItem(AUTH_STORAGE_KEY, JSON.stringify({
        name: account.name || '',
        username: account.username || '',
        homeAccountId: account.homeAccountId || ''
    }));
}

async function signInWithMicrosoft() {
    if (!msalInstance) return;

    setText('authStatusText', 'Opening Microsoft sign-in...');

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
            account: msalInstance.getActiveAccount() || msalInstance.getAllAccounts()[0],
            postLogoutRedirectUri: getLoginPageUrl()
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
    applyThemeClass(savedTheme);
    syncThemeSwitch(savedTheme);
    if (!localStorage.getItem(THEME_STORAGE_KEY)) {
        localStorage.setItem(THEME_STORAGE_KEY, savedTheme);
    }
}

function applyThemeClass(theme) {
    document.body.classList.remove(...THEME_CLASSES);
    document.body.classList.add(theme);
}

// Toggle Theme between dark and light
function toggleTheme() {
    if (state.theme === 'dark-theme') {
        state.theme = 'light-theme';
    } else {
        state.theme = 'dark-theme';
    }
    applyThemeClass(state.theme);
    syncThemeSwitch(state.theme);
    localStorage.setItem(THEME_STORAGE_KEY, state.theme);
}

function syncThemeSwitch(theme) {
    const switchEl = byId('themeToggleSwitch');
    if (switchEl) {
        switchEl.checked = theme === 'dark-theme';
    }
}

function getNamePool(gender) {
    return NAME_POOLS[gender] || NAME_POOLS.Neutral;
}

function slugifyNamePart(value) {
    return String(value || '')
        .toLowerCase()
        .replace(/[^a-z0-9]+/g, '.')
        .replace(/^\.+|\.+$/g, '');
}

function createPhoneNumber(prefix, index, offset) {
    const suffix = String(((index + 1) * 137 + offset * 911) % 1000000).padStart(6, '0');
    return `${prefix}${suffix}`;
}

function looksLikeLegacySeed(records) {
    if (!Array.isArray(records) || records.length === 0) {
        return false;
    }

    const sample = records.slice(0, Math.min(24, records.length));
    const uniqueStudentMobiles = new Set(sample.map(record => record['Mobile No.1'])).size;
    const uniqueGuardianMobiles = new Set(sample.map(record => record['Parent / Guardian Mobile'])).size;
    const hasLegacyEmailDomain = sample.every(record => String(record['Email Address'] || '').includes('@fictional-applicants.edu'));

    return hasLegacyEmailDomain && uniqueStudentMobiles <= 2 && uniqueGuardianMobiles <= 2;
}

function diversifySeedRecord(record, index) {
    const namePool = getNamePool(record['Gender']);
    const firstName = namePool[index % namePool.length];
    const lastName = LAST_NAMES[Math.floor(index / namePool.length) % LAST_NAMES.length];
    const relationship = record['Parent / Guardian Relationship'];
    const guardianPool = RELATIONSHIP_NAMES[relationship] || RELATIONSHIP_NAMES.Guardian;
    const guardianFirstName = guardianPool[(index * 3) % guardianPool.length];
    const schoolName = SCHOOL_NAMES[(index * 5) % SCHOOL_NAMES.length];
    const grades = GRADE_SUMMARIES[(index * 7) % GRADE_SUMMARIES.length];
    const note = NOTE_VARIATIONS[(index * 11) % NOTE_VARIATIONS.length];
    const applicantSlug = `${slugifyNamePart(firstName)}.${slugifyNamePart(lastName)}`;
    const guardianSlug = `${slugifyNamePart(guardianFirstName)}.${slugifyNamePart(lastName)}`;
    const serialDigits = String(record['App Serial No.'] || '').replace(/\D/g, '').slice(-4) || String(index + 1).padStart(4, '0');

    return {
        ...record,
        'First Name': firstName,
        'Last Name': lastName,
        'Mobile No.1': createPhoneNumber('9733', index, 17),
        'Email Address': `${applicantSlug}${serialDigits}@applicants.example`,
        'Parent / Guardian Name': `${guardianFirstName} ${lastName}`,
        'Parent / Guardian Mobile': createPhoneNumber('9736', index, 29),
        'Parent / Guardian Email': `${guardianSlug}${serialDigits}@familymail.example`,
        'School Name or Awarding Body': schoolName,
        'Existing / Predicted Grades': grades,
        'Exception Notes': note
    };
}

function getStageWeight(stage) {
    return STAGE_WEIGHTS[stage] || 0.65;
}

function buildRandomStageAssignments(recordCount, stages) {
    if (!recordCount || !Array.isArray(stages) || stages.length === 0) {
        return [];
    }

    const counts = new Array(stages.length).fill(0);
    let remaining = recordCount;

    if (recordCount >= stages.length) {
        for (let i = 0; i < stages.length; i++) {
            counts[i] = 1;
            remaining--;
        }
    }

    const weighted = stages.map(stage => {
        const variance = 0.75 + Math.random() * 0.85;
        return Math.max(0.1, getStageWeight(stage) * variance);
    });
    const totalWeight = weighted.reduce((sum, value) => sum + value, 0);

    const fractions = weighted.map((weight, idx) => {
        const exact = totalWeight > 0 ? (weight / totalWeight) * remaining : 0;
        const chunk = Math.floor(exact);
        counts[idx] += chunk;
        return exact - chunk;
    });

    let assigned = counts.reduce((sum, value) => sum + value, 0);
    while (assigned < recordCount) {
        const randomBump = Math.random() * 0.2;
        let target = 0;
        let bestScore = -1;

        fractions.forEach((fraction, idx) => {
            const score = fraction + (idx === target ? 0 : randomBump * Math.random());
            if (score > bestScore) {
                bestScore = score;
                target = idx;
            }
        });

        counts[target] += 1;
        assigned += 1;
    }

    const assignments = [];
    counts.forEach((count, idx) => {
        for (let i = 0; i < count; i++) {
            assignments.push(stages[idx]);
        }
    });

    for (let i = assignments.length - 1; i > 0; i--) {
        const j = Math.floor(Math.random() * (i + 1));
        [assignments[i], assignments[j]] = [assignments[j], assignments[i]];
    }

    return assignments;
}

function applyStageDataProfile(record, stage, index) {
    const isLateStage = stage === 'Offer In Process' || stage === 'Conditional Offer Issued' || stage === 'Unconditional Offer Issued' || stage === 'Conversion' || stage === 'Registration In Process' || stage === 'Registered';
    const hasFullVerification = stage === 'Registered' || stage === 'Registration In Process' || stage === 'Unconditional Offer Issued' || (isLateStage && Math.random() > 0.12);
    const isOfferConditional = stage === 'Conditional Offer Issued';
    const isOfferUnconditional = stage === 'Unconditional Offer Issued' || stage === 'Registered';

    const mobileStatus = hasFullVerification ? 'Verified' : (stage === 'Applicant Data Incomplete' ? 'Attempted' : (Math.random() > 0.45 ? 'Verified' : 'Attempted'));
    const emailStatus = hasFullVerification ? 'Verified' : (stage === 'Applicant Data Incomplete' ? 'Verification Email Sent' : (Math.random() > 0.4 ? 'Verified' : 'Verification Email Sent'));

    let offerStatus = 'In Process';
    if (isOfferConditional) offerStatus = 'Conditional';
    if (isOfferUnconditional) offerStatus = 'Unconditional';
    if (stage === 'Rejected') offerStatus = 'Rejected';

    let registrationStatus = 'Not Started';
    if (stage === 'Unconditional Offer Issued') registrationStatus = 'Ready for Registration';
    if (stage === 'Registration In Process') registrationStatus = 'Registration In Process';
    if (stage === 'Registered') registrationStatus = 'Registered';

    const verificationBlocked = mobileStatus !== 'Verified' || emailStatus !== 'Verified';

    return {
        ...record,
        'Current Pipeline Stage': stage,
        'Mobile Verification Status': mobileStatus,
        'Email Verification Status': emailStatus,
        'Contact Verification Gate': verificationBlocked ? 'Not Passed' : 'Passed',
        'Verification Blocker': verificationBlocked ? 'Missing Email or Mobile Verification' : '',
        'Offer Status': offerStatus,
        'Registration Status': registrationStatus,
        'Student ID': registrationStatus === 'Registered' ? `STU${String(260000 + index + 1)}` : ''
    };
}

function buildDiversifiedSeed(records) {
    const pipelineStages = (ADMISSIONS_DATA.dropdowns['Current Pipeline Stage'] || []).filter(Boolean);
    const stageAssignments = buildRandomStageAssignments(records.length, pipelineStages);

    return records.map((record, index) => {
        const diversified = diversifySeedRecord(record, index);
        const stage = stageAssignments[index] || diversified['Current Pipeline Stage'] || 'Applicant Data Incomplete';
        return applyStageDataProfile(diversified, stage, index);
    });
}

// Load records from local storage or fallback to spreadsheet data
function loadRecords() {
    const stored = localStorage.getItem(LOCAL_STORAGE_KEY);
    const storedSeedVariant = localStorage.getItem(SEED_VARIANT_STORAGE_KEY);
    let shouldPersist = false;

    if (stored) {
        try {
            state.records = JSON.parse(stored);
        } catch (e) {
            console.error("Error parsing stored records, resetting to seed data", e);
            state.records = buildDiversifiedSeed(ADMISSIONS_DATA.records);
            shouldPersist = true;
        }
    } else {
        state.records = buildDiversifiedSeed(ADMISSIONS_DATA.records);
        shouldPersist = true;
    }

    if (!Array.isArray(state.records)) {
        state.records = buildDiversifiedSeed(ADMISSIONS_DATA.records);
        shouldPersist = true;
    }

    if (storedSeedVariant !== CURRENT_SEED_VARIANT && looksLikeLegacySeed(state.records)) {
        state.records = buildDiversifiedSeed(state.records);
        shouldPersist = true;
    }

    if (shouldPersist) {
        localStorage.setItem(SEED_VARIANT_STORAGE_KEY, CURRENT_SEED_VARIANT);
        saveRecords();
    }

    state.filteredRecords = [...state.records];
}

// Save current records state to local storage
function saveRecords() {
    localStorage.setItem(LOCAL_STORAGE_KEY, JSON.stringify(state.records));
}

// Initialize dynamic filters based on sheet contents
function initFilterDropdowns() {
    addSelectOptions('filterStage', ADMISSIONS_DATA.dropdowns['Current Pipeline Stage'] || []);

    // Advisor filter is derived from records because ownership lives in the source data.
    const advisors = [...new Set(state.records.map(r => r['Advisor / Admissions Owner']).filter(Boolean))].sort();
    addSelectOptions('filterAdvisor', advisors);
    addSelectOptions('filterOffer', ADMISSIONS_DATA.dropdowns['Offer Status'] || []);
}

// Generate the 88 edit fields dynamically in their respective tabs inside the drawer
function renderFormFields() {
    ADMISSIONS_DATA.columns.forEach(col => {
        const sectionId = SECTION_MAPPING[col.group];
        if (!sectionId) return;

        const container = byId(sectionId);
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
            
            inputEl.appendChild(createOption('', '-- Select --'));

            const options = ADMISSIONS_DATA.dropdowns[col.name] || [];
            options.forEach(optVal => {
                inputEl.appendChild(createOption(optVal));
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

        // Audit/system fields are populated by code and locked to protect record history.
        if (READ_ONLY_FIELDS.includes(col.name)) {
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
    updateDashboardActiveStates();
    calculateKPIs();
    renderCharts();
    filterAndSortRecords();
    renderTable();
}

function setDashboardFilter(filter) {
    const isSameFilter = state.dashboardFilter && filter &&
        state.dashboardFilter.type === filter.type &&
        state.dashboardFilter.value === filter.value;

    if (!isSameFilter) {
        resetControls();
    }

    state.dashboardFilter = isSameFilter ? null : filter;
    state.currentPage = 1;
    updateUI();
}

function matchesDashboardFilter(record) {
    if (!state.dashboardFilter) {
        return true;
    }

    const filterChecks = {
        'kpi-total': () => true,
        'kpi-verified': isFullyVerified,
        'kpi-offers': isOfferIssued,
        'kpi-registered': isRegistrationComplete,
        stage: item => item['Current Pipeline Stage'] === state.dashboardFilter.value,
        advisor: item => item['Advisor / Admissions Owner'] === state.dashboardFilter.value
    };

    return (filterChecks[state.dashboardFilter.type] || (() => true))(record);
}

function updateDashboardActiveStates() {
    document.querySelectorAll('.dashboard-clickable').forEach(el => {
        el.classList.remove('dashboard-active');
    });

    if (!state.dashboardFilter) {
        return;
    }

    if (state.dashboardFilter.type.startsWith('kpi-')) {
        const card = document.querySelector(`.kpi-card[data-dashboard-key="${state.dashboardFilter.type}"]`);
        if (card) {
            card.classList.add('dashboard-active');
        }
    }
}

function bindDashboardClickTargets() {
    const cardBindings = [
        { selector: '.kpi-total', key: 'kpi-total' },
        { selector: '.kpi-verified', key: 'kpi-verified' },
        { selector: '.kpi-offers', key: 'kpi-offers' },
        { selector: '.kpi-registered', key: 'kpi-registered' }
    ];

    cardBindings.forEach(binding => {
        const card = document.querySelector(binding.selector);
        if (!card) return;

        card.dataset.dashboardKey = binding.key;
        makeClickableFilterTarget(card, { type: binding.key }, 'Click to filter records in the table');
    });
}

// Compute dashboard statistics dynamically
function calculateKPIs() {
    const totalCount = state.records.length;
    setText('kpiTotalCount', totalCount);

    // Contact Verification Gate requires both verified email and verified mobile.
    const verifiedCount = state.records.filter(isFullyVerified).length;
    setText('kpiVerifiedCount', verifiedCount);
    const verifiedPercent = totalCount > 0 ? Math.round((verifiedCount / totalCount) * 100) : 0;
    setText('kpiVerifiedPercent', `${verifiedPercent}% Fully Verified`);

    const offersCount = state.records.filter(isOfferIssued).length;
    setText('kpiOffersCount', offersCount);
    const offersPercent = totalCount > 0 ? Math.round((offersCount / totalCount) * 100) : 0;
    setText('kpiOffersPercent', `${offersPercent}% of applicants`);

    const registeredCount = state.records.filter(isRegistrationComplete).length;
    setText('kpiRegisteredCount', registeredCount);
    const registeredPercent = totalCount > 0 ? Math.round((registeredCount / totalCount) * 100) : 0;
    setText('kpiRegistrationPercent', `${registeredPercent}% Conversion Rate`);
}

// Render dynamic visual SVG charts
function renderCharts() {
    // 1. Pipeline Stages volume chart
    const stageContainer = byId('pipelineStageChart');
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
        col.dataset.stage = stage;
        makeClickableFilterTarget(col, { type: 'stage', value: stage }, `Filter table by ${stage}`);

        if (state.dashboardFilter?.type === 'stage' && state.dashboardFilter.value === stage) {
            col.classList.add('dashboard-active');
        }
        
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
    const advisorContainer = byId('advisorWorkloadChart');
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
        row.className = 'advisor-chart-row';
        row.style.width = '100%';
        row.style.marginBottom = '0.5rem';
        makeClickableFilterTarget(row, { type: 'advisor', value: adv }, `Filter table by advisor ${adv}`);

        if (state.dashboardFilter?.type === 'advisor' && state.dashboardFilter.value === adv) {
            row.classList.add('dashboard-active');
        }
        
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
    const searchVal = (byId('searchInput')?.value || '').toLowerCase().trim();
    const stageVal = byId('filterStage')?.value || '';
    const advisorVal = byId('filterAdvisor')?.value || '';
    const offerVal = byId('filterOffer')?.value || '';
    const verificationVal = byId('filterVerification')?.value || '';

    state.filteredRecords = state.records.filter(r => {
        // Search is intentionally broad so staff can locate records without knowing the exact field.
        const name = `${r['First Name'] || ''} ${r['Last Name'] || ''}`.toLowerCase();
        const searchableValues = [
            name,
            r['App Serial No.'],
            r['Nationality'],
            r['Email Address'],
            r['Advisor / Admissions Owner']
        ].map(value => String(value || '').toLowerCase());

        const searchMatches = !searchVal || searchableValues.some(value => value.includes(searchVal));
        const stageMatches = !stageVal || r['Current Pipeline Stage'] === stageVal;
        const advisorMatches = !advisorVal || r['Advisor / Admissions Owner'] === advisorVal;
        const offerMatches = !offerVal || r['Offer Status'] === offerVal;
        
        // Verification filter mirrors the compliance gate used before issuing offers.
        let verificationMatches = true;
        if (verificationVal === 'Blocked') {
            verificationMatches = r['Verification Blocker'] || 
                (r['Mobile Verification Status'] !== 'Verified' || r['Email Verification Status'] !== 'Verified');
        } else if (verificationVal === 'Unblocked') {
            verificationMatches = !r['Verification Blocker'] && 
                r['Mobile Verification Status'] === 'Verified' && 
                r['Email Verification Status'] === 'Verified';
        }

        const dashboardMatches = matchesDashboardFilter(r);

        return searchMatches && stageMatches && advisorMatches && offerMatches && verificationMatches && dashboardMatches;
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
    const tableBody = byId('tableBody');
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
    setText('resultCountText', `Showing ${totalRecords > 0 ? startIndex + 1 : 0} to ${endIndex} of ${totalRecords} records`);
    setText('pageNumberDisplay', `Page ${state.currentPage} of ${totalPages}`);
    byId('prevPageBtn').disabled = state.currentPage === 1;
    byId('nextPageBtn').disabled = state.currentPage === totalPages;

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
    setText('drawerSerialBadge', serialNo);
    setText('drawerTitle', `${record['First Name'] || ''} ${record['Last Name'] || ''}`);

    fillFormValues(record);

    // Select default tab
    switchTab('stage-sec-applicant');
    
    // Evaluate verification alerts
    validateBusinessRules();
    
    // Show delete button
    byId('deleteRecordBtn').style.display = 'block';

    setDrawerOpen(true);
}

// Initialize blank form for adding a new record
function openNewRecord() {
    state.currentRecordId = null;
    setText('drawerSerialBadge', 'NEW');
    setText('drawerTitle', 'New Applicant Profile');

    fillFormValues({});

    // Populate default and audit fields
    const today = currentDate();
    const nowTime = currentTimestamp();
    
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
    
    setValue('fld_App Serial No.', newSerial);
    setValue('fld_Application Date', today);
    setValue('fld_Current Pipeline Stage', 'Applicant Data Incomplete');
    setValue('fld_Current Pipeline Stage Date', today);
    setValue('fld_Created', nowTime);
    setValue('fld_Created By', SYSTEM_USER_EMAIL);
    setValue('fld_Modified', nowTime);
    setValue('fld_Modified By', SYSTEM_USER_EMAIL);
    setValue('fld_Record Status', 'Active');

    // Switch to first tab
    switchTab('stage-sec-applicant');
    
    // Clear validation banner
    byId('validationBanner').classList.add('hidden');
    
    // Hide delete button for new profiles
    byId('deleteRecordBtn').style.display = 'none';

    setDrawerOpen(true);
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
    const banner = byId('validationBanner');
    const title = byId('bannerTitle');
    const msg = byId('bannerMessage');
    
    const emailVer = byId('fld_Email Verification Status').value;
    const mobVer = byId('fld_Mobile Verification Status').value;
    const offerStatus = byId('fld_Offer Status').value;

    const regStatus = byId('fld_Registration Status').value;
    const studentId = byId('fld_Student ID').value;
    const feeStatus = byId('fld_Fees / Deposit Status').value;

    // Rule 1: Offer issue is blocked until both contact channels are verified.
    if (isOfferBlocked(emailVer, mobVer, offerStatus)) {
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
    const serialNo = byId('fld_App Serial No.').value;
    if (!serialNo) {
        alert("App Serial No. is required.");
        return;
    }

    const emailVer = byId('fld_Email Verification Status').value;
    const mobVer = byId('fld_Mobile Verification Status').value;
    const offerStatus = byId('fld_Offer Status').value;

    // This hard stop enforces the same rule as the visible warning banner.
    if (isOfferBlocked(emailVer, mobVer, offerStatus)) {
        alert("Action Denied: You cannot issue an offer while applicant email or mobile details are unverified.");
        return;
    }

    const formVals = readFormValues();

    const nowTime = currentTimestamp();
    const today = currentDate();

    // Audit timestamps are refreshed only after validation passes.
    if (state.currentRecordId) {
        const idx = state.records.findIndex(r => r['App Serial No.'] === state.currentRecordId);
        if (idx !== -1) {
            const oldRecord = state.records[idx];
            
            if (oldRecord['Current Pipeline Stage'] !== formVals['Current Pipeline Stage']) {
                formVals['Current Pipeline Stage Date'] = today;
            }
            
            // Check if Registration Status changed to Registered
            if (oldRecord['Registration Status'] !== 'Registered' && formVals['Registration Status'] === 'Registered') {
                formVals['Registration Date'] = today;
            }

            formVals['Modified'] = nowTime;
            formVals['Modified By'] = ADVISOR_USER_EMAIL;
            formVals['Created'] = oldRecord['Created'];
            formVals['Created By'] = oldRecord['Created By'];

            state.records[idx] = formVals;
        }
    } else {
        formVals['Created'] = nowTime;
        formVals['Created By'] = ADVISOR_USER_EMAIL;
        formVals['Modified'] = nowTime;
        formVals['Modified By'] = ADVISOR_USER_EMAIL;
        
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
    setDrawerOpen(false);
    state.currentRecordId = null;
}

// Wire events
function setupEventListeners() {
    bindDashboardClickTargets();

    // Text search
    byId('searchInput').addEventListener('input', () => {
        state.currentPage = 1;
        updateUI();
    });

    // Filters
    FILTER_CONTROL_IDS.filter(id => id !== 'searchInput').forEach(id => {
        byId(id).addEventListener('change', () => {
            state.currentPage = 1;
            updateUI();
        });
    });

    // Clear Filters
    byId('clearFiltersBtn').addEventListener('click', () => {
        resetControls();
        state.dashboardFilter = null;
        state.currentPage = 1;
        updateUI();
    });

    // Pagination
    byId('prevPageBtn').addEventListener('click', () => {
        if (state.currentPage > 1) {
            state.currentPage--;
            renderTable();
        }
    });

    byId('nextPageBtn').addEventListener('click', () => {
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

    // Theme Toggle Switch
    const themeSwitch = byId('themeToggleSwitch');
    if (themeSwitch) {
        themeSwitch.addEventListener('change', () => {
            state.theme = themeSwitch.checked ? 'dark-theme' : 'light-theme';
            applyThemeClass(state.theme);
            localStorage.setItem(THEME_STORAGE_KEY, state.theme);
        });
    }

    // Export JSON
    byId('exportJsonBtn').addEventListener('click', exportDatabaseToJson);

    // Reset Data
    byId('resetDataBtn').addEventListener('click', () => {
        if (confirm("Are you sure you want to reset the admissions database to its original Excel state? All your custom modifications will be lost.")) {
            localStorage.removeItem(LOCAL_STORAGE_KEY);
            localStorage.removeItem(SEED_VARIANT_STORAGE_KEY);
            loadRecords();
            updateUI();
        }
    });

    // New record profile
    byId('addApplicantBtn').addEventListener('click', openNewRecord);

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
        const el = byId(id);
        if (el) {
            el.addEventListener('change', validateBusinessRules);
        }
    });

    // Drawer closing bindings
    byId('closeDrawerBtn').addEventListener('click', closeDrawer);
    byId('cancelDrawerBtn').addEventListener('click', closeDrawer);
    byId('drawerBackdrop').addEventListener('click', closeDrawer);

    // Form save action
    byId('saveRecordBtn').addEventListener('click', saveRecordChanges);
    
    // Delete action
    byId('deleteRecordBtn').addEventListener('click', deleteRecord);
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
