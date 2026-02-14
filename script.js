// --- CONFIGURATION ---
const VILLAGES = [
    "জালশুকা", "গাড়াডোব", "আযান", "যুগিন্দা", "বাহাগুন্দা", "ঢেপা",
    "পাকুড়িয়া", "চিৎলা", "নিত্যনন্দনপুর", "খরমপুর", "কচুইখালি",
    "ভাটপড়া", "দীঘলকান্দি", "কসবা", "ধানখোলা", "দক্ষিণ তেতুলবাড়িয়া",
    "মহিষাখোলা", "আড়পাড়া", "চান্দামারী", "বেড়", "শানঘাট"
];

const DEFAULT_VILLAGE = "জালশুকা";
const EXCEL_PATH = "excel/"; 

// --- STATE ---
let currentVillage = "";
let currentGender = "M";
let currentData = [];
let unlockedVillages = JSON.parse(localStorage.getItem('vvdas_guardian_keys')) || [];

// --- INITIALIZATION ---
window.onload = function() {
    // 1. Populate Dropdown in Login Modal
    const select = document.getElementById('loginVillageSelect');
    select.innerHTML = "";
    VILLAGES.forEach(v => {
        let opt = document.createElement('option');
        opt.value = v;
        opt.innerText = v;
        select.appendChild(opt);
    });

    // 2. Setup Search Listener
    document.getElementById('searchInput').addEventListener('keyup', handleSearch);
};

// --- LOGIN LOGIC ---
function performLogin() {
    const village = document.getElementById('loginVillageSelect').value;
    const u = document.getElementById('username').value.trim();
    const p = document.getElementById('password').value.trim();
    const err = document.getElementById('loginError');

    // Simulate Server Check (Replace with real logic if needed)
    fetch('CD.json')
        .then(res => res.json())
        .then(users => {
            const user = users.find(acc => acc.username === u && acc.password === p);
            
            let allow = false;
            if(user) {
                if(user.access === 'all' || user.access === village) {
                    allow = true;
                }
            }

            if(allow || (unlockedVillages.includes(village) && u==='' && p==='')) { 
                // Allow re-login if already unlocked in this browser (optional convenience)
                // BUT stricter: must enter pass.
            }
            
            if (allow) {
                // SUCCESS
                currentVillage = village;
                
                // Hide Modal, Show App
                document.getElementById('loginModal').style.display = 'none';
                document.getElementById('mainApp').style.display = 'block';
                
                // Update Header
                updateHeader();
                
                // Trigger Load
                triggerFileLoad();
            } else {
                err.style.display = 'block';
            }
        })
        .catch(() => {
            err.innerText = "❌ সার্ভার এরর (CD.json missing)";
            err.style.display = 'block';
        });
}

function updateHeader() {
    const gText = currentGender === 'M' ? 'পুরুষ' : 'নারী';
    document.getElementById('pageTitle').innerText = `${currentVillage} গ্রাম`;
    document.getElementById('pageSubTitle').innerText = `ভোটার তালিকা: ${gText}`;
}

function changeGender(gender) {
    currentGender = gender;
    document.getElementById('btnM').className = gender === 'M' ? 'toggle-btn active' : 'toggle-btn';
    document.getElementById('btnF').className = gender === 'F' ? 'toggle-btn active' : 'toggle-btn';
    updateHeader();
    
    // Clear previous data views
    document.getElementById('dataArea').style.display = 'none';
    document.getElementById('detailViewSection').style.display = 'none';
    document.getElementById('siblingSection').style.display = 'none';
    
    triggerFileLoad();
}

// --- FILE SYSTEM ---
function triggerFileLoad() {
    const statusDiv = document.getElementById('statusSection');
    const statusMsg = document.getElementById('statusMessage');
    const uploadBox = document.getElementById('uploadContainer');
    const dataArea = document.getElementById('dataArea');

    // Reset UI
    statusDiv.style.display = 'block';
    uploadBox.style.display = 'none';
    dataArea.style.display = 'none'; // Hide table/search until load success
    statusMsg.innerHTML = `🔄 ডেটা লোড হচ্ছে... (${currentVillage}_${currentGender}.xlsx)`;

    const fileName = `${currentVillage}_${currentGender}.xlsx`;
    
    fetch(EXCEL_PATH + fileName)
        .then(res => {
            if(!res.ok) throw new Error("Not Found");
            return res.arrayBuffer();
        })
        .then(data => {
            processExcel(data);
            statusDiv.style.display = 'none';
            dataArea.style.display = 'block'; // Show content
        })
        .catch(() => {
            statusMsg.innerHTML = `<span style="color:var(--red)">❌ অটোমেটিক লোড ব্যর্থ! ফাইল আপলোড করুন:</span>`;
            uploadBox.style.display = 'block';
        });
}

function handleFileUpload(input) {
    const file = input.files[0];
    if(!file) return;
    const reader = new FileReader();
    reader.onload = (e) => {
        processExcel(e.target.result);
        document.getElementById('statusSection').style.display = 'none';
        document.getElementById('dataArea').style.display = 'block';
    };
    reader.readAsArrayBuffer(file);
}

function processExcel(buffer) {
    const workbook = XLSX.read(buffer, {type: 'array'});
    const sheetName = workbook.SheetNames[0];
    const jsonData = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {defval: ""});
    currentData = jsonData;
    renderTable(jsonData);
}

// --- RENDERING ---
function renderTable(data) {
    const tbody = document.getElementById('tableBody');
    tbody.innerHTML = '';

    data.forEach((row, index) => {
        // Robust Key Matching
        const sl = row['ক্রমিক'] || row['SL'] || Object.values(row)[0] || '-';
        const name = row['নাম'] || row['Name'] || Object.values(row)[1] || '-';
        const father = row['পিতা'] || row['স্বামী'] || Object.values(row)[2] || '-';
        const mother = row['মাতা'] || row['Mother'] || Object.values(row)[3] || '-';
        
        // Store full row data in button
        const rowDataStr = encodeURIComponent(JSON.stringify(row));

        let tr = document.createElement('tr');
        tr.innerHTML = `
            <td>
                <button class="view-btn" onclick="showDetail('${index}')">👁️ ভিউ</button>
            </td>
            <td>${sl}</td>
            <td style="color:var(--gold); font-weight:500;">${name}</td>
            <td>${father}</td>
            <td>${mother}</td>
        `;
        tbody.appendChild(tr);
    });
}

// --- DETAIL VIEW LOGIC ---
function showDetail(index) {
    const row = currentData[index];
    const section = document.getElementById('detailViewSection');
    const siblingSec = document.getElementById('siblingSection');
    
    // Normalize Data
    const sl = row['ক্রমিক'] || Object.values(row)[0] || '-';
    const name = row['নাম'] || Object.values(row)[1] || '-';
    const father = row['পিতা'] || row['স্বামী'] || Object.values(row)[2] || '-';
    const mother = row['মাতা'] || Object.values(row)[3] || '-';
    const voterNo = row['ভোটার নং'] || row['Voter No'] || Object.values(row)[4] || 'N/A';

    section.style.display = 'block';
    siblingSec.style.display = 'none'; // Reset sibling view initially

    section.innerHTML = `
        <div class="detail-card-grid">
            <div class="serial-box">${sl}</div>
            <div class="info-box">
                <div class="name-row">${name}</div>
                <div class="parent-row">
                    <div>
                        <span class="data-label">পিতা/স্বামী</span>
                        <span class="data-val father-link" onclick="findSiblings('${father.replace(/'/g, "\\'")}')">${father} (ক্লিক করুন)</span>
                    </div>
                    <div>
                        <span class="data-label">মাতা</span>
                        <span class="data-val">${mother}</span>
                    </div>
                </div>
                <div class="voter-row">
                    VOTER ID: <span style="color:white; font-weight:bold;">${voterNo}</span>
                </div>
            </div>
        </div>
    `;

    // Scroll to view
    section.scrollIntoView({behavior: 'smooth', block: 'center'});
}

function findSiblings(fatherName) {
    if(!fatherName || fatherName.length < 3) return;
    
    const siblings = currentData.filter(row => {
        const f = row['পিতা'] || row['স্বামী'] || Object.values(row)[2] || '';
        return String(f).trim() === String(fatherName).trim();
    });

    const box = document.getElementById('siblingSection');
    const list = document.getElementById('siblingList');
    list.innerHTML = '';

    if(siblings.length > 0) {
        box.style.display = 'block';
        siblings.forEach(sib => {
            const sName = sib['নাম'] || Object.values(sib)[1];
            const sSL = sib['ক্রমিক'] || Object.values(sib)[0];
            
            // Find index in main data to link view
            const realIndex = currentData.indexOf(sib);

            let tag = document.createElement('span');
            tag.className = 'sibling-tag';
            tag.innerHTML = `${sName} <span style="color:var(--cyan)">(${sSL})</span>`;
            tag.onclick = () => showDetail(realIndex);
            list.appendChild(tag);
        });
    }
}

// --- SEARCH ---
function handleSearch(e) {
    const val = e.target.value.toLowerCase();
    const rows = document.querySelectorAll('#tableBody tr');

    rows.forEach(row => {
        const text = row.innerText.toLowerCase();
        row.style.display = text.includes(val) ? '' : 'none';
    });
}
