// --- CONFIGURATION ---
const VILLAGES = [
    "জালশুকা", "গাড়াডোব", "আযান", "যুগিন্দা", "বাহাগুন্দা", "ঢেপা",
    "পাকুড়িয়া", "চিৎলা", "নিত্যনন্দনপুর", "খরমপুর", "কচুইখালি",
    "ভাটপড়া", "দীঘলকান্দি", "কসবা", "ধানখোলা", "দক্ষিণ তেতুলবাড়িয়া",
    "মহিষাখোলা", "আড়পাড়া", "চান্দামারী", "বেড়", "শানঘাট"
];

const EXCEL_PATH = "excel/"; 

// --- STATE MANAGEMENT ---
let currentVillage = "";
let currentGender = "M";
let currentData = []; // Main Data Store
let filteredData = []; // For Search View

// --- INITIALIZATION ---
window.onload = function() {
    // 1. Setup Village Dropdown
    const select = document.getElementById('loginVillageSelect');
    select.innerHTML = "";
    VILLAGES.forEach(v => {
        let opt = document.createElement('option');
        opt.value = v;
        opt.innerText = v;
        select.appendChild(opt);
    });

    // 2. Event Listeners
    document.getElementById('searchInput').addEventListener('keyup', handleLiveSearch);
    
    // 3. CHECK AUTO LOGIN (Persistent Login)
    checkAutoLogin();
};

// --- AUTH SYSTEM (Login/Logout) ---
function checkAutoLogin() {
    const savedVillage = localStorage.getItem('vvdas_village');
    const savedUser = localStorage.getItem('vvdas_user');
    const savedPass = localStorage.getItem('vvdas_pass');

    if(savedVillage && savedUser && savedPass) {
        // Auto fill and trigger login
        document.getElementById('loginVillageSelect').value = savedVillage;
        document.getElementById('username').value = savedUser;
        document.getElementById('password').value = savedPass;
        performLogin(true); // true = silent mode
    }
}

function performLogin(isAuto = false) {
    const village = document.getElementById('loginVillageSelect').value;
    const u = document.getElementById('username').value.trim();
    const p = document.getElementById('password').value.trim();
    const err = document.getElementById('loginError');

    // Fetch credentials securely (In real app, use server-side validation)
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

            if (allow) {
                // SAVE CREDENTIALS FOR NEXT TIME
                localStorage.setItem('vvdas_village', village);
                localStorage.setItem('vvdas_user', u);
                localStorage.setItem('vvdas_pass', p);

                currentVillage = village;
                document.getElementById('loginModal').style.display = 'none';
                document.getElementById('mainApp').style.display = 'block';
                
                updateHeader();
                triggerFileLoad();
            } else {
                if(!isAuto) err.style.display = 'block';
                // If auto login fails, clear garbage data
                if(isAuto) logout(); 
            }
        })
        .catch(() => {
            if(!isAuto) {
                err.innerText = "❌ সার্ভার এরর অথবা ফাইল মিসিং!";
                err.style.display = 'block';
            }
        });
}

function logout() {
    localStorage.removeItem('vvdas_village');
    localStorage.removeItem('vvdas_user');
    localStorage.removeItem('vvdas_pass');
    location.reload();
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
    
    // Reset Views
    document.getElementById('searchInput').value = '';
    document.getElementById('detailViewSection').style.display = 'none';
    document.getElementById('siblingSection').style.display = 'none';
    updateHeader();
    triggerFileLoad();
}

// --- FILE HANDLING ---
function triggerFileLoad() {
    const statusDiv = document.getElementById('statusSection');
    const statusMsg = document.getElementById('statusMessage');
    const uploadBox = document.getElementById('uploadContainer');
    const dataArea = document.getElementById('dataArea');

    statusDiv.style.display = 'block';
    uploadBox.style.display = 'none';
    dataArea.style.display = 'none';
    statusMsg.innerHTML = `🔄 ডেটা লোড হচ্ছে... (${currentVillage}_${currentGender}.xlsx)`;

    const fileName = `${currentVillage}_${currentGender}.xlsx`;
    
    fetch(EXCEL_PATH + fileName)
        .then(res => {
            if(!res.ok) throw new Error("File Not Found");
            return res.arrayBuffer();
        })
        .then(data => {
            processExcel(data);
            statusDiv.style.display = 'none';
            dataArea.style.display = 'block';
        })
        .catch(() => {
            statusMsg.innerHTML = `<span style="color:var(--red)">⚠️ অটো লোড ব্যর্থ! ফাইল ম্যানুয়ালি আপলোড করুন:</span>`;
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
    renderTable(currentData); // Render ALL Data initially
}

// --- TABLE RENDERING SYSTEM (Optimized for Large Data) ---
function renderTable(data) {
    const tbody = document.getElementById('tableBody');
    const countBadge = document.getElementById('totalDataCount');
    
    // Update English Counter
    countBadge.innerText = `Total Voters: ${data.length}`;

    // Optimization: Build HTML string instead of createElements loop for speed
    if(data.length === 0) {
        tbody.innerHTML = `<tr><td colspan="5" style="text-align:center; padding: 20px;">কোন তথ্য পাওয়া যায়নি</td></tr>`;
        return;
    }

    // Limit rendering for safety if > 5000 rows, otherwise render all
    // Since user asked for ALL, we use map join. 
    // WARNING: huge datasets might lag slightly.
    
    const rowsHTML = data.map((row) => {
        // Find original index for "View" button to work correctly with Details
        const originalIndex = currentData.indexOf(row);
        
        const sl = row['ক্রমিক'] || row['SL'] || Object.values(row)[0] || '-';
        const name = row['নাম'] || row['Name'] || Object.values(row)[1] || '-';
        const father = row['পিতা'] || row['স্বামী'] || Object.values(row)[2] || '-';
        const mother = row['মাতা'] || row['Mother'] || Object.values(row)[3] || '-';

        return `
            <tr>
                <td><button class="view-btn" onclick="showDetail('${originalIndex}')">👁️</button></td>
                <td>${sl}</td>
                <td class="name-cell">${name}</td>
                <td>${father}</td>
                <td>${mother}</td>
            </tr>
        `;
    }).join('');

    tbody.innerHTML = rowsHTML;
}

// --- SEARCH LOGIC (Filters Table Directly) ---
function handleLiveSearch(e) {
    const val = e.target.value.toLowerCase().trim();
    const type = document.querySelector('input[name="searchType"]:checked').value;
    
    if(val.length === 0) {
        renderTable(currentData); // Restore full table
        return;
    }

    // Filter Logic
    const matches = currentData.filter(row => {
        let text = '';
        // 0=SL, 1=Name, 2=Father
        if(type === 'sl') text = String(row['ক্রমিক'] || Object.values(row)[0]);
        if(type === 'name') text = String(row['নাম'] || Object.values(row)[1]);
        if(type === 'father') text = String(row['পিতা'] || Object.values(row)[2]);
        
        return text.toLowerCase().includes(val);
    });

    renderTable(matches); // Update Table with Filtered Results
}

// --- HELPERS (Bangla Number Removed for Total Count, kept for others if needed) ---
function en2bn(num) {
    return String(num).replace(/\d/g, d => "০১২৩৪৫৬৭৮৯"[d]);
}

// --- DETAIL VIEW LOGIC ---
function showDetail(index) {
    const row = currentData[index];
    const section = document.getElementById('detailViewSection');
    const siblingSec = document.getElementById('siblingSection');
    
    // Extract Data
    const sl = row['ক্রমিক'] || Object.values(row)[0] || '-';
    const name = row['নাম'] || Object.values(row)[1] || '-';
    const father = row['পিতা'] || row['স্বামী'] || Object.values(row)[2] || '-';
    const mother = row['মাতা'] || Object.values(row)[3] || '-';
    const voterNo = row['ভোটার নং'] || row['Voter No'] || Object.values(row)[4] || 'N/A';

    section.style.display = 'block';
    siblingSec.style.display = 'none';

    section.innerHTML = `
        <div class="detail-card-grid">
            <div class="serial-box">${sl}</div>
            <div class="info-box">
                <div class="name-row">${name}</div>
                <div class="parent-row">
                    <div>
                        <span class="data-label">পিতা/স্বামী</span>
                        <span class="data-val father-link" onclick="findRelatives('${father.replace(/'/g, "\\'")}', 'father')">
                            ${father} ↗
                        </span>
                    </div>
                    <div>
                        <span class="data-label">মাতা</span>
                        <span class="data-val father-link" onclick="findRelatives('${mother.replace(/'/g, "\\'")}', 'mother')">
                            ${mother} ↗
                        </span>
                    </div>
                </div>
                <div class="voter-row">
                    VOTER ID: <span style="color:#fff; font-weight:bold;">${voterNo}</span>
                </div>
            </div>
            <button onclick="document.getElementById('detailViewSection').style.display='none'" class="close-detail">✖</button>
        </div>
    `;

    section.scrollIntoView({behavior: 'smooth', block: 'center'});
}

function findRelatives(parentName, type) {
    if(!parentName || parentName.length < 2) return;
    
    const relatives = currentData.filter(row => {
        let pName = '';
        if(type === 'father') pName = row['পিতা'] || row['স্বামী'] || Object.values(row)[2] || '';
        if(type === 'mother') pName = row['মাতা'] || Object.values(row)[3] || '';
        return String(pName).trim() === String(parentName).trim();
    });

    const box = document.getElementById('siblingSection');
    const list = document.getElementById('siblingList');
    list.innerHTML = '';
    
    const labelText = type === 'father' ? 'পিতার নামে' : 'মায়ের নামে';
    box.querySelector('small').innerText = `${labelText} মিল পাওয়া সদস্যগণ (${relatives.length} জন):`;

    if(relatives.length > 0) {
        box.style.display = 'block';
        relatives.forEach(sib => {
            const sName = sib['নাম'] || Object.values(sib)[1];
            const sSL = sib['ক্রমিক'] || Object.values(sib)[0];
            const realIndex = currentData.indexOf(sib);

            let tag = document.createElement('span');
            tag.className = 'sibling-tag';
            tag.innerHTML = `${sName} <span style="color:var(--cyan)">(${sSL})</span>`;
            tag.onclick = () => showDetail(realIndex);
            list.appendChild(tag);
        });
        box.scrollIntoView({behavior: 'smooth', block: 'center'});
    }
}
