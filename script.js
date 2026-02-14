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
let currentData = []; // Stores all rows
let unlockedVillages = JSON.parse(localStorage.getItem('vvdas_guardian_keys')) || [];

// --- INITIALIZATION ---
window.onload = function() {
    // 1. Login Modal - Village Dropdown
    const select = document.getElementById('loginVillageSelect');
    select.innerHTML = "";
    VILLAGES.forEach(v => {
        let opt = document.createElement('option');
        opt.value = v;
        opt.innerText = v;
        select.appendChild(opt);
    });

    // 2. Search Event Listener
    document.getElementById('searchInput').addEventListener('keyup', handleLiveSearch);
    
    // Close suggestion box when clicking outside
    document.addEventListener('click', function(e) {
        if (!e.target.closest('.search-input-box')) {
            document.getElementById('suggestionBox').style.display = 'none';
        }
    });
};

// --- LOGIN LOGIC ---
function performLogin() {
    const village = document.getElementById('loginVillageSelect').value;
    const u = document.getElementById('username').value.trim();
    const p = document.getElementById('password').value.trim();
    const err = document.getElementById('loginError');

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
            
            // Allow if previously unlocked (Optional security feature)
            // if(unlockedVillages.includes(village) && u === '' && p === '') allow = true; 

            if (allow) {
                currentVillage = village;
                document.getElementById('loginModal').style.display = 'none';
                document.getElementById('mainApp').style.display = 'block';
                updateHeader();
                triggerFileLoad();
            } else {
                err.style.display = 'block';
            }
        })
        .catch(() => {
            err.innerText = "❌ সার্ভার এরর!";
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
    
    // Reset Views
    document.getElementById('dataArea').style.display = 'none';
    document.getElementById('detailViewSection').style.display = 'none';
    document.getElementById('siblingSection').style.display = 'none';
    updateHeader();
    triggerFileLoad();
}

// --- FILE LOAD SYSTEM ---
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
            statusMsg.innerHTML = `<span style="color:var(--red)">❌ অটো লোড ব্যর্থ! ফাইল আপলোড করুন:</span>`;
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
    
    // Update Total Count
    document.getElementById('totalDataCount').innerText = `মোট ভোটার: ${en2bn(jsonData.length)}`;
    
    renderTable(jsonData);
}

// --- HELPER: English to Bangla Number ---
function en2bn(num) {
    return String(num).replace(/\d/g, d => "০১২৩৪৫৬৭৮৯"[d]);
}

// --- TABLE RENDER (Initial View) ---
function renderTable(data) {
    const tbody = document.getElementById('tableBody');
    tbody.innerHTML = '';
    // Show first 50 rows initially to boost performance
    const initialData = data.slice(0, 50);

    initialData.forEach((row) => {
        // Find Index in main array
        const realIndex = currentData.indexOf(row);
        createTableRow(row, realIndex, tbody);
    });
}

function createTableRow(row, index, parent) {
    const sl = row['ক্রমিক'] || row['SL'] || Object.values(row)[0] || '-';
    const name = row['নাম'] || row['Name'] || Object.values(row)[1] || '-';
    const father = row['পিতা'] || row['স্বামী'] || Object.values(row)[2] || '-';
    const mother = row['মাতা'] || row['Mother'] || Object.values(row)[3] || '-';

    let tr = document.createElement('tr');
    tr.innerHTML = `
        <td><button class="view-btn" onclick="showDetail('${index}')">👁️ ভিউ</button></td>
        <td>${sl}</td>
        <td style="color:var(--gold); font-weight:500;">${name}</td>
        <td>${father}</td>
        <td>${mother}</td>
    `;
    parent.appendChild(tr);
}

// --- LIVE SEARCH LOGIC (Dropdown) ---
function handleLiveSearch(e) {
    const val = e.target.value.toLowerCase().trim();
    const type = document.querySelector('input[name="searchType"]:checked').value;
    const box = document.getElementById('suggestionBox');
    
    // Hide if empty
    if(val.length === 0) {
        box.style.display = 'none';
        renderTable(currentData); // Restore full table
        return;
    }

    // Filter Logic (Partial Match)
    const matches = currentData.map((row, index) => ({row, index})).filter(item => {
        let text = '';
        // 0=SL, 1=Name, 2=Father
        if(type === 'sl') text = String(item.row['ক্রমিক'] || Object.values(item.row)[0]);
        if(type === 'name') text = String(item.row['নাম'] || Object.values(item.row)[1]);
        if(type === 'father') text = String(item.row['পিতা'] || Object.values(item.row)[2]);
        
        return text.toLowerCase().includes(val);
    });

    // Populate Dropdown
    box.innerHTML = '';
    if(matches.length > 0) {
        box.style.display = 'block';
        
        // Add Header
        let header = document.createElement('div');
        header.className = 'suggestion-header-label';
        header.innerHTML = `<span>ক্রমিক</span> <span>নাম</span> <span>পিতা</span>`;
        box.appendChild(header);

        // Limit to 10 suggestions for speed
        matches.slice(0, 10).forEach(m => {
            const sl = m.row['ক্রমিক'] || Object.values(m.row)[0];
            const name = m.row['নাম'] || Object.values(m.row)[1];
            const father = m.row['পিতা'] || Object.values(m.row)[2];

            const div = document.createElement('div');
            div.className = 'suggestion-item';
            div.innerHTML = `
                <span style="color:var(--cyan)">${sl}</span>
                <span style="font-weight:bold; color:var(--gold)">${name}</span>
                <span>${father}</span>
            `;
            div.onclick = () => {
                showDetail(m.index);
                document.getElementById('searchInput').value = ''; // Clear input
                box.style.display = 'none';
            };
            box.appendChild(div);
        });
    } else {
        box.style.display = 'none';
    }
}

// --- DETAIL CARD & RELATIVES LOGIC ---
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

    // Show Section
    section.style.display = 'block';
    siblingSec.style.display = 'none'; // Reset siblings

    // Render Card (Added Mother OnClick)
    section.innerHTML = `
        <div class="detail-card-grid">
            <div class="serial-box">${sl}</div>
            <div class="info-box">
                <div class="name-row">${name}</div>
                <div class="parent-row">
                    <div>
                        <span class="data-label">পিতা/স্বামী</span>
                        <span class="data-val father-link" onclick="findRelatives('${father.replace(/'/g, "\\'")}', 'father')">
                            ${father} (ক্লিক)
                        </span>
                    </div>
                    <div>
                        <span class="data-label">মাতা</span>
                        <span class="data-val father-link" onclick="findRelatives('${mother.replace(/'/g, "\\'")}', 'mother')">
                            ${mother} (ক্লিক)
                        </span>
                    </div>
                </div>
                <div class="voter-row">
                    VOTER ID: <span style="color:white; font-weight:bold;">${voterNo}</span>
                </div>
            </div>
        </div>
    `;

    // Scroll to Details
    section.scrollIntoView({behavior: 'smooth', block: 'center'});
}

// Unified Function for Father/Mother Search
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
    
    // Label update based on type
    const labelText = type === 'father' ? 'পিতার নামে' : 'মায়ের নামে';
    box.querySelector('small').innerText = `${labelText} মিল পাওয়া সদস্যগণ (${en2bn(relatives.length)} জন):`;

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
    }
}
