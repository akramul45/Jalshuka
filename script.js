// --- CONFIGURATION ---
const VILLAGES = [
    "জালশুকা", "গাড়াডোব", "আযান", "যুগিন্দা", "বাহাগুন্দা", "ঢেপা",
    "পাকুড়িয়া", "চিৎলা", "নিত্যনন্দনপুর", "খরমপুর", "কচুইখালি",
    "ভাটপড়া", "দীঘলকান্দি", "কসবা", "ধানখোলা", "দক্ষিণ তেতুলবাড়িয়া",
    "মহিষাখোলা", "আড়পাড়া", "চান্দামারী", "বেড়", "শানঘাট"
];

const DEFAULT_VILLAGE = "জালশুকা";
const EXCEL_PATH = "excel/"; 

// --- STATE MANAGEMENT ---
let currentVillage = DEFAULT_VILLAGE;
let currentGender = "M"; // Default to Male file
let currentData = [];
let unlockedVillages = JSON.parse(localStorage.getItem('vvdas_guardian_keys')) || [];

// --- INITIALIZATION ---
window.onload = function() {
    populateVillageDropdown();
    
    // Check initial load
    handleVillageChange();
    
    // Search Listener
    document.getElementById('searchInput').addEventListener('keyup', handleSearch);
    document.addEventListener('click', (e) => {
        if (!e.target.closest('.search-input-box')) {
            document.getElementById('suggestionBox').style.display = 'none';
        }
    });
};

function populateVillageDropdown() {
    const select = document.getElementById('villageSelect');
    select.innerHTML = "";
    VILLAGES.forEach(v => {
        let opt = document.createElement('option');
        opt.value = v;
        opt.innerText = v;
        if(v === DEFAULT_VILLAGE) opt.selected = true;
        select.appendChild(opt);
    });
}

// --- CORE LOGIC FLOW ---

function handleVillageChange() {
    const village = document.getElementById('villageSelect').value;
    currentVillage = village;
    
    // Update Page Title (SEO & Tab)
    updatePageTitle();
    
    // Check Security
    if (checkAccess(village)) {
        document.getElementById('loginModal').style.display = 'none';
        triggerFileLoad();
    } else {
        showLoginModal(village);
    }
}

function updatePageTitle() {
    const genderText = currentGender === 'M' ? 'পুরুষ' : 'নারী';
    document.title = `VVDAS - ${currentVillage} (${genderText})`;
    document.getElementById('pageTitle').innerText = `${currentVillage} গ্রাম`;
    document.getElementById('pageSubTitle').innerText = `ভোটার তালিকা: ${genderText}`;
}

function changeGender(gender) {
    currentGender = gender;
    document.getElementById('btnM').className = gender === 'M' ? 'toggle-btn active' : 'toggle-btn';
    document.getElementById('btnF').className = gender === 'F' ? 'toggle-btn active' : 'toggle-btn';
    
    updatePageTitle();
    
    if (document.getElementById('loginModal').style.display === 'none') {
        triggerFileLoad();
    }
}

// --- SECURITY & LOGIN ---

function checkAccess(village) {
    return unlockedVillages.includes('all') || unlockedVillages.includes(village);
}

function showLoginModal(village) {
    const modal = document.getElementById('loginModal');
    document.getElementById('loginSubtitle').innerText = `${village} ডেটাবেস লক করা আছে`;
    document.getElementById('tableBody').innerHTML = ''; // Clear table for security
    currentData = [];
    modal.style.display = 'flex';
}

function performLogin() {
    const u = document.getElementById('username').value.trim();
    const p = document.getElementById('password').value.trim();
    const err = document.getElementById('loginError');
    
    fetch('CD.json')
        .then(res => res.json())
        .then(users => {
            const user = users.find(acc => acc.username === u && acc.password === p);
            
            if(user) {
                // Success
                if(user.access === 'all') {
                    if(!unlockedVillages.includes('all')) unlockedVillages.push('all');
                } else if(user.access === currentVillage) {
                    if(!unlockedVillages.includes(currentVillage)) unlockedVillages.push(currentVillage);
                } else {
                    err.innerText = "❌ এই গ্রামের অনুমতি নেই!";
                    err.style.display = 'block';
                    return;
                }
                
                // Save and Unlock
                localStorage.setItem('vvdas_guardian_keys', JSON.stringify(unlockedVillages));
                document.getElementById('loginModal').style.display = 'none';
                document.getElementById('username').value = '';
                document.getElementById('password').value = '';
                err.style.display = 'none';
                
                triggerFileLoad();
            } else {
                err.innerText = "❌ ইউজারনেম বা পাসওয়ার্ড ভুল!";
                err.style.display = 'block';
            }
        })
        .catch(() => {
            err.innerText = "❌ সার্ভার কানেকশন এরর (CD.json)";
            err.style.display = 'block';
        });
}

// --- FILE LOADING SYSTEM ---

function triggerFileLoad() {
    const statusDiv = document.getElementById('statusSection');
    const statusMsg = document.getElementById('statusMessage');
    const uploadBox = document.getElementById('uploadContainer');
    
    statusDiv.style.display = 'block';
    uploadBox.style.display = 'none';
    statusMsg.innerHTML = `<span style="color:var(--guardian-cyan)">🔄 সার্ভার থেকে ডেটা লোড হচ্ছে... (${currentVillage}_${currentGender}.xlsx)</span>`;
    
    const fileName = `${currentVillage}_${currentGender}.xlsx`;
    const filePath = EXCEL_PATH + fileName;

    fetch(filePath)
        .then(res => {
            if(!res.ok) throw new Error("File Missing");
            return res.arrayBuffer();
        })
        .then(data => {
            processExcel(data);
            statusDiv.style.display = 'none';
        })
        .catch(err => {
            console.warn(err);
            statusMsg.innerHTML = `<span style="color:var(--guardian-red)">❌ অটোমেটিক লোড ব্যর্থ হয়েছে!</span>`;
            uploadBox.style.display = 'block'; // Show manual upload
        });
}

function handleFileUpload(input) {
    const file = input.files[0];
    if(!file) return;
    
    const reader = new FileReader();
    reader.onload = function(e) {
        processExcel(e.target.result);
        document.getElementById('statusSection').style.display = 'none';
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

// --- TABLE & SEARCH ---

function renderTable(data) {
    const tbody = document.getElementById('tableBody');
    tbody.innerHTML = '';

    data.forEach((row, index) => {
        // Dynamic Key Matching
        const sl = row['ক্রমিক'] || row['SL'] || Object.values(row)[0] || '-';
        const name = row['নাম'] || row['Name'] || Object.values(row)[1] || '-';
        const father = row['পিতা'] || row['স্বামী'] || row['Father'] || Object.values(row)[2] || '-';
        
        let tr = document.createElement('tr');
        tr.id = `row-${index}`;
        tr.innerHTML = `
            <td>${sl}</td>
            <td style="color:var(--guardian-gold); font-weight:600;">${name}</td>
            <td>${father}</td>
            <td><button class="guardian-btn" style="padding:5px 10px; font-size:12px; width:auto;" onclick='openDetails(${JSON.stringify(row)})'>👁️</button></td>
        `;
        tbody.appendChild(tr);
    });
}

function handleSearch(e) {
    const val = e.target.value.toLowerCase().trim();
    const type = document.querySelector('input[name="searchType"]:checked').value;
    const box = document.getElementById('suggestionBox');
    
    if(val.length === 0) {
        box.style.display = 'none';
        return;
    }

    const matches = currentData.map((row, index) => ({row, index})).filter(item => {
        let text = '';
        if(type === 'sl') text = String(item.row['ক্রমিক'] || Object.values(item.row)[0]);
        if(type === 'name') text = String(item.row['নাম'] || Object.values(item.row)[1]);
        if(type === 'father') text = String(item.row['পিতা'] || Object.values(item.row)[2]);
        
        return text.toLowerCase().includes(val);
    });

    box.innerHTML = '';
    if(matches.length > 0) {
        box.style.display = 'block';
        matches.slice(0, 8).forEach(m => {
            const name = m.row['নাম'] || Object.values(m.row)[1];
            const div = document.createElement('div');
            div.className = 'suggestion-item';
            div.innerHTML = `<span>${name}</span> <small style="color:var(--guardian-cyan)">${m.row['ক্রমিক']||'SL'}</small>`;
            div.onclick = () => {
                const el = document.getElementById(`row-${m.index}`);
                el.scrollIntoView({behavior: 'smooth', block: 'center'});
                el.style.background = 'rgba(6, 182, 212, 0.3)';
                setTimeout(() => el.style.background = '', 2000);
                box.style.display = 'none';
            };
            box.appendChild(div);
        });
    } else {
        box.style.display = 'none';
    }
}

// --- DETAILS ---
function openDetails(row) {
    const card = document.getElementById('detailsCard');
    const content = document.getElementById('detailsContent');
    content.innerHTML = '';
    
    Object.entries(row).forEach(([k, v]) => {
        content.innerHTML += `
            <div class="info-item">
                <span class="info-label">${k}</span>
                <span class="info-val">${v}</span>
            </div>
        `;
    });
    card.classList.add('active');
}

function closeDetails() {
    document.getElementById('detailsCard').classList.remove('active');
}