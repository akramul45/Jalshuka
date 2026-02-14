// --- CONFIGURATION ---
const VILLAGES = [
    "জালশুকা", "গাড়াডোব", "আযান", "যুগিন্দা", "বাহাগুন্দা", "ঢেপা",
    "পাকুড়িয়া", "চিৎলা", "নিত্যনন্দনপুর", "খরমপুর", "কচুইখালি",
    "ভাটপড়া", "দীঘলকান্দি", "কসবা", "ধানখোলা", "দক্ষিণ তেতুলবাড়িয়া",
    "মহিষাখোলা", "আড়পাড়া", "চান্দামারী", "বেড়", "শানঘাট"
];

const DEFAULT_VILLAGE = "জালশুকা";
const EXCEL_PATH = "excel/"; // Folder where .xlsx files are stored

// --- STATE ---
let currentVillage = DEFAULT_VILLAGE;
let currentGender = "M"; // M or F
let currentData = []; // Holds the full JSON data of current file
let unlockedVillages = JSON.parse(localStorage.getItem('vvdas_unlocked')) || [];

// --- INITIALIZATION ---
window.onload = function() {
    initVillageDropdown();
    checkAuthAndLoad(currentVillage);
    
    // Suggestion Search Listeners
    document.getElementById('searchInput').addEventListener('keyup', handleSearchInput);
    document.addEventListener('click', (e) => {
        if (!e.target.closest('.input-wrapper')) {
            document.getElementById('suggestionBox').style.display = 'none';
        }
    });
};

function initVillageDropdown() {
    const select = document.getElementById('villageSelect');
    VILLAGES.forEach(v => {
        let opt = document.createElement('option');
        opt.value = v;
        opt.innerText = v;
        if(v === DEFAULT_VILLAGE) opt.selected = true;
        select.appendChild(opt);
    });
}

// --- AUTHENTICATION SYSTEM ---

function checkAuthAndLoad(village) {
    currentVillage = village;
    const modal = document.getElementById('loginModal');
    const msg = document.getElementById('loginMessage');
    
    // Check if user is Super Admin or has unlocked this village
    if (unlockedVillages.includes('all') || unlockedVillages.includes(village)) {
        modal.style.display = 'none';
        loadExcelData(village, currentGender);
    } else {
        // Show Login
        msg.innerText = `${village} গ্রামের তথ্য সুরক্ষিত। পাসওয়ার্ড দিন।`;
        modal.style.display = 'flex';
        // Clear background data for security
        document.getElementById('tableBody').innerHTML = '';
        currentData = [];
    }
}

function performLogin() {
    const user = document.getElementById('username').value.trim();
    const pass = document.getElementById('password').value.trim();
    const err = document.getElementById('loginError');

    fetch('CD.json')
        .then(res => res.json())
        .then(users => {
            const validUser = users.find(u => u.username === user && u.password === pass);
            
            if (validUser) {
                // Login Success
                if (validUser.access === 'all') {
                    if (!unlockedVillages.includes('all')) unlockedVillages.push('all');
                } else {
                    if (!unlockedVillages.includes(validUser.access)) unlockedVillages.push(validUser.access);
                }
                
                // Check if the unlocked village matches current request
                if (validUser.access === 'all' || validUser.access === currentVillage) {
                    localStorage.setItem('vvdas_unlocked', JSON.stringify(unlockedVillages));
                    document.getElementById('loginModal').style.display = 'none';
                    loadExcelData(currentVillage, currentGender);
                    
                    // Clear inputs
                    document.getElementById('username').value = '';
                    document.getElementById('password').value = '';
                    err.style.display = 'none';
                } else {
                    err.innerText = "আপনার এই গ্রামে প্রবেশাধিকার নেই!";
                    err.style.display = 'block';
                }
            } else {
                err.innerText = "ভুল ইউজারনেম বা পাসওয়ার্ড!";
                err.style.display = 'block';
            }
        })
        .catch(e => {
            alert("Database Error: CD.json not found.");
            console.error(e);
        });
}

// --- DATA HANDLING ---

function changeVillage() {
    const v = document.getElementById('villageSelect').value;
    checkAuthAndLoad(v);
}

function changeGender(gender) {
    currentGender = gender;
    // Update buttons
    document.getElementById('btnM').className = gender === 'M' ? 'active' : '';
    document.getElementById('btnF').className = gender === 'F' ? 'active' : '';
    
    // Reload data if already authenticated
    if (document.getElementById('loginModal').style.display === 'none') {
        loadExcelData(currentVillage, currentGender);
    }
}

function loadExcelData(village, gender) {
    const tableBody = document.getElementById('tableBody');
    const loading = document.getElementById('loading');
    
    tableBody.innerHTML = '';
    loading.style.display = 'block';
    
    // English filename mapping needed or use Bengali directly if file system supports
    // Assuming filenames like "Garadob_M.xlsx" or "গাড়াডোব_M.xlsx". 
    // Here using Bengali name + suffix as per prompt requirement implies mapped names or direct use.
    // For safety, let's assume the files are named exactly as village string + "_" + gender + ".xlsx"
    
    // Mapping for English filenames if needed (Optional, otherwise use village variable directly)
    // const fileName = `${getEnglishName(village)}_${gender}.xlsx`; 
    const fileName = `${village}_${gender}.xlsx`; 

    const filePath = EXCEL_PATH + fileName;

    fetch(filePath)
        .then(res => {
            if(!res.ok) throw new Error("File not found");
            return res.arrayBuffer();
        })
        .then(data => {
            const workbook = XLSX.read(data, {type: 'array'});
            const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
            // Convert to JSON array
            const jsonData = XLSX.utils.sheet_to_json(firstSheet, {defval: ""});
            
            currentData = jsonData;
            renderTable(currentData);
            loading.style.display = 'none';
        })
        .catch(err => {
            loading.style.display = 'none';
            tableBody.innerHTML = `<tr><td colspan="4" style="text-align:center; color:red;">তথ্য পাওয়া যায়নি (ফাইল মিসিং: ${fileName})</td></tr>`;
            console.error(err);
        });
}

function renderTable(data) {
    const tbody = document.getElementById('tableBody');
    tbody.innerHTML = ''; // Clear existing
    
    // Assuming Column Headers in Excel: "SL", "Name", "Father", "OtherInfo"
    // Adjust keys based on your actual Excel headers
    
    data.forEach((row, index) => {
        // Detect keys dynamically or fix them
        const keys = Object.keys(row);
        const sl = row['ক্রমিক'] || row['SL'] || row['No'] || keys[0] ? row[keys[0]] : '';
        const name = row['নাম'] || row['Name'] || keys[1] ? row[keys[1]] : '';
        const father = row['পিতা'] || row['স্বামী'] || row['Father'] || keys[2] ? row[keys[2]] : '';
        
        let tr = document.createElement('tr');
        tr.id = `row-${index}`;
        tr.onclick = () => showDetails(row);
        
        tr.innerHTML = `
            <td>${sl}</td>
            <td style="color:var(--accent); font-weight:500;">${name}</td>
            <td>${father}</td>
            <td><button class="view-btn" style="padding:2px 8px; font-size:12px;">দেখুন</button></td>
        `;
        tbody.appendChild(tr);
    });
}

// --- ADVANCED SEARCH & SUGGESTION ---

function handleSearchInput(e) {
    const val = e.target.value.toLowerCase().trim();
    const type = document.querySelector('input[name="searchType"]:checked').value;
    const box = document.getElementById('suggestionBox');
    
    if (val.length === 0) {
        box.style.display = 'none';
        return;
    }

    // Filter Logic based on selected type ONLY
    const matches = currentData.filter((row, index) => {
        row.__index = index; // Store original index for scrolling
        let target = '';

        if (type === 'sl') {
            target = String(row['ক্রমিক'] || row['SL'] || Object.values(row)[0] || '');
        } else if (type === 'name') {
            target = String(row['নাম'] || row['Name'] || Object.values(row)[1] || '');
        } else if (type === 'father') {
            target = String(row['পিতা'] || row['স্বামী'] || row['Father'] || Object.values(row)[2] || '');
        }

        return target.toLowerCase().includes(val);
    });

    // Render Suggestions (Max 10)
    box.innerHTML = '';
    if (matches.length > 0) {
        box.style.display = 'block';
        matches.slice(0, 10).forEach(match => {
            const sl = match['ক্রমিক'] || match['SL'] || Object.values(match)[0];
            const name = match['নাম'] || match['Name'] || Object.values(match)[1];
            const father = match['পিতা'] || match['স্বামী'] || Object.values(match)[2];

            let div = document.createElement('div');
            div.className = 'suggestion-item';
            div.innerHTML = `
                <span><b>${sl}</b> - ${name}</span>
                <span style="font-size:11px; color:#94a3b8;">${father}</span>
            `;
            div.onclick = () => scrollToRow(match.__index);
            box.appendChild(div);
        });
    } else {
        box.style.display = 'none';
    }
}

function scrollToRow(index) {
    const row = document.getElementById(`row-${index}`);
    if (row) {
        row.scrollIntoView({ behavior: 'smooth', block: 'center' });
        
        // Remove previous highlights
        document.querySelectorAll('.highlight-row').forEach(r => r.classList.remove('highlight-row'));
        
        // Add highlight
        row.classList.add('highlight-row');
        
        // Hide suggestion box
        document.getElementById('suggestionBox').style.display = 'none';
        document.getElementById('searchInput').value = ''; // Optional: clear search
    }
}

// --- DETAILS POPUP ---

function showDetails(row) {
    const content = document.getElementById('detailsContent');
    const card = document.getElementById('detailsCard');
    
    let html = `<h3>ভোটার বিস্তারিত</h3><div style="display:grid; grid-template-columns: 1fr 1fr; gap:10px; text-align:left;">`;
    
    for (const [key, value] of Object.entries(row)) {
        if(key === '__index') continue; // Skip internal index
        html += `
            <div style="background:rgba(255,255,255,0.05); padding:8px; border-radius:5px;">
                <small style="color:#94a3b8;">${key}</small><br>
                <span style="color:#f8fafc;">${value}</span>
            </div>
        `;
    }
    html += `</div>`;
    
    content.innerHTML = html;
    card.classList.add('show');
}

function closeDetails() {
    document.getElementById('detailsCard').classList.remove('show');
}