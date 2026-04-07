import { SUPABASE_CONFIG } from './config.js';

const { createClient } = supabase;
const supabaseClient = createClient(SUPABASE_CONFIG.url, SUPABASE_CONFIG.anonKey);

// --- GLOBAL STATE & CACHE ---
let currentTemplate = null;
let allTemplates = [];
let localEntries = [];
let dateSortAsc = true;
const categoryCache = {}; // Stores { template, entries } for every visited module

window.currentTemplate = currentTemplate;
window.supabaseClient = supabaseClient;
window.categoryCache = categoryCache;

// --- INITIALIZATION ---
window.onload = async function() {
    await refreshCategories();
};

// ==========================================
// 1. MODAL CONTROLS 
// ==========================================
// Category Modal
window.openModal = function() {
    const modal = document.getElementById('categoryModal');
    if (modal) {
        modal.style.display = 'block';
        document.getElementById('newCategoryName').focus();
    }
};
window.closeModal = function() {
    const modal = document.getElementById('categoryModal');
    if (modal) {
        modal.style.display = 'none';
        document.getElementById('newCategoryName').value = "";
    }
};

// Column Modal (NEW)
window.openColumnModal = function() {
    const modal = document.getElementById('columnModal');
    if (modal) {
        modal.style.display = 'block';
        document.getElementById('newColumnName').focus();
    }
};
window.closeColumnModal = function() {
    const modal = document.getElementById('columnModal');
    if (modal) {
        modal.style.display = 'none';
        document.getElementById('newColumnName').value = "";
    }
};

// Handle clicking outside of EITHER modal
window.onclick = function(event) {
    const catModal = document.getElementById('categoryModal');
    const colModal = document.getElementById('columnModal');
    
    if (event.target === catModal) closeModal();
    if (event.target === colModal) closeColumnModal();
};

// ==========================================
// UPDATE: ADD COLUMN FUNCTION
// ==========================================
window.addColumnToActive = async function() {
    const colNameInput = document.getElementById('newColumnName');
    const colName = colNameInput.value.trim();
    const colType = document.getElementById('newColumnType').value;
    
    if (!colName) return alert("Please enter a column name.");

    // Visual feedback during save
    const btn = event.target;
    const originalText = btn.innerText;
    btn.innerText = "Adding...";
    btn.disabled = true;

    const newOrder = currentTemplate.doc_columns.length + 1;
    const { error } = await supabaseClient.from('doc_columns').insert([{
        template_id: currentTemplate.id,
        column_name: colName,
        column_type: colType,
        display_order: newOrder
    }]);

    btn.innerText = originalText;
    btn.disabled = false;

    if (!error) {
        delete categoryCache[currentTemplate.name]; 
        colNameInput.value = "";
        
        // NEW: Close the modal automatically after success
        closeColumnModal(); 
        
        await switchCategory(currentTemplate.name); 
    } else {
        alert("Error adding column: " + error.message);
    }
};

// ==========================================
// 2. CATEGORY MANAGEMENT
// ==========================================
async function refreshCategories() {
    const { data, error } = await supabaseClient.from('doc_templates').select('*');
    if (error) return console.error("Error fetching categories:", error);
    
    allTemplates = data || [];
    const container = document.getElementById('categoryCards');
    
    // Generate the cards
    container.innerHTML = allTemplates.map(t => `
        <div class="category-card" id="card-${t.name}" onclick="switchCategory('${t.name}')">
            <div class="card-icon">${t.name.charAt(0).toUpperCase()}</div>
            <span class="card-label">${t.name}</span>
        </div>
    `).join('');
}

window.createNewCategory = async function() {
    const input = document.getElementById('newCategoryName');
    const name = input.value.trim();

    if (!name) return alert("Please enter a category name.");

    // OPTIMIZATION: Disable button to prevent double-submission
    const btn = document.querySelector('.confirm-btn');
    if (btn) {
        btn.disabled = true;
        btn.innerText = "Creating...";
    }

    const { error } = await supabaseClient.from('doc_templates').insert([{ name }]);

    if (btn) {
        btn.disabled = false;
        btn.innerText = "Create Category";
    }

    if (!error) {
        closeModal();
        await refreshCategories(); // Refresh the list without reloading the page
    } else {
        alert("Error: " + error.message);
    }
};

window.addColumnToActive = async function() {
    const colNameInput = document.getElementById('newColumnName');
    const colName = colNameInput.value.trim();
    const colType = document.getElementById('newColumnType').value;
    
    if (!colName) return alert("Please enter a column name.");

    // OPTIMIZATION: Visual feedback during save
    const btn = event.target;
    const originalText = btn.innerText;
    btn.innerText = "Adding...";
    btn.disabled = true;

    const newOrder = currentTemplate.doc_columns.length + 1;
    const { error } = await supabaseClient.from('doc_columns').insert([{
        template_id: currentTemplate.id,
        column_name: colName,
        column_type: colType,
        display_order: newOrder
    }]);

    btn.innerText = originalText;
    btn.disabled = false;

    if (!error) {
        delete categoryCache[currentTemplate.name]; // Clear cache to force refetch
        colNameInput.value = "";
        await switchCategory(currentTemplate.name); // Reload current view
    } else {
        alert("Error adding column: " + error.message);
    }
};
/**
 * SWITCH CATEGORY: The core logic for the SAP-style UI.
 * Handles card highlighting, workspace visibility, and data fetching.
 */
window.switchCategory = async function(name) {
    // 1. UI UPDATE: Highlight the selected card
    document.querySelectorAll('.category-card').forEach(card => card.classList.remove('active'));
    const activeCard = document.getElementById(`card-${name}`);
    if (activeCard) {
        activeCard.classList.add('active');
    }

    // 2. SAP-STYLE REVEAL: Show the hidden workspace
    const workspace = document.getElementById('moduleWorkspace');
    if (workspace) {
        workspace.style.display = 'block';
    }

    // 3. Update UI Labels
    const activeLabel = document.getElementById('activeCategoryName');
    if (activeLabel) {
        activeLabel.innerText = name;
    }

    // 4. Loading State
    const tableBody = document.getElementById('tableData');
    tableBody.innerHTML = '<tr><td colspan="100%" style="text-align:center; padding:20px;">Loading module data...</td></tr>';

    // 5. CACHE CHECK: Load instantly if already visited
    if (categoryCache[name]) {
        currentTemplate = categoryCache[name].template;
        localEntries = categoryCache[name].entries;
        
        // CRITICAL: Sync with global window for stress testing
        window.currentTemplate = currentTemplate; 
        
        renderUI();    
        renderTable(localEntries); 
        return; 
    }

    // 6. NETWORK FETCH: Get Template and Columns
    const { data: template, error: tErr } = await supabaseClient
        .from('doc_templates')
        .select('id, name, doc_columns(*)')
        .eq('name', name)
        .single();

    if (tErr) {
        console.error("Template Fetch Error:", tErr);
        tableBody.innerHTML = '<tr><td colspan="100%">Error loading category.</td></tr>';
        return;
    }

    // 7. NETWORK FETCH: Get Entries
    const { data: entries, error: eErr } = await supabaseClient
        .from('doc_entries')
        .select('*')
        .eq('template_id', template.id)
        .order('created_at', { ascending: false });

    if (eErr) console.error("Entries Fetch Error:", eErr);

    // 8. DATA PROCESSING
    template.doc_columns.sort((a, b) => a.display_order - b.display_order);
    
    // Save to Cache
    categoryCache[name] = { 
        template: template, 
        entries: entries || [] 
    };

    // 9. UPDATE LOCAL & GLOBAL STATE
    currentTemplate = template;
    localEntries = entries || [];

    // CRITICAL: Sync with global window so Console scripts can see it
    window.currentTemplate = currentTemplate;

    // 10. FINAL RENDER
    renderUI();    
    renderTable(localEntries); 
};

// ==========================================
// 4. RENDERING ENGINE
// ==========================================
function renderUI() {
    const form = document.getElementById('dynamicForm');
    form.innerHTML = currentTemplate.doc_columns.map(c => `
        <div class="input-box">
            <label>${c.column_name}</label>
            <input type="${c.column_type}" id="input_${c.column_name}" placeholder="Enter ${c.column_name}">
        </div>
    `).join('') + `<button onclick="saveData()" class="save-btn">Save Entry</button>`;

    const headers = document.getElementById('tableHeaders');
    headers.innerHTML = `<tr>${currentTemplate.doc_columns.map(c => `<th>${c.column_name}</th>`).join('')}<th>Action</th></tr>`;
}

function renderTable(entries) {
    const body = document.getElementById('tableData');
    if (!entries || entries.length === 0) {
        body.innerHTML = `<tr><td colspan="100%" style="text-align:center; padding: 20px;">No records found for this category.</td></tr>`;
        return;
    }

    body.innerHTML = entries.map(e => {
        const cells = currentTemplate.doc_columns.map(c => `<td>${e.content[c.column_name] || '-'}</td>`).join('');
        return `<tr>${cells}<td><button class="del-btn" onclick="deleteEntry('${e.id}')">Delete</button></td></tr>`;
    }).join('');
}

// ==========================================
// 5. DATA CRUD & EXPORT ACTIONS
// ==========================================
window.saveData = async function() {
    const content = {};
    let hasData = false;

    // OPTIMIZATION: Check if the user actually typed anything
    currentTemplate.doc_columns.forEach(c => {
        const val = document.getElementById(`input_${c.column_name}`).value.trim();
        content[c.column_name] = val;
        if (val) hasData = true;
    });

    if (!hasData) return alert("Please fill in at least one field.");

    // OPTIMIZATION: Loading state for save button
    const btn = document.querySelector('.save-btn');
    if (btn) {
        btn.innerText = "Saving...";
        btn.disabled = true;
    }

    const { data, error } = await supabaseClient.from('doc_entries').insert([{
        template_id: currentTemplate.id,
        content: content
    }]).select();

    if (btn) {
        btn.innerText = "Save Entry";
        btn.disabled = false;
    }

    if (!error && data) {
        localEntries.unshift(data[0]); // Add to top of local array
        categoryCache[currentTemplate.name].entries = localEntries; // Update cache
        renderTable(localEntries); // Render fast
        
        // Clear inputs
        currentTemplate.doc_columns.forEach(c => document.getElementById(`input_${c.column_name}`).value = "");
    } else {
        alert("Error saving data: " + (error?.message || "Unknown error"));
    }
};

window.deleteEntry = async function(id) {
    if(!confirm("Are you sure you want to delete this record?")) return;

    const { error } = await supabaseClient.from('doc_entries').delete().eq('id', id);
    if (!error) {
        localEntries = localEntries.filter(e => e.id !== id);
        categoryCache[currentTemplate.name].entries = localEntries;
        renderTable(localEntries);
    } else {
        alert("Error deleting record.");
    }
};

window.searchData = function() {
    const term = document.getElementById('search').value.toLowerCase();
    const filtered = localEntries.filter(e => 
        JSON.stringify(e.content).toLowerCase().includes(term)
    );
    renderTable(filtered);
};

// OPTIMIZATION: Handle missing/empty dates properly during sort
window.sortByDate = function () {
    if (!currentTemplate) return;

    const dateColumn = currentTemplate.doc_columns.find(c => c.column_type === 'date');
    if (!dateColumn) return alert("No date column found in this category.");

    const colName = dateColumn.column_name;

    localEntries.sort((a, b) => {
        // Fallback to year 1970 if date is missing to push them to the bottom/top predictably
        const dateA = a.content[colName] ? new Date(a.content[colName]) : new Date(0);
        const dateB = b.content[colName] ? new Date(b.content[colName]) : new Date(0);

        return dateSortAsc ? dateA - dateB : dateB - dateA;
    });

    dateSortAsc = !dateSortAsc;
    categoryCache[currentTemplate.name].entries = localEntries;
    renderTable(localEntries);
};

// ADDED: Missing Export function to support your HTML button
window.exportToExcel = function() {
    if (!localEntries || localEntries.length === 0) {
        return alert("No data to export for this category.");
    }
    
    // Flatten JSONB content for Excel
    const exportData = localEntries.map(e => e.content);
    
    const ws = XLSX.utils.json_to_sheet(exportData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, currentTemplate.name);
    
    // Generate file
    XLSX.writeFile(wb, `${currentTemplate.name}_Records.xlsx`);
};