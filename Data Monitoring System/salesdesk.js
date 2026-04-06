import { SUPABASE_CONFIG } from './config.js';

const { createClient } = supabase;
const supabaseClient = createClient(SUPABASE_CONFIG.url, SUPABASE_CONFIG.anonKey);

// --- GLOBAL STATE & CACHE ---
let currentTemplate = null;
let allTemplates = [];
let localEntries = [];
let dateSortAsc = true;
const categoryCache = {}; // Stores { template, entries } for every visited module

window.onload = async function() {
    await refreshCategories();
};

/** * 1. INITIAL FETCH: Get all available categories (Modules)
 */
async function refreshCategories() {
    const { data, error } = await supabaseClient.from('doc_templates').select('*');
    if (error) return console.error("Error fetching categories:", error);
    
    allTemplates = data || [];
    const dropdown = document.getElementById('categoryDropdown');
    dropdown.innerHTML = allTemplates.map(t => `<option value="${t.name}">${t.name}</option>`).join('');
    
    // Auto-load the first category on start
    if (allTemplates.length > 0 && !currentTemplate) {
        switchCategory(allTemplates[0].name);
    }
}

/** * 2. SMART SWITCHING: The core of the "No-Lag" system
 */
window.switchCategory = async function(name) {
    // Show loading state in the table only
    document.getElementById('tableData').innerHTML = '<tr><td colspan="100%">Loading data...</td></tr>';

    // CHECK CACHE FIRST
    if (categoryCache[name]) {
        console.log(`%c Loading ${name} from Cache (Instant)`, "color: green; font-weight: bold;");
        currentTemplate = categoryCache[name].template;
        localEntries = categoryCache[name].entries;
        renderUI();
        renderTable(localEntries);
        return;
    }

    // IF NOT IN CACHE: Fetch from network
    console.log(`%c Fetching ${name} from Database (Network)`, "color: orange; font-weight: bold;");
    const { data: template, error: tErr } = await supabaseClient
        .from('doc_templates')
        .select('id, name, doc_columns(*)')
        .eq('name', name)
        .single();

    if (tErr) return console.error("Template Error:", tErr);

    const { data: entries, error: eErr } = await supabaseClient
        .from('doc_entries')
        .select('*')
        .eq('template_id', template.id)
        .order('created_at', { ascending: false });

    // Store in Cache
    template.doc_columns.sort((a, b) => a.display_order - b.display_order);
    categoryCache[name] = { template, entries: entries || [] };

    // Set Global State
    currentTemplate = template;
    localEntries = entries || [];

    document.getElementById('activeCategoryName').innerText = name;
    renderUI();
    renderTable(localEntries);
};

/** * 3. RENDERING ENGINE: Builds the Form and Headers
 */
function renderUI() {
    // Build Form Inputs
    const form = document.getElementById('dynamicForm');
    form.innerHTML = currentTemplate.doc_columns.map(c => `
        <div class="input-box">
            <label>${c.column_name}</label>
            <input type="${c.column_type}" id="input_${c.column_name}" placeholder="Enter ${c.column_name}">
        </div>
    `).join('') + `<button onclick="saveData()" class="save-btn">Save Entry</button>`;

    // Build Table Headers
    const headers = document.getElementById('tableHeaders');
    headers.innerHTML = `<tr>${currentTemplate.doc_columns.map(c => `<th>${c.column_name}</th>`).join('')}<th>Action</th></tr>`;
}

/** * 4. TABLE RENDERER: Fast DOM injection
 */
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

/** * 5. DATA ACTIONS: Save, Delete, and Update Cache
 */
window.saveData = async function() {
    const content = {};
    currentTemplate.doc_columns.forEach(c => {
        const val = document.getElementById(`input_${c.column_name}`).value;
        content[c.column_name] = val;
    });

    const { data, error } = await supabaseClient.from('doc_entries').insert([{
        template_id: currentTemplate.id,
        content: content
    }]).select();

    if (!error) {
        // Update Local State & Cache so we don't have to refetch from DB
        localEntries.unshift(data[0]);
        categoryCache[currentTemplate.name].entries = localEntries;
        
        renderTable(localEntries);
        // Clear inputs
        currentTemplate.doc_columns.forEach(c => document.getElementById(`input_${c.column_name}`).value = "");
    }
};

window.deleteEntry = async function(id) {
    if(!confirm("Are you sure you want to delete this record?")) return;

    const { error } = await supabaseClient.from('doc_entries').delete().eq('id', id);
    if (!error) {
        // Update Local State & Cache
        localEntries = localEntries.filter(e => e.id !== id);
        categoryCache[currentTemplate.name].entries = localEntries;
        renderTable(localEntries);
    }
};

/** * 6. SEARCH: Instance Search in Memory (No Lag)
 */
window.searchData = function() {
    const term = document.getElementById('search').value.toLowerCase();
    // We search the 'localEntries' variable which is already in RAM
    const filtered = localEntries.filter(e => 
        JSON.stringify(e.content).toLowerCase().includes(term)
    );
    renderTable(filtered);
};

/** * 7. ADMIN: Add New Category & New Column
 */
window.createNewCategory = async function() {
    const name = document.getElementById('newCategoryName').value;
    if (!name) return;

    const { error } = await supabaseClient.from('doc_templates').insert([{ name }]);
    if (!error) {
        document.getElementById('newCategoryName').value = "";
        await refreshCategories();
    }
};

window.addColumnToActive = async function() {
    const colName = document.getElementById('newColumnName').value;
    const colType = document.getElementById('newColumnType').value;
    if (!colName) return;

    const newOrder = currentTemplate.doc_columns.length + 1;
    const { error } = await supabaseClient.from('doc_columns').insert([{
        template_id: currentTemplate.id,
        column_name: colName,
        column_type: colType,
        display_order: newOrder
    }]);

    if (!error) {
        // Clear Cache for this category so it refetches the new layout
        delete categoryCache[currentTemplate.name];
        document.getElementById('newColumnName').value = "";
        await switchCategory(currentTemplate.name); 
    }
};

/** * 8. Sort by date function
 */
window.sortByDate = function () {
    if (!currentTemplate) return;

    // hanapin yung date column
    const dateColumn = currentTemplate.doc_columns.find(c => c.column_type === 'date');

    if (!dateColumn) {
        alert("No date column found in this category.");
        return;
    }

    const colName = dateColumn.column_name;

    localEntries.sort((a, b) => {
        const dateA = new Date(a.content[colName] || 0);
        const dateB = new Date(b.content[colName] || 0);

        return dateSortAsc ? dateA - dateB : dateB - dateA;
    });

    // toggle asc/desc
    dateSortAsc = !dateSortAsc;

    // update cache
    categoryCache[currentTemplate.name].entries = localEntries;

    renderTable(localEntries);
};