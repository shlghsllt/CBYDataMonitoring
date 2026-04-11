import { SUPABASE_CONFIG } from './config.js';

const { createClient } = supabase;
const supabaseClient = createClient(SUPABASE_CONFIG.url, SUPABASE_CONFIG.anonKey);

export const AppCore = {
    state: {
        moduleName: '',
        currentTemplate: null,
        allTemplates: [],
        localEntries: [],
        dateSortAsc: true,
        editingId: null,
        cache: {},
        isLoading: false,
        _importWorkbook: null,
        _importColumnsWorkbook: null,
        _columnsToImport: []
    },

    // ============================================================
    // INIT
    // ============================================================
    init: async function(moduleName) {
        this.state.moduleName = moduleName;
        this.syncWithWindow();
        await this.refreshCategories();
    },

    syncWithWindow: function() {
        window.switchCategory    = (name)     => this.switchCategory(name);
        window.saveData          = ()          => this.saveData();
        window.editEntry         = (id)        => this.editEntry(id);
        window.deleteEntry       = (id)        => this.deleteEntry(id);
        window.searchData        = ()          => this.searchData();
        window.sortByDate        = ()          => this.sortByDate();
        window.exportToExcel     = ()          => this.exportToExcel();

        window.openModal         = ()          => document.getElementById('categoryModal').style.display = 'block';
        window.closeModal        = ()          => document.getElementById('categoryModal').style.display = 'none';
        window.openColumnModal   = ()          => this.openColumnModal();
        window.closeColumnModal  = ()          => document.getElementById('columnModal').style.display = 'none';

        window.createNewCategory = ()          => this.createNewCategory();
        window.addColumnToActive = ()          => this.addColumnToActive();
        window.deleteColumn      = (id, name)  => this.deleteColumn(id, name);
        window.deleteCategory    = (id, name)  => this.deleteCategory(id, name);
        window.toggleMenu        = (event, id) => this.toggleMenu(event, `menu-${id}`);

        window.openImportModal   = ()          => document.getElementById('importModal').style.display = 'block';
        window.closeImportModal  = ()          => this.closeImportModal();
        window.loadSheets        = ()          => this.loadSheets();
        window.confirmImport     = ()          => this.confirmImport();

        window.openImportColumnsModal = ()     => this.openImportColumnsModal();
        window.closeImportColumnsModal = ()    => this.closeImportColumnsModal();
        window.loadColumnsSheets = ()          => this.loadColumnsSheets();
        window.confirmImportColumns = ()       => this.confirmImportColumns();

        window.addEventListener('click', () => {
            document.querySelectorAll('.dropdown').forEach(d => d.style.display = 'none');
        });
    },

    // ============================================================
    // TOAST
    // ============================================================
    showToast: function(message, type = 'success') {
        const toast = document.createElement('div');
        toast.className = `toast toast-${type} show`;
        toast.innerText = message;
        document.body.appendChild(toast);
        setTimeout(() => {
            toast.classList.remove('show');
            setTimeout(() => toast.remove(), 500);
        }, 3000);
    },

    // ============================================================
    // CATEGORY SWITCHING
    // ============================================================
    switchCategory: async function(name) {
        if (this.state.isLoading) return;

        this.updateActiveUI(name);
        const workspace = document.getElementById('moduleWorkspace');
        workspace.style.display      = 'block';
        workspace.style.opacity      = '0.4';
        workspace.style.pointerEvents = 'none';
        this.state.isLoading = true;

        requestAnimationFrame(async () => {
            try {
                if (this.state.cache[name]) {
                    this.state.currentTemplate = this.state.cache[name].template;
                    this.state.localEntries    = this.state.cache[name].entries;
                } else {
                    const templateId = this.state.allTemplates.find(t => t.name === name)?.id;

                    const [tRes, eRes] = await Promise.all([
                        supabaseClient.from('doc_templates').select('*, doc_columns(*)').eq('id', templateId).single(),
supabaseClient.from('doc_entries').select('*').eq('template_id', templateId).order('id', { ascending: true })                    ]);

                    if (tRes.error) throw tRes.error;

                    const template = tRes.data;
                    template.doc_columns.sort((a, b) => a.display_order - b.display_order);

                    this.state.currentTemplate = template;
                    this.state.localEntries    = eRes.data || [];
                    this.state.cache[name]     = { template, entries: this.state.localEntries };
                }

                this.renderAll();

            } catch (err) {
                this.showToast('Switch failed: ' + err.message, 'error');
            } finally {
                workspace.style.opacity       = '1';
                workspace.style.pointerEvents = 'auto';
                this.state.isLoading          = false;
            }
        });
    },

    updateActiveUI: function(name) {
        document.querySelectorAll('.category-card').forEach(c => c.classList.remove('active'));
        const activeCard = document.getElementById(`card-${name}`);
        if (activeCard) activeCard.classList.add('active');

        const label = document.getElementById('activeCategoryName');
        if (label) label.innerText = name;
    },

    // ============================================================
    // RENDER
    // ============================================================
    renderAll: function() {
        const form = document.getElementById('dynamicForm');
        if (!form) return;

        form.innerHTML = this.state.currentTemplate.doc_columns.map(c => `
            <div class="input-box">
                <label>${c.column_name}</label>
                <input type="${c.column_type}" id="input_${c.column_name}" placeholder="Enter ${c.column_name}">
            </div>
        `).join('') + `<button onclick="saveData()" class="save-btn" id="mainSaveBtn">Save Record</button>`;

        const headers = document.getElementById('tableHeaders');
        headers.innerHTML = `<tr>
            ${this.state.currentTemplate.doc_columns.map(c => `<th>${c.column_name}</th>`).join('')}
            <th>Actions</th>
        </tr>`;

        this.renderTable(this.state.localEntries);
        this.setupTableEditing();
    },

    renderTable: function(entries) {
        const body = document.getElementById('tableData');
        if (!body) return;

        body.innerHTML = entries.length
            ? entries.map(e => `
                <tr data-entry-id="${e.id}">
                    ${this.state.currentTemplate.doc_columns.map(c => `<td contenteditable="true" data-col-name="${c.column_name}">${e.content[c.column_name] || ''}</td>`).join('')}
                    <td class="action-buttons">
                        <button class="edit-btn" onclick="editEntry('${e.id}')">Edit</button>
                        <button class="del-btn" onclick="deleteEntry('${e.id}')">Delete</button>
                    </td>
                </tr>
            `).join('')
            : '<tr><td colspan="100%" style="text-align:center;padding:40px;color:#94a3b8;">No records found.</td></tr>';
    },

    setupTableEditing: function() {
        if (this.tableEventsInitialized) return;
        const body = document.getElementById('tableData');
        if (!body) return;

        body.addEventListener('focusin', (e) => {
            const td = e.target;
            if (td.tagName !== 'TD' || !td.isContentEditable) return;
            td.classList.add('cell-focused');
        });

        body.addEventListener('focusout', (e) => {
            const td = e.target;
            if (td.tagName !== 'TD' || !td.isContentEditable) return;
            td.classList.remove('cell-focused');
            this.onTableCellBlur(td);
        });

        body.addEventListener('keydown', (e) => this.onTableCellKeyDown(e));
        body.addEventListener('paste', (e) => this.onTableCellPaste(e));

        this.tableEventsInitialized = true;
    },

    parseTabular: function(text) {
        return text.replace(/\r/g, '').split('\n').filter(line => line !== '').map(line => line.split('\t'));
    },

    getTableCellInfo: function(td) {
        if (!td || td.tagName !== 'TD') return null;
        const row = td.closest('tr');
        const entryId = row?.dataset.entryId;
        const colName = td.dataset.colName;
        return entryId && colName ? { entryId, colName } : null;
    },

    onTableCellBlur: function(td) {
        const info = this.getTableCellInfo(td);
        if (!info) return;
        const entry = this.state.localEntries.find(e => e.id === info.entryId);
        if (!entry) return;

        const newValue = td.textContent.trim();
        const currentValue = entry.content[info.colName] || '';
        if (newValue === currentValue) return;

        entry.content[info.colName] = newValue;
        this.saveEntryField(entry.id, entry.content);
    },

    onTableCellKeyDown: function(e) {
        const td = e.target;
        if (td.tagName !== 'TD' || !td.isContentEditable) return;
        if (e.key === 'Enter') {
            e.preventDefault();
            td.blur();
        }
    },

    onTableCellPaste: function(e) {
        const td = e.target;
        if (td.tagName !== 'TD' || !td.isContentEditable) return;

        const text = e.clipboardData?.getData('text/plain') || '';
        if (!text) return;
        e.preventDefault();

        const pasted = this.parseTabular(text);
        if (!pasted.length) return;

        const tbody = td.closest('tbody');
        const rows = Array.from(tbody.rows);
        const startRowIndex = rows.indexOf(td.closest('tr'));
        const columns = this.state.currentTemplate.doc_columns.map(c => c.column_name);
        const startColIndex = columns.indexOf(td.dataset.colName);
        const changedEntries = new Map();

        pasted.forEach((rowValues, rowOffset) => {
            const targetRow = rows[startRowIndex + rowOffset];
            if (!targetRow) return;
            const entryId = targetRow.dataset.entryId;
            const entry = this.state.localEntries.find(e => e.id === entryId);
            if (!entry) return;

            rowValues.forEach((cellValue, colOffset) => {
                const colName = columns[startColIndex + colOffset];
                if (!colName) return;
                const cell = targetRow.querySelector(`td[data-col-name="${colName}"]`);
                if (!cell) return;

                const normalized = cellValue.trim();
                if ((entry.content[colName] || '') === normalized) return;

                entry.content[colName] = normalized;
                cell.textContent = normalized;
                changedEntries.set(entryId, entry);
            });
        });

        if (!changedEntries.size) return;
        changedEntries.forEach((entry) => this.saveEntryField(entry.id, entry.content));
    },

    saveEntryField: async function(entryId, content) {
        try {
            const { error } = await supabaseClient
                .from('doc_entries')
                .update({ content })
                .eq('id', entryId);
            if (error) throw error;

            if (this.state.cache[this.state.currentTemplate.name]) {
                this.state.cache[this.state.currentTemplate.name].entries = this.state.localEntries;
            }
        } catch (err) {
            this.showToast('Save failed: ' + err.message, 'error');
        }
    },

    // ============================================================
    // CATEGORIES
    // ============================================================
    refreshCategories: async function() {
        const { data, error } = await supabaseClient
            .from('doc_templates')
            .select('*')
            .eq('module', this.state.moduleName);
        if (error) return;
        this.state.allTemplates = data || [];
        this.renderCategoryCards();
    },

    renderCategoryCards: function() {
        const container = document.getElementById('categoryCards');
        if (!container) return;
        container.innerHTML = this.state.allTemplates.map(t => {
            const color = `hsla(${Math.abs(t.name.split('').reduce((a, b) => (((a << 5) - a) + b.charCodeAt(0)) | 0, 0)) % 360}, 70%, 85%, 1)`;
            return `
            <div class="category-card" id="card-${t.name}" style="background-color:${color};" onclick="switchCategory('${t.name}')">
                <div class="card-menu" onclick="event.stopPropagation()">
                    <button class="menu-btn" onclick="toggleMenu(event, '${t.id}')">⋮</button>
                    <div class="dropdown" id="menu-${t.id}">
                        <button onclick="deleteCategory('${t.id}', '${t.name}')">🗑️ Delete</button>
                    </div>
                </div>
                <div class="card-icon">${t.name.substring(0, 2).toUpperCase()}</div>
                <span class="card-label">${t.name}</span>
            </div>`;
        }).join('');
    },

    createNewCategory: async function() {
        const name = document.getElementById('newCategoryName').value.trim();
        if (!name) return this.showToast('Category name is required.', 'error');

        try {
            const { data, error } = await supabaseClient
                .from('doc_templates')
                .insert([{ name, module: this.state.moduleName }])
                .select();
            if (error) throw error;

            this.state.allTemplates.push(data[0]);
            this.renderCategoryCards();
            this.showToast('Category created!');
            document.getElementById('newCategoryName').value = '';
            window.closeModal();
        } catch (err) {
            this.showToast('Failed: ' + err.message, 'error');
        }
    },

    deleteCategory: async function(id, name) {
        if (!confirm(`Delete "${name}"? All data will be lost.`)) return;
        try {
            const { error } = await supabaseClient.from('doc_templates').delete().eq('id', id);
            if (error) throw error;

            this.state.allTemplates = this.state.allTemplates.filter(t => t.id !== id);
            delete this.state.cache[name];

            if (this.state.currentTemplate?.id === id) {
                this.state.currentTemplate = null;
                document.getElementById('moduleWorkspace').style.display = 'none';
            }

            this.renderCategoryCards();
            this.showToast('Category deleted.');
        } catch (err) {
            this.showToast('Failed: ' + err.message, 'error');
        }
    },

    // ============================================================
    // COLUMNS
    // ============================================================
    addColumnToActive: async function() {
        const name = document.getElementById('newColumnName').value.trim();
        const type = document.getElementById('newColumnType').value;

        if (!name) return this.showToast('Column name is required.', 'error');
        if (!this.state.currentTemplate) return this.showToast('No category selected.', 'error');

        const order = this.state.currentTemplate.doc_columns.length;

        try {
            const { data, error } = await supabaseClient
                .from('doc_columns')
                .insert([{
                    template_id:   this.state.currentTemplate.id,
                    column_name:   name,
                    column_type:   type,
                    display_order: order
                }])
                .select();
            if (error) throw error;

            this.state.currentTemplate.doc_columns.push(data[0]);
            delete this.state.cache[this.state.currentTemplate.name];

            this.renderAll();
            this.showToast('Column added!');
            document.getElementById('newColumnName').value = '';
            window.closeColumnModal();
        } catch (err) {
            this.showToast('Failed to add column: ' + err.message, 'error');
        }
    },

    deleteColumn: async function(id, name) {
        if (!confirm(`Delete column "${name}"? This will affect all records.`)) return;
        try {
            const { error } = await supabaseClient.from('doc_columns').delete().eq('id', id);
            if (error) throw error;

            this.state.currentTemplate.doc_columns = this.state.currentTemplate.doc_columns.filter(c => c.id !== id);
            delete this.state.cache[this.state.currentTemplate.name];

            this.renderAll();
            this.showToast('Column deleted.');
        } catch (err) {
            this.showToast('Failed: ' + err.message, 'error');
        }
    },

    // ============================================================
    // ENTRIES — SAVE / EDIT / DELETE
    // ============================================================
    saveData: async function() {
        const content = {};
        this.state.currentTemplate.doc_columns.forEach(c => {
            content[c.column_name] = document.getElementById(`input_${c.column_name}`).value;
        });

        const btn = document.getElementById('mainSaveBtn');
        btn.disabled = true;

        try {
            let res;
            if (this.state.editingId) {
                res = await supabaseClient
                    .from('doc_entries')
                    .update({ content })
                    .eq('id', this.state.editingId)
                    .select();
                const idx = this.state.localEntries.findIndex(e => e.id === this.state.editingId);
                this.state.localEntries[idx] = res.data[0];
            } else {
                res = await supabaseClient
                    .from('doc_entries')
                    .insert([{ template_id: this.state.currentTemplate.id, content }])
                    .select();
                this.state.localEntries.push(res.data[0]);
                this.state.localEntries.sort((a, b) => a.id - b.id);
            }

            this.state.cache[this.state.currentTemplate.name].entries = this.state.localEntries;
            this.state.editingId = null;
            btn.innerText = 'Save Record';
            this.renderTable(this.state.localEntries);
            this.state.currentTemplate.doc_columns.forEach(c => {
                document.getElementById(`input_${c.column_name}`).value = '';
            });
            this.showToast('Saved successfully!');
        } catch (err) {
            this.showToast(err.message, 'error');
        } finally {
            btn.disabled = false;
        }
    },

    editEntry: function(id) {
        const entry = this.state.localEntries.find(e => e.id === id);
        if (!entry) return;

        this.state.currentTemplate.doc_columns.forEach(c => {
            const input = document.getElementById(`input_${c.column_name}`);
            if (input) input.value = entry.content[c.column_name] || '';
        });

        this.state.editingId = id;
        const btn = document.getElementById('mainSaveBtn');
        if (btn) btn.innerText = 'Update Record';

        document.getElementById('dynamicForm')?.scrollIntoView({ behavior: 'smooth', block: 'start' });
    },

    deleteEntry: async function(id) {
        if (!confirm('Are you sure you want to delete this record?')) return;
        try {
            const { error } = await supabaseClient.from('doc_entries').delete().eq('id', id);
            if (error) throw error;

            this.state.localEntries = this.state.localEntries.filter(e => e.id !== id);
            this.state.cache[this.state.currentTemplate.name].entries = this.state.localEntries;

            this.renderTable(this.state.localEntries);
            this.showToast('Record deleted.');
        } catch (err) {
            this.showToast('Delete failed: ' + err.message, 'error');
        }
    },

    // ============================================================
    // SEARCH & SORT
    // ============================================================
    searchData: function() {
        const term = document.getElementById('search').value.toLowerCase();
        const filtered = this.state.localEntries.filter(e =>
            JSON.stringify(e.content).toLowerCase().includes(term)
        );
        this.renderTable(filtered);
    },

    sortByDate: function() {
        const dateCol = this.state.currentTemplate.doc_columns.find(c => c.column_type === 'date');
        if (!dateCol) return this.showToast('No date column found.', 'error');

        this.state.localEntries.sort((a, b) => {
            const d1 = new Date(a.content[dateCol.column_name] || 0);
            const d2 = new Date(b.content[dateCol.column_name] || 0);
            return this.state.dateSortAsc ? d1 - d2 : d2 - d1;
        });

        this.state.dateSortAsc = !this.state.dateSortAsc;
        this.renderTable(this.state.localEntries);
    },

    // ============================================================
    // EXPORT
    // ============================================================
    exportToExcel: function() {
        if (!this.state.currentTemplate) return;
        const formatted = this.state.localEntries.map(e => e.content);
        const ws = XLSX.utils.json_to_sheet(formatted);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, this.state.currentTemplate.name);
        XLSX.writeFile(wb, `${this.state.currentTemplate.name}.xlsx`);
    },

    // ============================================================
    // IMPORT FROM EXCEL
    // ============================================================
    openImportModal: function() {
        if (!this.state.currentTemplate)
            return this.showToast('Select a category first.', 'error');
        document.getElementById('importModal').style.display = 'block';
    },

    closeImportModal: function() {
        document.getElementById('importModal').style.display = 'none';
        document.getElementById('importFile').value = '';
        document.getElementById('importSheet').innerHTML = '<option>— load a file first —</option>';
        document.getElementById('importSheet').disabled = true;
        document.getElementById('importConfirmBtn').disabled = true;
        document.getElementById('importPreview').innerHTML = '';
        this.state._importWorkbook = null;
    },
    loadSheets: function() {
        const file = document.getElementById('importFile').files[0];
        if (!file) return;

        const reader = new FileReader();
        reader.onload = (e) => {
            const workbook = XLSX.read(e.target.result, { type: 'array' });
            this.state._importWorkbook = workbook;

            const sheetSelect = document.getElementById('importSheet');
            sheetSelect.innerHTML = workbook.SheetNames.map(
                name => `<option value="${name}">${name}</option>`
            ).join('');
            sheetSelect.disabled = false;
            sheetSelect.onchange = () => this.previewSheet();

            this.previewSheet();
        };
        reader.readAsArrayBuffer(file);
    },

    // ============================================================
// IMPORT FROM EXCEL - FIXED
// ============================================================
getImportRawRows: function(ws) {
    // Get all rows as 2D array
    const rawRows = XLSX.utils.sheet_to_json(ws, { defval: '', header: 1 });
    // Filter out completely empty rows AND rows that are just formulas with no values
    return rawRows.filter(row => {
        // Check if row has any actual content (not just empty strings)
        return row.some(cell => {
            const str = String(cell).trim();
            return str !== '' && !str.startsWith('=');
        });
    });
},

detectHeaderRow: function(rawRows, cols) {
    if (!rawRows.length) return false;
    
    // Look at first few rows to find the actual header
    for (let i = 0; i < Math.min(5, rawRows.length); i++) {
        const row = rawRows[i].map(cell => String(cell).trim().toUpperCase());
        const matchCount = cols.filter(col => row.includes(col.toUpperCase())).length;
        const threshold = Math.max(2, Math.ceil(cols.length / 2));
        if (matchCount >= threshold) {
            return i; // Return the index of the header row
        }
    }
    return -1; // No header found
},

mapImportRows: function(rawRows, cols) {
    const headerIndex = this.detectHeaderRow(rawRows, cols);
    
    let dataStartIndex;
    let headers;
    
    if (headerIndex >= 0) {
        // Use the detected header row
        headers = rawRows[headerIndex].map(cell => String(cell).trim());
        dataStartIndex = headerIndex + 1;
    } else {
        // No header found - use column order
        headers = cols;
        dataStartIndex = 0;
    }
    
    // Map rows starting from dataStartIndex
    const mappedRows = [];
    for (let i = dataStartIndex; i < rawRows.length; i++) {
        const row = rawRows[i];
        // Skip empty rows
        if (!row.some(cell => String(cell).trim() !== '')) continue;
        
        const mappedRow = {};
        cols.forEach(col => {
            if (headerIndex >= 0) {
                // Find by header name (case-insensitive)
                const colIndex = headers.findIndex(h => 
                    h.toUpperCase() === col.toUpperCase()
                );
                if (colIndex >= 0 && colIndex < row.length) {
                    const value = String(row[colIndex] || '').trim();
                    // Skip formula results that are empty
                    mappedRow[col] = value;
                } else {
                    mappedRow[col] = '';
                }
            } else {
                // Map by position
                const colPosition = cols.indexOf(col);
                if (colPosition < row.length) {
                    const value = String(row[colPosition] || '').trim();
                    mappedRow[col] = value;
                } else {
                    mappedRow[col] = '';
                }
            }
        });
        
        // Only add if at least one column has a value
        if (Object.values(mappedRow).some(v => v !== '')) {
            mappedRows.push(mappedRow);
        }
    }
    
    return mappedRows;
},

previewSheet: function() {
    const sheetName = document.getElementById('importSheet').value;
    const ws = this.state._importWorkbook.Sheets[sheetName];
    const rawRows = this.getImportRawRows(ws);

    const preview = document.getElementById('importPreview');
    const confirmBtn = document.getElementById('importConfirmBtn');
    const cols = this.state.currentTemplate.doc_columns.map(c => c.column_name);
    
    const previewRows = this.mapImportRows(rawRows, cols);

    if (!previewRows.length) {
        preview.innerHTML = '<span style="color:#ef4444;">No data rows found in this sheet. Make sure the sheet contains data matching your columns.</span>';
        confirmBtn.disabled = true;
        return;
    }

    const headerIndex = this.detectHeaderRow(rawRows, cols);
    const headerNote = headerIndex >= 0
        ? `Found header row at position ${headerIndex + 1}. Mapping by column names.`
        : 'No header row found. Importing rows by column order.';

    // Show sample of first 3 rows in preview
    const sampleRows = previewRows.slice(0, 3).map(row => 
        Object.entries(row).map(([k, v]) => `${k}: ${v || '(empty)'}`).join(', ')
    ).join('<br>');

    preview.innerHTML = `
        <div style="font-size:12px;line-height:1.8;padding:10px 12px;background:#f8fafc;border-radius:7px;border:1px solid #e2e8f0;">
            <div><strong>${previewRows.length} rows</strong> will be imported</div>
            <div>${headerNote}</div>
            <div>Columns: <span style="color:#16a34a;font-weight:600;">${cols.join(', ')}</span></div>
            <div style="margin-top:8px;padding-top:8px;border-top:1px solid #e2e8f0;">
                <strong>Sample data:</strong><br>
                <span style="font-family:monospace;font-size:11px;">${sampleRows}</span>
            </div>
        </div>
    `;
    confirmBtn.disabled = false;
},

confirmImport: async function() {
    const sheetName = document.getElementById('importSheet').value;
    const ws = this.state._importWorkbook.Sheets[sheetName];
    const rawRows = this.getImportRawRows(ws);
    const cols = this.state.currentTemplate.doc_columns.map(c => c.column_name);
    const rows = this.mapImportRows(rawRows, cols);

    if (!rows.length) {
        return this.showToast('No data rows to import.', 'error');
    }

    const entries = rows.map(row => ({
        template_id: this.state.currentTemplate.id,
        content: Object.fromEntries(cols.map(c => [c, row[c] || '']))
    }));

    const confirmBtn = document.getElementById('importConfirmBtn');
    confirmBtn.disabled = true;
    confirmBtn.innerText = 'Importing...';

    try {
        const batchSize = 100;
        for (let i = 0; i < entries.length; i += batchSize) {
            const batch = entries.slice(i, i + batchSize);
            const { error } = await supabaseClient.from('doc_entries').insert(batch);
            if (error) throw error;
        }

        // Refresh the entries
                const { data, error } = await supabaseClient
            .from('doc_entries')
            .select('*')
            .eq('template_id', this.state.currentTemplate.id)
            .order('id', { ascending: true });

        if (error) throw error;

        this.state.cache[this.state.currentTemplate.name] = {
        template: this.state.currentTemplate,
        entries: this.state.localEntries
        };
        if (this.state.cache[this.state.currentTemplate.name]) {
            this.state.cache[this.state.currentTemplate.name].entries = this.state.localEntries;
        }

        this.renderTable(this.state.localEntries);
        this.closeImportModal();
        this.showToast(`${entries.length} rows imported successfully!`);

    } catch (err) {
        this.showToast('Import failed: ' + err.message, 'error');
    } finally {
        confirmBtn.disabled = false;
        confirmBtn.innerText = 'Import';
    }
},

    // ============================================================
    // IMPORT COLUMNS FROM EXCEL
    // ============================================================
    addImportColumnsButton: function() {
        const columnModal = document.getElementById('columnModal');
        if (!columnModal) return;
        const actions = columnModal.querySelector('.modal-actions');
        if (!actions || actions.querySelector('.import-columns-btn')) return;

        const btn = document.createElement('button');
        btn.className = 'import-columns-btn';
        btn.onclick = () => this.openImportColumnsModal();
        btn.innerText = 'Import Columns';
        const addBtn = actions.querySelector('button[onclick*="addColumnToActive"]');
        if (addBtn) {
            actions.insertBefore(btn, addBtn);
        } else {
            actions.appendChild(btn);
        }
    },

    openImportColumnsModal: function() {
        if (!this.state.currentTemplate) return this.showToast('Select a category first.', 'error');

        let modal = document.getElementById('importColumnsModal');
        if (!modal) {
            modal = document.createElement('div');
            modal.id = 'importColumnsModal';
            modal.className = 'modal';
            modal.innerHTML = `
                <div class="modal-content">
                    <span class="close-btn" onclick="closeImportColumnsModal()">&times;</span>
                    <h3>Import Columns from Excel</h3>

                    <label class="modal-label">Choose File</label>
                    <input type="file" id="importColumnsFile" accept=".xlsx,.xls" onchange="loadColumnsSheets()">

                    <label class="modal-label">Select Sheet</label>
                    <select id="importColumnsSheet" disabled>
                        <option>— load a file first —</option>
                    </select>

                    <div id="importColumnsPreview"></div>

                    <div class="modal-actions">
                        <button onclick="closeImportColumnsModal()">Cancel</button>
                        <button id="importColumnsConfirmBtn" onclick="confirmImportColumns()" disabled>Import</button>
                    </div>
                </div>
            `;
            document.body.appendChild(modal);
        }
        modal.style.display = 'block';
    },

    closeImportColumnsModal: function() {
        const modal = document.getElementById('importColumnsModal');
        if (modal) modal.style.display = 'none';
        const fileInput = document.getElementById('importColumnsFile');
        if (fileInput) fileInput.value = '';
        const sheetSelect = document.getElementById('importColumnsSheet');
        if (sheetSelect) {
            sheetSelect.innerHTML = '<option>— load a file first —</option>';
            sheetSelect.disabled = true;
        }
        const confirmBtn = document.getElementById('importColumnsConfirmBtn');
        if (confirmBtn) {
            confirmBtn.disabled = true;
            confirmBtn.innerText = 'Import';
        }
        const preview = document.getElementById('importColumnsPreview');
        if (preview) preview.innerHTML = '';
        this.state._importColumnsWorkbook = null;
        this.state._columnsToImport = [];
    },

    loadColumnsSheets: function() {
        const fileInput = document.getElementById('importColumnsFile');
        const file = fileInput ? fileInput.files[0] : null;
        if (!file) return;

        const reader = new FileReader();
        reader.onload = (e) => {
            const workbook = XLSX.read(e.target.result, { type: 'array' });
            this.state._importColumnsWorkbook = workbook;

            const sheetSelect = document.getElementById('importColumnsSheet');
            if (sheetSelect) {
                sheetSelect.innerHTML = workbook.SheetNames.map(
                    name => `<option value="${name}">${name}</option>`
                ).join('');
                sheetSelect.disabled = false;
                sheetSelect.onchange = () => this.previewColumnsSheet();
            }

            this.previewColumnsSheet();
        };
        reader.readAsArrayBuffer(file);
    },

    previewColumnsSheet: function() {
        const sheetSelect = document.getElementById('importColumnsSheet');
        const sheetName = sheetSelect ? sheetSelect.value : '';
        if (!sheetName || !this.state._importColumnsWorkbook) return;

        const ws = this.state._importColumnsWorkbook.Sheets[sheetName];
        const rawRows = XLSX.utils.sheet_to_json(ws, { defval: '', header: 1 });

        const preview = document.getElementById('importColumnsPreview');
        const confirmBtn = document.getElementById('importColumnsConfirmBtn');

        if (!preview || !confirmBtn) return;

        if (rawRows.length < 2) {
            preview.innerHTML = '<span style="color:#ef4444;">No data found. Need at least header and one row.</span>';
            confirmBtn.disabled = true;
            return;
        }

        const headers = rawRows[0].map(h => String(h).trim().toLowerCase());
        const nameIndex = headers.findIndex(h => h.includes('name') && h.includes('column'));
        const altNameIndex = headers.indexOf('name');
        const finalNameIndex = nameIndex !== -1 ? nameIndex : altNameIndex;
        const typeIndex = headers.indexOf('type');
        const orderIndex = headers.findIndex(h => h.includes('order'));

        if (finalNameIndex === -1) {
            preview.innerHTML = '<span style="color:#ef4444;">Header must include a column with "name" (preferably "column name").</span>';
            confirmBtn.disabled = true;
            return;
        }

        const columnsToAdd = [];
        for (let i = 1; i < rawRows.length; i++) {
            const row = rawRows[i];
            const name = String(row[finalNameIndex] || '').trim();
            if (!name) continue;
            const type = String(row[typeIndex] || 'text').trim().toLowerCase();
            const validTypes = ['text', 'number', 'date'];
            const columnType = validTypes.includes(type) ? type : 'text';
            const order = parseInt(row[orderIndex] || 0) || 0;
            columnsToAdd.push({ column_name: name, column_type: columnType, display_order: order });
        }

        if (!columnsToAdd.length) {
            preview.innerHTML = '<span style="color:#ef4444;">No valid columns to add.</span>';
            confirmBtn.disabled = true;
            return;
        }

        // Sort by provided order, then assign sequential orders
        columnsToAdd.sort((a, b) => a.display_order - b.display_order);
        const currentMaxOrder = Math.max(...this.state.currentTemplate.doc_columns.map(c => c.display_order), -1);
        columnsToAdd.forEach((col, idx) => {
            col.display_order = currentMaxOrder + 1 + idx;
        });

        this.state._columnsToImport = columnsToAdd;

        const sample = columnsToAdd.slice(0, 5).map(c => `${c.column_name} (${c.column_type})`).join(', ');
        preview.innerHTML = `
            <div style="font-size:12px;line-height:1.8;padding:10px 12px;background:#f8fafc;border-radius:7px;border:1px solid #e2e8f0;">
                <div><strong>${columnsToAdd.length} columns</strong> will be added</div>
                <div>Sample: ${sample}${columnsToAdd.length > 5 ? '...' : ''}</div>
            </div>
        `;
        confirmBtn.disabled = false;
    },

    confirmImportColumns: async function() {
        if (!this.state._columnsToImport || !this.state._columnsToImport.length) return;

        const confirmBtn = document.getElementById('importColumnsConfirmBtn');
        if (confirmBtn) {
            confirmBtn.disabled = true;
            confirmBtn.innerText = 'Importing...';
        }

        try {
            const inserts = this.state._columnsToImport.map(col => ({
                template_id: this.state.currentTemplate.id,
                column_name: col.column_name,
                column_type: col.column_type,
                display_order: col.display_order
            }));

            const { data, error } = await supabaseClient
                .from('doc_columns')
                .insert(inserts)
                .select();
            if (error) throw error;

            this.state.currentTemplate.doc_columns.push(...data);
            this.state.currentTemplate.doc_columns.sort((a, b) => a.display_order - b.display_order);
            delete this.state.cache[this.state.currentTemplate.name];

            this.renderAll();
            this.closeImportColumnsModal();
            this.showToast(`${data.length} columns added!`);
        } catch (err) {
            this.showToast('Import failed: ' + err.message, 'error');
        } finally {
            if (confirmBtn) {
                confirmBtn.disabled = false;
                confirmBtn.innerText = 'Import';
            }
        }
    },

    openColumnModal: function() {
        document.getElementById('columnModal').style.display = 'block';
        this.addImportColumnsButton();
    },

    // ============================================================
    // DROPDOWN TOGGLE
    // ============================================================
    toggleMenu: function(event, menuId) {
        event.stopPropagation();
        const menu   = document.getElementById(menuId);
        if (!menu) return;
        const isOpen = menu.style.display === 'block';
        document.querySelectorAll('.dropdown').forEach(d => d.style.display = 'none');
        if (!isOpen) menu.style.display = 'block';
    }
};