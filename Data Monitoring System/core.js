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
        _importWorkbook: null
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
        window.openColumnModal   = ()          => document.getElementById('columnModal').style.display = 'block';
        window.closeColumnModal  = ()          => document.getElementById('columnModal').style.display = 'none';

        window.createNewCategory = ()          => this.createNewCategory();
        window.addColumnToActive = ()          => this.addColumnToActive();
        window.deleteColumn      = (id, name)  => this.deleteColumn(id, name);
        window.deleteCategory    = (id, name)  => this.deleteCategory(id, name);
        window.toggleMenu        = (event, id) => this.toggleMenu(event, `menu-${id}`);

        window.openImportModal   = ()          => this.openImportModal();
        window.closeImportModal  = ()          => this.closeImportModal();
        window.loadSheets        = ()          => this.loadSheets();
        window.confirmImport     = ()          => this.confirmImport();

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
                        supabaseClient.from('doc_entries').select('*').eq('template_id', templateId).order('created_at', { ascending: false })
                    ]);

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
    },

    renderTable: function(entries) {
        const body = document.getElementById('tableData');
        if (!body) return;

        body.innerHTML = entries.length
            ? entries.map(e => `
                <tr>
                    ${this.state.currentTemplate.doc_columns.map(c => `<td>${e.content[c.column_name] || '-'}</td>`).join('')}
                    <td class="action-buttons">
                        <button class="edit-btn" onclick="editEntry('${e.id}')">Edit</button>
                        <button class="del-btn" onclick="deleteEntry('${e.id}')">Delete</button>
                    </td>
                </tr>
            `).join('')
            : '<tr><td colspan="100%" style="text-align:center;padding:40px;color:#94a3b8;">No records found.</td></tr>';
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
                this.state.localEntries.unshift(res.data[0]);
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

    previewSheet: function() {
        const sheetName = document.getElementById('importSheet').value;
        const ws        = this.state._importWorkbook.Sheets[sheetName];
        const rows      = XLSX.utils.sheet_to_json(ws, { defval: '' });

        const preview    = document.getElementById('importPreview');
        const confirmBtn = document.getElementById('importConfirmBtn');

        if (!rows.length) {
            preview.innerHTML = '<span style="color:#ef4444;">No data found in this sheet.</span>';
            confirmBtn.disabled = true;
            return;
        }

        const cols      = this.state.currentTemplate.doc_columns.map(c => c.column_name);
        const excelCols = Object.keys(rows[0]);
        const matched   = cols.filter(c => excelCols.includes(c));
        const unmatched = cols.filter(c => !excelCols.includes(c));

        preview.innerHTML = `
            <div style="font-size:12px;line-height:1.8;padding:10px 12px;background:#f8fafc;border-radius:7px;border:1px solid #e2e8f0;">
                <div><strong>${rows.length} rows</strong> found in sheet</div>
                <div>Matched: <span style="color:#16a34a;font-weight:600;">${matched.join(', ') || 'none'}</span></div>
                ${unmatched.length ? `<div>Will be blank: <span style="color:#b45309;">${unmatched.join(', ')}</span></div>` : ''}
            </div>
        `;
        confirmBtn.disabled = matched.length === 0;
    },

    confirmImport: async function() {
        const sheetName = document.getElementById('importSheet').value;
        const ws        = this.state._importWorkbook.Sheets[sheetName];
        const rows      = XLSX.utils.sheet_to_json(ws, { defval: '' });
        const cols      = this.state.currentTemplate.doc_columns.map(c => c.column_name);

        const entries = rows.map(row => ({
            template_id: this.state.currentTemplate.id,
            content: Object.fromEntries(
                cols.map(c => [c, row[c] !== undefined ? String(row[c]) : ''])
            )
        }));

        const confirmBtn     = document.getElementById('importConfirmBtn');
        confirmBtn.disabled  = true;
        confirmBtn.innerText = 'Importing...';

        try {
            const batchSize = 100;
            for (let i = 0; i < entries.length; i += batchSize) {
                const batch = entries.slice(i, i + batchSize);
                const { error } = await supabaseClient.from('doc_entries').insert(batch);
                if (error) throw error;
            }

            const { data } = await supabaseClient
                .from('doc_entries')
                .select('*')
                .eq('template_id', this.state.currentTemplate.id)
                .order('created_at', { ascending: false });

            this.state.localEntries = data || [];
            this.state.cache[this.state.currentTemplate.name].entries = this.state.localEntries;

            this.renderTable(this.state.localEntries);
            this.closeImportModal();
            this.showToast(`${entries.length} rows imported successfully!`);

        } catch (err) {
            this.showToast('Import failed: ' + err.message, 'error');
        } finally {
            confirmBtn.disabled  = false;
            confirmBtn.innerText = 'Import';
        }
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