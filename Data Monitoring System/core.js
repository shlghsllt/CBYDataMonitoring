import { SUPABASE_CONFIG } from './config.js';

const { createClient } = supabase;
const supabaseClient = createClient(SUPABASE_CONFIG.url, SUPABASE_CONFIG.anonKey);

export const AppCore = {
    state: {
        moduleName:             '',
        currentTemplate:        null,
        allTemplates:           [],
        localEntries:           [],
        dateSortAsc:            true,
        editingId:              null,
        cache:                  {},
        isLoading:              false,
        _importWorkbook:        null,
        _importExcelCols:       [],   // detected Excel column names
        tableEventsInitialized: false
    },

    // ============================================================
    // INIT
    // ============================================================
    init: async function (moduleName) {
        this.state.moduleName = moduleName;
        this.syncWithWindow();
        await this.refreshCategories();
    },    

    syncWithWindow: function () {
        window.switchCategory    = (name)     => this.switchCategory(name);
        window.saveData          = ()          => this.saveData();
        window.editEntry         = (id)        => this.editEntry(id);
        window.closeEditModal    = ()          => this.closeEditModal();
        window.saveEditEntry     = ()          => this.saveEditEntry();
        window.deleteEntry       = (id)        => this.deleteEntry(id);
        window.searchData        = ()          => this.searchData();
        window.sortByDate        = ()          => this.sortByDate();
        window.exportToExcel     = ()          => this.exportToExcel();


        window.openModal         = ()          => document.getElementById('categoryModal').style.display = 'block';
        window.closeModal        = ()          => document.getElementById('categoryModal').style.display = 'none';
        window.openColumnModal   = ()          => document.getElementById('columnModal').style.display = 'block';
        window.closeColumnModal  = ()          => document.getElementById('columnModal').style.display = 'none';

        window.filterCategoryCards = () => this.filterCategoryCards();

        window.createNewCategory = ()          => this.createNewCategory();
        window.addColumnToActive = ()          => this.addColumnToActive();
        window.deleteColumn      = (id, name)  => this.deleteColumn(id, name);
        window.deleteCategory    = (id, name)  => this.deleteCategory(id, name);
        window.toggleMenu        = (event, id) => this.toggleMenu(event, `menu-${id}`);

        window.openImportModal   = ()          => this.openImportModal();
        window.closeImportModal  = ()          => this.closeImportModal();
        window.loadSheets        = ()          => this.loadSheets();
        window.previewSheet      = ()          => this.previewSheet();
        window.confirmImport     = ()          => this.confirmImport();
        window.deleteSelected    = ()          => this.deleteSelected();

        window.addEventListener('click', () => {
            document.querySelectorAll('.dropdown').forEach(d => d.style.display = 'none');
        });

        this.ensureContextMenu();

        // CONTEXT MENU ACTIONS
        document.getElementById('ctxEdit')?.addEventListener('click', () => {
            if (this.state.currentRowId) {
                this.editEntry(this.state.currentRowId);
            }
        });

        document.getElementById('ctxDelete')?.addEventListener('click', () => {
            if (this.state.currentRowId) {
                this.deleteEntry(this.state.currentRowId);
            }
        });     
    },

    ensureContextMenu: function () {
        this.injectContextMenuStyles();
        if (document.getElementById('contextMenu')) return;

        const menu = document.createElement('div');
        menu.id = 'contextMenu';
        menu.className = 'context-menu';

        const editBtn = document.createElement('button');
        editBtn.id = 'ctxEdit';
        editBtn.type = 'button';
        editBtn.textContent = 'Edit Row';

        const deleteBtn = document.createElement('button');
        deleteBtn.id = 'ctxDelete';
        deleteBtn.type = 'button';
        deleteBtn.textContent = 'Delete Row';
        deleteBtn.className = 'delete';

        menu.appendChild(editBtn);
        menu.appendChild(deleteBtn);
        document.body.appendChild(menu);
        this.injectContextMenuStyles();
    },

    injectContextMenuStyles: function () {
        if (document.getElementById('appcore-context-menu-styles')) return;

        const style = document.createElement('style');
        style.id = 'appcore-context-menu-styles';
        style.textContent = `
            .context-menu {
                position: absolute;
                display: none;
                min-width: 160px;
                background: #ffffff;
                border-radius: 10px;
                box-shadow: 0 8px 25px rgba(0,0,0,0.15);
                padding: 6px 0;
                z-index: 9999;
                animation: fadeInMenu 0.15s ease;
                border: 1px solid #eee;
            }

            .context-menu button {
                width: 100%;
                padding: 10px 14px;
                border: none;
                background: transparent;
                text-align: left;
                font-size: 14px;
                cursor: pointer;
                display: flex;
                align-items: center;
                gap: 10px;
                transition: background 0.2s ease, padding-left 0.2s ease;
            }

            .context-menu button:hover {
                background: #f5f7fb;
                padding-left: 18px;
            }

            .context-menu button:active {
                background: #eaeef5;
            }

            .context-menu button.delete {
                color: #ef4444;
            }

            .context-menu button.delete:hover {
                background: #fee2e2;
            }

            tbody td.cell-focused,
            tbody td[contenteditable="true"]:focus {
                background: #eff6ff;
                box-shadow: inset 0 0 0 1.5px #2563eb;
                border-radius: 4px;
            }

            tbody td.cell-selected {
                background: #dbeafe !important;
                box-shadow: inset 0 0 0 1px #3b82f6;
                border-radius: 2px;
                user-select: none;
            }

            @keyframes fadeInMenu {
                from {
                    opacity: 0;
                    transform: scale(0.95) translateY(-5px);
                }
                to {
                    opacity: 1;
                    transform: scale(1) translateY(0);
                }
            }

            tr.row-selected td {
                background: #e0f2fe !important;
            }
        `;

        document.head.appendChild(style);
    },

    // ============================================================
    // TOAST
    // ============================================================
    showToast: function (message, type = 'success') {
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
    switchCategory: async function (name) {
        if (this.state.isLoading) return;

        this.updateActiveUI(name);
        const workspace = document.getElementById('moduleWorkspace');
        workspace.style.display       = 'block';
        workspace.style.opacity       = '0.4';
        workspace.style.pointerEvents = 'none';
        this.state.isLoading = true;

        await new Promise(resolve => {
            requestAnimationFrame(async () => {
                try {
                    if (this.state.cache[name]) {
                        this.state.currentTemplate = this.state.cache[name].template;
                        this.state.localEntries    = this.state.cache[name].entries;
                    } else {
                        const templateId = this.state.allTemplates.find(t => t.name === name)?.id;
                        if (!templateId) throw new Error('Template not found');

                        const [tRes, eRes] = await Promise.all([
                            supabaseClient.from('doc_templates').select('*, doc_columns(*)').eq('id', templateId).single(),
                            supabaseClient.from('doc_entries').select('*').eq('template_id', templateId)
                        ]);

                        if (tRes.error) throw tRes.error;
                        if (eRes.error) throw eRes.error;

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
                    resolve();
                }
            });
        });
    },

    updateActiveUI: function (name) {
        document.querySelectorAll('.category-card').forEach(c => c.classList.remove('active'));
        const activeCard = document.getElementById(`card-${name}`);
        if (activeCard) activeCard.classList.add('active');
        const label = document.getElementById('activeCategoryName');
        if (label) label.innerText = name;
    },

    // ============================================================
    // RENDER
    // ============================================================
    renderAll: function () {
        const form = document.getElementById('dynamicForm');
        if (!form) return;

        // FIX #2: use correct input type per column_type
        form.innerHTML = this.state.currentTemplate.doc_columns.map(c => `
            <div class="input-box">
                <label>${c.column_name}</label>
                <input type="text"
                    id="input_${c.column_name}"
                    placeholder="Enter ${c.column_name}"
                    step="${c.column_type === 'number' ? 'any' : ''}">
            </div>
        `).join('') + `<button onclick="saveData()" class="save-btn" id="mainSaveBtn">Save Record</button>`;

        const headers = document.getElementById('tableHeaders');
        headers.innerHTML = `<tr>
            <th><input type="checkbox" id="selectAll"></th>
            ${this.state.currentTemplate.doc_columns.map(c => `
                <th>
                    <div class="th-inner">
                        <span>${c.column_name}</span>
                        <button class="del-col-btn" title="Delete column" onclick="deleteColumn('${c.id}', '${c.column_name}')">✕</button>
                    </div>
                </th>
            `).join('')}
        </tr>`;

        this.renderTable(this.state.localEntries);
        this.setupTableEditing();
    },

    renderTable: function (entries) {
        const body = document.getElementById('tableData');
        if (!body) return;

        if (!entries.length) {
            body.innerHTML = `<tr><td colspan="100%" class="no-data">No records found.</td></tr>`;
            return;
        }

        body.innerHTML = entries.map(e => `
            <tr data-entry-id="${e.id}">
                <td><input type="checkbox" class="rowCheckbox" data-id="${e.id}"></td>
                ${this.state.currentTemplate.doc_columns.map(c => {
                    const rawVal = e.content ? (e.content[c.column_name] ?? '') : '';
                    const val = this.formatDisplayValue(rawVal, c.column_type);
                    return `<td contenteditable="true" data-col-name="${c.column_name}">${val}</td>`;
                }).join('')}
            </tr>
        `).join('');
    },

    formatDisplayValue: function (raw, colType) {
        if (colType === 'date') {
            if (raw instanceof Date && !isNaN(raw.getTime())) {
                return this.formatDateDisplay(raw);
            }
            return String(raw ?? '');
        }
        if (raw instanceof Date && !isNaN(raw.getTime())) {
            return raw.toString();
        }
        return String(raw ?? '');
    },

    // ============================================================
    // INLINE TABLE EDITING + MULTI-CELL SELECTION
    // ============================================================
    setupTableEditing: function () {
        if (this.state.tableEventsInitialized) return;
        const body = document.getElementById('tableData');
        if (!body) return;

        // ---- selection state ----
        let isSelecting   = false;
        let selStartTd    = null;
        let selEndTd      = null;

        const getCellPos = (td) => {
            const row = td.closest('tr');
            const rows = Array.from(body.rows);
            const ri = rows.indexOf(row);
            const cells = Array.from(row.querySelectorAll('td[data-col-name]'));
            const ci = cells.indexOf(td);
            return { ri, ci };
        };

        const clearSelection = () => {
            body.querySelectorAll('td.cell-selected').forEach(c => c.classList.remove('cell-selected'));
            selStartTd = null;
            selEndTd   = null;
        };

        const applySelection = (startTd, endTd) => {
            body.querySelectorAll('td.cell-selected').forEach(c => c.classList.remove('cell-selected'));
            const s = getCellPos(startTd);
            const e = getCellPos(endTd);
            const minR = Math.min(s.ri, e.ri), maxR = Math.max(s.ri, e.ri);
            const minC = Math.min(s.ci, e.ci), maxC = Math.max(s.ci, e.ci);
            const rows = Array.from(body.rows);
            for (let r = minR; r <= maxR; r++) {
                if (!rows[r]) continue;
                const cells = Array.from(rows[r].querySelectorAll('td[data-col-name]'));
                for (let c = minC; c <= maxC; c++) {
                    if (cells[c]) cells[c].classList.add('cell-selected');
                }
            }
        };

        // mousedown — start selection or focus single cell
        body.addEventListener('mousedown', (e) => {
            const td = e.target.closest('td[data-col-name]');
            if (!td) { clearSelection(); return; }

            if (e.shiftKey && selStartTd) {
                // Shift+click extends selection
                selEndTd = td;
                applySelection(selStartTd, selEndTd);
                e.preventDefault();
                return;
            }

            clearSelection();
            isSelecting = true;
            selStartTd  = td;
            selEndTd    = td;
            td.classList.add('cell-selected');
        });

        // mouseover — drag to extend selection
        body.addEventListener('mouseover', (e) => {
            if (!isSelecting || !selStartTd) return;
            const td = e.target.closest('td[data-col-name]');
            if (!td) return;
            selEndTd = td;
            applySelection(selStartTd, selEndTd);
        });

        document.addEventListener('mouseup', () => { isSelecting = false; });

        // focusin — single cell edit focus
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

        // keydown — Ctrl+C copies selection, Enter blurs cell
        body.addEventListener('keydown', (e) => {
            // Ctrl+C or Cmd+C — copy selected cells
            if ((e.ctrlKey || e.metaKey) && e.key === 'c') {
                const selected = body.querySelectorAll('td.cell-selected');
                if (!selected.length) return; // let browser handle normal copy

                e.preventDefault();
                // Group selected cells by row for TSV output
                const rowMap = new Map();
                selected.forEach(td => {
                    const row = td.closest('tr');
                    const rows = Array.from(body.rows);
                    const ri = rows.indexOf(row);
                    if (!rowMap.has(ri)) rowMap.set(ri, []);
                    rowMap.get(ri).push(td.textContent);
                });
                const tsv = Array.from(rowMap.keys()).sort((a,b) => a-b)
                    .map(ri => rowMap.get(ri).join('\t')).join('\n');
                navigator.clipboard.writeText(tsv).catch(() => {});
                this.showToast(`Copied ${selected.length} cell(s)`);
                return;
            }

            this.onTableCellKeyDown(e);
        });

        body.addEventListener('paste', (e) => this.onTableCellPaste(e));

        // Click outside table clears selection
        document.addEventListener('click', (e) => {
            if (!body.contains(e.target)) clearSelection();
        });

        this.state.tableEventsInitialized = true;

        //Select All Funtion para sa mga cells to ya
        document.addEventListener('change', (e) => {
            if (e.target.id === 'selectAll') {
                document.querySelectorAll('.rowCheckbox')
                    .forEach(cb => cb.checked = e.target.checked);
            }
        });
        
        //Right Click Logic
        this.state.currentRowId = null;
        const menu = document.getElementById('contextMenu');
        
        body.addEventListener('contextmenu', (e) => {
            const td = e.target.closest('td[data-col-name]');
            if (!td) return;

            e.preventDefault();

            const row = td.closest('tr');

            // CLEAR previous row highlight
            body.querySelectorAll('tr.row-selected')
                .forEach(r => r.classList.remove('row-selected'));

            // HIGHLIGHT buong row
            row.classList.add('row-selected');

            this.state.currentRowId = row.dataset.entryId;

            // Show menu
            menu.style.display = 'block';
            menu.style.top = e.pageY + 'px';
            menu.style.left = e.pageX + 'px';
        });

        // Hide menu on click
        document.addEventListener('click', () => {
            menu.style.display = 'none';

            document.querySelectorAll('tr.row-selected')
                .forEach(r => r.classList.remove('row-selected'));
        });
    },

    parseTabular: function (text) {
        return text.replace(/\r/g, '').split('\n').filter(l => l !== '').map(l => l.split('\t'));
    },

    getTableCellInfo: function (td) {
        if (!td || td.tagName !== 'TD') return null;
        const row     = td.closest('tr');
        const entryId = row?.dataset.entryId;
        const colName = td.dataset.colName;
        return entryId && colName ? { entryId, colName } : null;
    },

    onTableCellBlur: function (td) {
        const info = this.getTableCellInfo(td);
        if (!info) return;
        const entry = this.state.localEntries.find(e => e.id === info.entryId);
        if (!entry) return;
        const newValue     = td.textContent.trim();
        const currentValue = String(entry.content[info.colName] ?? '');
        if (newValue === currentValue) return;
        entry.content[info.colName] = newValue;
        this.saveEntryField(entry.id, entry.content);
    },

    onTableCellKeyDown: function (e) {
        const td = e.target;
        if (td.tagName !== 'TD' || !td.isContentEditable) return;
        if (e.key === 'Enter') { e.preventDefault(); td.blur(); }
    },

    onTableCellPaste: function (e) {
        const text = e.clipboardData?.getData('text/plain') || '';
        if (!text) return;

        // Use focused cell or the first selected cell as paste anchor
        const td = e.target?.closest?.('td[data-col-name]')
            || document.querySelector('#tableData td.cell-selected');
        if (!td) return;

        e.preventDefault();

        const pasted = this.parseTabular(text);
        if (!pasted.length) return;

        const tbody          = document.getElementById('tableData');
        const rows           = Array.from(tbody.rows);
        const startRowIndex  = rows.indexOf(td.closest('tr'));
        const columns        = this.state.currentTemplate.doc_columns.map(c => c.column_name);
        const startColIndex  = columns.indexOf(td.dataset.colName);
        const changedEntries = new Map();

        pasted.forEach((rowValues, rowOffset) => {
            const targetRow = rows[startRowIndex + rowOffset];
            if (!targetRow) return;
            const entryId = targetRow.dataset.entryId;
            const entry   = this.state.localEntries.find(e => e.id === entryId);
            if (!entry) return;
            rowValues.forEach((cellValue, colOffset) => {
                const colName = columns[startColIndex + colOffset];
                if (!colName) return;
                const cell = targetRow.querySelector(`td[data-col-name="${colName}"]`);
                if (!cell) return;
                const normalized = cellValue.trim();
                if (String(entry.content[colName] ?? '') === normalized) return;
                entry.content[colName] = normalized;
                cell.textContent = normalized;
                changedEntries.set(entryId, entry);
            });
        });

        if (!changedEntries.size) return;
        changedEntries.forEach((entry) => this.saveEntryField(entry.id, entry.content));
    },

    saveEntryField: async function (entryId, content) {
        try {
            const { error } = await supabaseClient
                .from('doc_entries').update({ content }).eq('id', entryId);
            if (error) throw error;
            if (this.state.cache[this.state.currentTemplate.name]) {
                this.state.cache[this.state.currentTemplate.name].entries = this.state.localEntries;
            }
        } catch (err) {
            this.showToast('Save failed: ' + err.message, 'error');
        }
    },

    // ============================================================
    // EDIT ENTRY — POPUP MODAL
    // ============================================================
    editEntry: function (id) {
        const entry = this.state.localEntries.find(e => e.id === id);
        if (!entry) return;
        this.state.editingId = id;

        const editForm = document.getElementById('editForm');
        // FIX #2: use correct input types in edit modal too
        editForm.innerHTML = this.state.currentTemplate.doc_columns.map(c => {
            const inputType = 'text';
            const rawVal    = String(entry.content[c.column_name] ?? '');
            // For date inputs, value must be YYYY-MM-DD
            const val = (inputType === 'date' && rawVal && !/^\d{4}-\d{2}-\d{2}$/.test(rawVal))
                ? this.anyDateToISO(rawVal) || rawVal
                : rawVal;
            return `
            <div class="input-box">
                <label>${c.column_name}</label>
                <input type="${inputType}"
                    id="edit_input_${c.column_name}"
                    value="${val.replace(/"/g, '&quot;')}"
                    step="${inputType === 'number' ? 'any' : ''}"
                    placeholder="Enter ${c.column_name}">
            </div>`;
        }).join('');

        document.getElementById('editModal').style.display = 'block';
    },

    closeEditModal: function () {
        document.getElementById('editModal').style.display = 'none';
        this.state.editingId = null;
    },

    saveEditEntry: async function () {
        if (!this.state.editingId) return;
        const content = {};
        this.state.currentTemplate.doc_columns.forEach(c => {
            const el = document.getElementById(`edit_input_${c.column_name}`);
            content[c.column_name] = el ? el.value : '';
        });

        const saveBtn     = document.getElementById('editSaveBtn');
        saveBtn.disabled  = true;
        saveBtn.innerText = 'Saving...';

        try {
            const { data, error } = await supabaseClient
                .from('doc_entries').update({ content }).eq('id', this.state.editingId).select();
            if (error) throw error;

            const idx = this.state.localEntries.findIndex(e => e.id === this.state.editingId);
            if (idx !== -1) this.state.localEntries[idx] = data[0];

            if (this.state.cache[this.state.currentTemplate.name]) {
                this.state.cache[this.state.currentTemplate.name].entries = this.state.localEntries;
            }

            this.renderTable(this.state.localEntries);
            this.closeEditModal();
            this.showToast('Record updated!');
        } catch (err) {
            this.showToast('Update failed: ' + err.message, 'error');
        } finally {
            saveBtn.disabled  = false;
            saveBtn.innerText = 'Save Changes';
        }
    },

    // ============================================================
    // CATEGORIES
    // ============================================================
    refreshCategories: async function () {
        const { data, error } = await supabaseClient
            .from('doc_templates').select('*').eq('module', this.state.moduleName);
        if (error) { this.showToast('Failed to load categories.', 'error'); return; }
        this.state.allTemplates = data || [];
        this.renderCategoryCards();
    },

    renderCategoryCards: function (filteredTemplates) {
        const container = document.getElementById('categoryCards');
        if (!container) return;
        const templates = filteredTemplates || this.state.allTemplates;
        container.innerHTML = templates.map(t => {
            const hue   = Math.abs(t.name.split('').reduce((a, b) => (((a << 5) - a) + b.charCodeAt(0)) | 0, 0)) % 360;
            const color = `hsla(${hue}, 60%, 82%, 1)`;
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

    filterCategoryCards: function () {
        const searchInput = document.getElementById('categorySearch');
        if (!searchInput) return;
        const term = searchInput.value.toLowerCase();
        const filtered = this.state.allTemplates.filter(t => t.name.toLowerCase().includes(term));
        this.renderCategoryCards(filtered);
    },

    createNewCategory: async function () {
        const name = document.getElementById('newCategoryName').value.trim();
        if (!name) return this.showToast('Category name is required.', 'error');
        try {
            const { data, error } = await supabaseClient
                .from('doc_templates').insert([{ name, module: this.state.moduleName }]).select();
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

    deleteCategory: async function (id, name) {
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
    addColumnToActive: async function () {
        const name = document.getElementById('newColumnName').value.trim();
        const type = 'text';
        if (!name) return this.showToast('Column name is required.', 'error');
        if (!this.state.currentTemplate) return this.showToast('No category selected.', 'error');
        const order = this.state.currentTemplate.doc_columns.length;
        try {
            const { data, error } = await supabaseClient
                .from('doc_columns')
                .insert([{ template_id: this.state.currentTemplate.id, column_name: name, column_type: type, display_order: order }])
                .select();
            if (error) throw error;
            this.state.currentTemplate.doc_columns.push(data[0]);
            if (this.state.cache[this.state.currentTemplate.name]) {
                this.state.cache[this.state.currentTemplate.name].template = this.state.currentTemplate;
            }
            this.renderAll();
            this.showToast('Column added!');
            document.getElementById('newColumnName').value = '';
            window.closeColumnModal();
        } catch (err) {
            this.showToast('Failed to add column: ' + err.message, 'error');
        }
    },

    deleteColumn: async function (id, name) {
        if (!confirm(`Delete column "${name}"? This will affect all records.`)) return;
        const column = this.state.currentTemplate?.doc_columns.find(c => c.id === id);
        if (!column) return this.showToast('Column not found.', 'error');
        const columnName = column.column_name;

        try {
            const { error } = await supabaseClient.from('doc_columns').delete().eq('id', id);
            if (error) throw error;

            const { data: entries, error: fetchErr } = await supabaseClient
                .from('doc_entries')
                .select('id, content')
                .eq('template_id', this.state.currentTemplate.id);
            if (fetchErr) throw fetchErr;

            if (entries && entries.length) {
                const updates = entries
                    .filter(entry => entry.content && Object.prototype.hasOwnProperty.call(entry.content, columnName))
                    .map(entry => {
                        const updatedContent = { ...entry.content };
                        delete updatedContent[columnName];
                        return supabaseClient
                            .from('doc_entries')
                            .update({ content: updatedContent })
                            .eq('id', entry.id);
                    });

                if (updates.length) {
                    const results = await Promise.all(updates);
                    const failed = results.find(r => r.error);
                    if (failed) throw failed.error;
                }
            }

            if (this.state.cache[this.state.currentTemplate.name]) {
                delete this.state.cache[this.state.currentTemplate.name];
            }

            await this.reloadCurrentTemplate();
            this.showToast('Column deleted.');
        } catch (err) {
            this.showToast('Failed: ' + err.message, 'error');
        }
    },

    reloadCurrentTemplate: async function () {
        const templateId = this.state.currentTemplate?.id;
        if (!templateId) return;

        const [tRes, eRes] = await Promise.all([
            supabaseClient.from('doc_templates').select('*, doc_columns(*)').eq('id', templateId).single(),
            supabaseClient.from('doc_entries').select('*').eq('template_id', templateId)
        ]);

        if (tRes.error) throw tRes.error;
        if (eRes.error) throw eRes.error;

        const template = tRes.data;
        template.doc_columns.sort((a, b) => a.display_order - b.display_order);

        this.state.currentTemplate = template;
        this.state.localEntries    = eRes.data || [];
        this.state.cache[template.name] = { template, entries: this.state.localEntries };

        this.updateActiveUI(template.name);
        this.renderAll();
    },

    // ============================================================
    // ENTRIES — SAVE / DELETE
    // ============================================================
    saveData: async function () {
        const content = {};
        this.state.currentTemplate.doc_columns.forEach(c => {
            const el = document.getElementById(`input_${c.column_name}`);
            content[c.column_name] = el ? el.value : '';
        });

        const btn = document.getElementById('mainSaveBtn');
        btn.disabled = true;

        try {
            const { data, error } = await supabaseClient
                .from('doc_entries')
                .insert([{ template_id: this.state.currentTemplate.id, content }])
                .select();
            if (error) throw error;

            this.state.localEntries.push(data[0]);
            if (this.state.cache[this.state.currentTemplate.name]) {
                this.state.cache[this.state.currentTemplate.name].entries = this.state.localEntries;
            }
            this.renderTable(this.state.localEntries);
            this.state.currentTemplate.doc_columns.forEach(c => {
                const el = document.getElementById(`input_${c.column_name}`);
                if (el) el.value = '';
            });
            this.showToast('Saved successfully!');
        } catch (err) {
            this.showToast(err.message, 'error');
        } finally {
            btn.disabled = false;
        }
    },

    deleteEntry: async function (id) {
        if (!confirm('Are you sure you want to delete this record?')) return;
        try {
            const { error } = await supabaseClient.from('doc_entries').delete().eq('id', id);
            if (error) throw error;
            this.state.localEntries = this.state.localEntries.filter(e => e.id !== id);
            if (this.state.cache[this.state.currentTemplate.name]) {
                this.state.cache[this.state.currentTemplate.name].entries = this.state.localEntries;
            }
            this.renderTable(this.state.localEntries);
            this.showToast('Record deleted.');
        } catch (err) {
            this.showToast('Delete failed: ' + err.message, 'error');
        }
    },

    // ============================================================
    // SEARCH & SORT
    // ============================================================
    searchData: function () {
        const term     = document.getElementById('search').value.toLowerCase();
        const filtered = this.state.localEntries.filter(e =>
            JSON.stringify(e.content).toLowerCase().includes(term)
        );
        this.renderTable(filtered);
    },

    sortByDate: function () {
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
    exportToExcel: function () {
        if (!this.state.currentTemplate) return this.showToast('No category selected.', 'error');
        if (!this.state.localEntries.length) return this.showToast('No data to export.', 'error');
        const formatted = this.state.localEntries.map(e => e.content);
        const ws = XLSX.utils.json_to_sheet(formatted);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, this.state.currentTemplate.name);
        XLSX.writeFile(wb, `${this.state.currentTemplate.name}.xlsx`);
    },

    // ============================================================
    // DATE PARSING UTILITIES  (FIX #1)
    // Handles: serial, YYYY-MM-DD, DD-Mon-YY, DD-Mon-YYYY,
    //          MM/DD/YYYY, DD/MM/YYYY, Month DD YYYY, etc.
    // ============================================================
    MONTH_MAP: {
        jan:'01',feb:'02',mar:'03',apr:'04',may:'05',jun:'06',
        jul:'07',aug:'08',sep:'09',oct:'10',nov:'11',dec:'12'
    },

    excelSerialToISO: function (serial) {
        const num = parseInt(serial, 10);
        if (isNaN(num) || num < 1) return null;
        const date = new Date(Date.UTC(1899, 11, 30) + num * 86400000);
        return this._dateToISO(date);
    },

    _dateToISO: function (date) {
        const yyyy = date.getUTCFullYear();
        const mm   = String(date.getUTCMonth() + 1).padStart(2, '0');
        const dd   = String(date.getUTCDate()).padStart(2, '0');
        return `${yyyy}-${mm}-${dd}`;
    },

    formatDateDisplay: function (date) {
        if (!(date instanceof Date) || isNaN(date.getTime())) return '';
        const day   = date.getUTCDate();
        const mon   = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'][date.getUTCMonth()];
        const year  = String(date.getUTCFullYear()).slice(-2);
        return `${day}-${mon}-${year}`;
    },

    isExcelSerial: function (value) {
        const num = Number(value);
        return Number.isInteger(num) && num > 1 && num < 100000;
    },

    // FIX #1: Parse any common date string into YYYY-MM-DD
    anyDateToISO: function (raw) {
        if (!raw && raw !== 0) return '';
        if (raw instanceof Date && !isNaN(raw.getTime())) return this._dateToISO(raw);
        if (typeof raw === 'number') return this.excelSerialToISO(raw) || '';

        const s = String(raw).trim();
        if (!s) return '';

        // Already ISO: 2024-10-15
        if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;

        // Excel serial number
        if (this.isExcelSerial(s)) return this.excelSerialToISO(s);

        // DD-Mon-YY or DD-Mon-YYYY  e.g. 15-Oct-24, 15-Oct-2024
        const dMonY = s.match(/^(\d{1,2})[-\/\s]([A-Za-z]{3,})[-\/\s](\d{2,4})$/);
        if (dMonY) {
            const dd   = dMonY[1].padStart(2, '0');
            const mon  = this.MONTH_MAP[dMonY[2].toLowerCase().substring(0, 3)];
            let   year = dMonY[3];
            if (year.length === 2) year = parseInt(year) < 50 ? '20' + year : '19' + year;
            if (mon) return `${year}-${mon}-${dd}`;
        }

        // Mon-DD-YYYY or Mon DD YYYY  e.g. Oct-15-2024, October 15 2024
        const monDY = s.match(/^([A-Za-z]{3,})[-\/\s](\d{1,2})[-\/\s](\d{2,4})$/);
        if (monDY) {
            const mon  = this.MONTH_MAP[monDY[1].toLowerCase().substring(0, 3)];
            const dd   = monDY[2].padStart(2, '0');
            let   year = monDY[3];
            if (year.length === 2) year = parseInt(year) < 50 ? '20' + year : '19' + year;
            if (mon) return `${year}-${mon}-${dd}`;
        }

        // MM/DD/YYYY or DD/MM/YYYY — try MM/DD/YYYY first (US), fallback
        const slashParts = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
        if (slashParts) {
            let [, p1, p2, year] = slashParts;
            if (year.length === 2) year = parseInt(year) < 50 ? '20' + year : '19' + year;
            const mm = p1.padStart(2, '0');
            const dd = p2.padStart(2, '0');
            return `${year}-${mm}-${dd}`;
        }

        // Try native Date parse as last resort
        const d = new Date(s);
        if (!isNaN(d.getTime())) return this._dateToISO(d);

        return s; // return as-is if nothing matched
    },

    // ============================================================
    // IMPORT FROM EXCEL
    // ============================================================
    _el: function (id) { return document.getElementById(id); },

    openImportModal: function () {
        if (!this.state.currentTemplate)
            return this.showToast('Select a category first.', 'error');
        document.getElementById('importModal').style.display = 'block';
    },

    closeImportModal: function () {
        const modal = this._el('importModal');
        if (modal) modal.style.display = 'none';

        const safe = (id, prop, val) => { const el = this._el(id); if (el) el[prop] = val; };
        safe('importFile',        'value',     '');
        safe('importHeaderRow',   'value',     '1');
        safe('importSheet',       'innerHTML', '<option>— load a file first —</option>');
        safe('importSheet',       'disabled',  true);
        safe('importConfirmBtn',  'disabled',  true);
        safe('importPreview',     'innerHTML', '');
        safe('importExcelHeaders','innerHTML', '');
        safe('importColMapping',  'innerHTML', '');

        this.state._importWorkbook  = null;
        this.state._importExcelCols = [];
    },

    loadSheets: function () {
        const file = this._el('importFile')?.files[0];
        if (!file) return;
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const workbook = XLSX.read(e.target.result, { type: 'array', cellDates: true });
                this.state._importWorkbook = workbook;

                const sheetSelect = this._el('importSheet');
                sheetSelect.innerHTML = workbook.SheetNames.map(
                    name => `<option value="${name}">${name}</option>`
                ).join('');
                sheetSelect.disabled = false;
                sheetSelect.onchange = () => this.previewSheet();
                this.previewSheet();
            } catch (err) {
                this.showToast('Could not read file: ' + err.message, 'error');
            }
        };
        reader.readAsArrayBuffer(file);
    },

    previewSheet: function () {
        if (!this.state._importWorkbook) return;

        const sheetSelect = this._el('importSheet');
        const headerInput = this._el('importHeaderRow');
        const preview     = this._el('importPreview');
        const confirmBtn  = this._el('importConfirmBtn');
        const excelHdrs   = this._el('importExcelHeaders');
        const colMapping  = this._el('importColMapping');

        if (!sheetSelect || !preview || !confirmBtn) return;

        const sheetName = sheetSelect.value;
        const headerRow = Math.max(1, parseInt(headerInput?.value || '1') || 1);
        const ws        = this.state._importWorkbook.Sheets[sheetName];

        let rows = [];
        try {
            rows = XLSX.utils.sheet_to_json(ws, { defval: '', range: headerRow - 1, raw: false });
        } catch (err) {
            preview.innerHTML   = `<span style="color:#ef4444;">Error reading sheet: ${err.message}</span>`;
            confirmBtn.disabled = true;
            return;
        }

        if (!rows.length) {
            preview.innerHTML   = '<span style="color:#ef4444;">No data found. Try a different header row number.</span>';
            if (excelHdrs) excelHdrs.innerHTML = '';
            if (colMapping) colMapping.innerHTML = '';
            confirmBtn.disabled = true;
            return;
        }

        const excelCols = Object.keys(rows[0]);
        this.state._importExcelCols = excelCols;

        // Show detected Excel columns
        if (excelHdrs) {
            excelHdrs.innerHTML = `
                <p class="import-section-label">Detected Excel columns (row ${headerRow})</p>
                <div class="import-excel-cols">${excelCols.join(' &middot; ')}</div>
            `;
        }

        // FIX #4: Manual column mapping UI
        // Build a dropdown per DB column so user can pick which Excel column maps to it
        if (colMapping) {
            const dbCols = this.state.currentTemplate.doc_columns;
            const optionsHtml = `<option value="">(skip)</option>` +
                excelCols.map(c => `<option value="${c.replace(/"/g, '&quot;')}">${c}</option>`).join('');

            colMapping.innerHTML = `
                <p class="import-section-label" style="margin-top:12px;">Map columns</p>
                <div class="col-mapping-grid">
                    ${dbCols.map(c => {
                        // Auto-select exact match (case-insensitive)
                        const autoMatch = excelCols.find(
                            ec => ec.trim().toLowerCase() === c.column_name.trim().toLowerCase()
                        ) || '';
                        return `
                        <div class="col-mapping-row">
                            <span class="col-mapping-label" title="${c.column_type}">${c.column_name}
                                <small class="col-type-badge">${c.column_type}</small>
                            </span>
                            <select class="col-mapping-select" data-db-col="${c.column_name}" data-col-type="${c.column_type}">
                                ${optionsHtml}
                            </select>
                        </div>`;
                    }).join('')}
                </div>
            `;

            // Set auto-matched selections
            dbCols.forEach(c => {
                const autoMatch = excelCols.find(
                    ec => ec.trim().toLowerCase() === c.column_name.trim().toLowerCase()
                );
                if (autoMatch) {
                    const sel = colMapping.querySelector(`select[data-db-col="${c.column_name}"]`);
                    if (sel) sel.value = autoMatch;
                }
            });
        }

        preview.innerHTML = `
            <div class="import-preview-box">
                <div><strong>${rows.length.toLocaleString()} data rows</strong> found in sheet</div>
                <div style="font-size:11px;color:#64748b;margin-top:4px;">
                    Select which Excel column maps to each of your DB columns above.
                    Columns set to "(skip)" will not be stored — those fields will simply be absent.
                </div>
            </div>
        `;

        confirmBtn.disabled = false;
    },

    // FIX #1: Convert a value based on the column type
    convertValue: function (raw, colType) {
        if (colType === 'date') {
            if (raw instanceof Date && !isNaN(raw.getTime())) {
                return this.formatDateDisplay(raw);
            }
            if (typeof raw === 'number') {
                const date = new Date(Date.UTC(1899, 11, 30) + raw * 86400000);
                return this.formatDateDisplay(date);
            }
            const s = String(raw ?? '').trim();
            return s;
        }

        const s = String(raw ?? '').trim();
        if (!s) return '';


        if (colType === 'number') {
            // Strip commas from numbers like "1,440"
            const cleaned = s.replace(/,/g, '');
            return isNaN(Number(cleaned)) ? s : cleaned;
        }

        return s;
    },

    // FIX #3 + #5: confirmImport now uses mapping UI and re-fetches ordered data
    confirmImport: async function () {
        if (!this.state._importWorkbook) return this.showToast('No file loaded.', 'error');
        if (!this.state.currentTemplate)  return this.showToast('No category selected.', 'error');

        const sheetName = this._el('importSheet')?.value;
        const headerRow = Math.max(1, parseInt(this._el('importHeaderRow')?.value || '1') || 1);
        const ws        = this.state._importWorkbook.Sheets[sheetName];
        const rows      = XLSX.utils.sheet_to_json(ws, { defval: '', range: headerRow - 1, raw: false });

        if (!rows.length) return this.showToast('No data rows to import.', 'error');

        // Read the manual column mapping from UI
        const mappingSelects = document.querySelectorAll('.col-mapping-select');
        const mapping = {}; // dbColName -> excelColName
        mappingSelects.forEach(sel => {
            const dbCol  = sel.dataset.dbCol;
            const excelCol = sel.value;
            if (dbCol && excelCol) mapping[dbCol] = excelCol;
        });

        const dbCols = this.state.currentTemplate.doc_columns;

        const entries = rows.map(row => {
            const content = {};
            dbCols.forEach(c => {
                const excelColName = mapping[c.column_name] || '';
                // If skipped, omit the key entirely — do not store blank
                if (!excelColName) return;
                const rawVal = row[excelColName] ?? '';
                content[c.column_name] = this.convertValue(rawVal, c.column_type);
            });
            return { template_id: this.state.currentTemplate.id, content };
        });

        const confirmBtn     = this._el('importConfirmBtn');
        if (confirmBtn) { confirmBtn.disabled = true; confirmBtn.innerText = 'Importing...'; }

        try {
            const batchSize = 100;
            for (let i = 0; i < entries.length; i += batchSize) {
                const { error } = await supabaseClient
                    .from('doc_entries').insert(entries.slice(i, i + batchSize));
                if (error) throw error;
            }

            // FIX #3 + #5: re-fetch fresh data in correct order, then render immediately
            const { data: freshData, error: fetchErr } = await supabaseClient
                .from('doc_entries')
                .select('*')
                .eq('template_id', this.state.currentTemplate.id);
            if (fetchErr) throw fetchErr;

            this.state.localEntries = freshData || [];
            // Rebuild cache so it's not stale
            this.state.cache[this.state.currentTemplate.name] = {
                template: this.state.currentTemplate,
                entries:  this.state.localEntries
            };

            // FIX #3: render the table BEFORE closing modal so user sees data immediately
            this.renderTable(this.state.localEntries);
            this.closeImportModal();
            this.showToast(`${entries.length} rows imported successfully!`);

        } catch (err) {
            this.showToast('Import failed: ' + err.message, 'error');
        } finally {
            if (confirmBtn) { confirmBtn.disabled = false; confirmBtn.innerText = 'Import'; }
        }
    },

    // ============================================================
    // DROPDOWN TOGGLE
    // ============================================================
    toggleMenu: function (event, menuId) {
        event.stopPropagation();
        const menu   = document.getElementById(menuId);
        if (!menu) return;
        const isOpen = menu.style.display === 'block';
        document.querySelectorAll('.dropdown').forEach(d => d.style.display = 'none');
        if (!isOpen) menu.style.display = 'block';
    },

    // ============================================================
    //Delete Selection lang sa mga checkbox ituuu
    // ============================================================
    deleteSelected: async function () {
    const checked = Array.from(document.querySelectorAll('.rowCheckbox:checked'))
        .map(cb => cb.dataset.id);

    if (!checked.length) return this.showToast('No selected rows.', 'error');

    if (!confirm(`Delete ${checked.length} records?`)) return;

    try {
        const { error } = await supabaseClient
            .from('doc_entries')
            .delete()
            .in('id', checked);

        if (error) throw error;

        this.state.localEntries = this.state.localEntries.filter(e => !checked.includes(e.id));

        this.renderTable(this.state.localEntries);
        this.showToast('Deleted selected records.');
    } catch (err) {
        this.showToast('Delete failed: ' + err.message, 'error');
    }
    }
};