export const UI = {
    setLoading(isLoading) {
        const workspace = document.getElementById('moduleWorkspace');
        if (!workspace) return;
        workspace.style.opacity = isLoading ? "0.3" : "1";
        workspace.style.pointerEvents = isLoading ? "none" : "auto";
    },

    showToast(message, type = 'success') {
        const toast = document.createElement('div');
        toast.className = `toast toast-${type} show`;
        toast.innerText = message;
        document.body.appendChild(toast);
        setTimeout(() => { toast.classList.remove('show'); setTimeout(() => toast.remove(), 500); }, 3000);
    },

    renderCategoryCards(templates, activeName) {
        const container = document.getElementById('categoryCards');
        if (!container) return;
        container.innerHTML = templates.map(t => `
            <div class="category-card ${t.name === activeName ? 'active' : ''}" id="card-${t.id}" onclick="switchCategory('${t.name}')">
                <div class="card-menu" onclick="event.stopPropagation();">
                    <button class="menu-btn" onclick="toggleMenu(event, '${t.id}')">⋮</button>
                    <div id="menu-${t.id}" class="dropdown">
                        <button class="delete-menu-item" onclick="deleteCategory('${t.id}', '${t.name}')">🗑️ Delete</button>
                    </div>
                </div>
                <div class="card-icon">${t.name.substring(0, 2).toUpperCase()}</div>
                <span class="card-label">${t.name}</span>
            </div>
        `).join('');
    },

    renderForm(columns, onSave) {
        const container = document.getElementById('dynamicForm');
        container.innerHTML = columns.map(c => `
            <div class="input-box">
                <label>${c.column_name}</label>
                <input type="${c.column_type}" id="input_${c.column_name}" placeholder="Enter ${c.column_name}">
            </div>
        `).join('') + `<button id="mainSaveBtn" class="save-btn">Save Record</button>`;
        document.getElementById('mainSaveBtn').onclick = onSave;
    },

    renderTable(columns, entries) {
        const head = document.getElementById('tableHeaders');
        const body = document.getElementById('tableData');
        head.innerHTML = `<tr>${columns.map(c => `<th>${c.column_name}</th>`).join('')}<th>Actions</th></tr>`;
        body.innerHTML = entries.length ? entries.map(e => `
            <tr>
                ${columns.map(c => `<td>${e.content[c.column_name] || '-'}</td>`).join('')}
                <td class="action-buttons">
                    <button class="edit-btn" onclick="editEntry('${e.id}')">Edit</button>
                    <button class="del-btn" onclick="deleteEntry('${e.id}')">Del</button>
                </td>
            </tr>
        `).join('') : '<tr><td colspan="100%" style="text-align:center; padding:20px;">No records.</td></tr>';
    }
};