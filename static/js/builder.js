// TAMEER Scaffolding Quotation Builder JavaScript
// Comprehensive, easily viewable product selection & dynamic calculation engine

let catalogItems = [];
let activeCategoryFilter = "";
let sessionAddedCount = 0;

document.addEventListener("DOMContentLoaded", function () {
    fetch('/api/items')
        .then(response => response.json())
        .then(data => {
            catalogItems = data;
            initCategoryDropdowns();
        })
        .catch(err => console.error("Error loading scaffolding catalog:", err));

    attachGlobalListeners();
    updateTableUI();
});

// --- CATEGORIES & SEARCH ---
function initCategoryDropdowns() {
    const categories = [...new Set(catalogItems.map(item => item.category))].sort();
    const catSelect = document.getElementById("catalogCategorySelect");
    if (catSelect) {
        catSelect.innerHTML = '<option value="">All Categories (136 Items)</option>';
        categories.forEach(cat => {
            const count = catalogItems.filter(i => i.category === cat).length;
            const opt = document.createElement("option");
            opt.value = cat;
            opt.textContent = `${cat} (${count})`;
            catSelect.appendChild(opt);
        });
    }
}

// --- TOP QUICK PRODUCT SEARCH ---
function handleQuickSearch(input) {
    const query = input.value.toLowerCase().trim();
    const resultsContainer = document.getElementById("quickSearchResults");
    const clearBtn = document.getElementById("clearQuickSearchBtn");

    if (clearBtn) {
        clearBtn.style.display = query ? "block" : "none";
    }

    if (!resultsContainer) return;

    let filtered = catalogItems;

    if (activeCategoryFilter) {
        filtered = filtered.filter(item => item.category === activeCategoryFilter);
    }

    if (query) {
        filtered = filtered.filter(item => 
            item.item_name.toLowerCase().includes(query) || 
            item.category.toLowerCase().includes(query)
        );
    }

    // Limit to top 15 results for performance and clean viewing
    const displayItems = filtered.slice(0, 15);

    if (displayItems.length === 0) {
        resultsContainer.innerHTML = `
            <div class="p-3 text-center text-muted small">
                <i class="bi bi-search me-1"></i> No matching scaffolding items found for "${escapeHtml(query)}".
                <div class="mt-2">
                    <button type="button" class="btn btn-secondary-tmr btn-sm py-1" onclick="addCustomRowFromQuery('${escapeHtml(query)}')">
                        <i class="bi bi-plus"></i> Add "${escapeHtml(query)}" as Custom Item
                    </button>
                </div>
            </div>
        `;
        resultsContainer.classList.remove("d-none");
        return;
    }

    let html = `<div class="p-2 border-bottom bg-light d-flex justify-content-between align-items-center">
        <span class="text-muted small fw-semibold" style="font-size: 0.76rem;">FOUND ${filtered.length} COMPONENTS — CLICK TO ADD</span>
        <small class="text-muted" style="font-size: 0.72rem;">Press Esc to close</small>
    </div>`;

    displayItems.forEach(item => {
        html += `
            <div class="quick-search-item" onclick="addQuickSearchItem(${item.item_id})">
                <div>
                    <div class="item-title">${escapeHtml(item.item_name)}</div>
                    <div class="item-sub">
                        <span class="badge bg-secondary-subtle text-secondary me-1" style="font-size: 0.7rem;">${escapeHtml(item.category)}</span>
                        <span>Unit: <strong>${escapeHtml(item.unit || 'Pcs.')}</strong></span>
                    </div>
                </div>
                <div class="d-flex align-items-center gap-3">
                    <div class="text-end">
                        <div class="item-rate">SAR ${parseFloat(item.unit_price || 0).toFixed(2)}</div>
                        <small class="text-muted" style="font-size: 0.7rem;">Std. Sale Rate</small>
                    </div>
                    <button type="button" class="btn btn-primary-tmr btn-sm py-1 px-2">
                        <i class="bi bi-plus-lg"></i> Add
                    </button>
                </div>
            </div>
        `;
    });

    resultsContainer.innerHTML = html;
    resultsContainer.classList.remove("d-none");
}

function clearQuickSearch() {
    const input = document.getElementById("quickProductSearch");
    if (input) {
        input.value = "";
        input.focus();
    }
    const clearBtn = document.getElementById("clearQuickSearchBtn");
    if (clearBtn) clearBtn.style.display = "none";
    const resultsContainer = document.getElementById("quickSearchResults");
    if (resultsContainer) resultsContainer.classList.add("d-none");
}

function filterChipCategory(categoryName, chipBtn) {
    activeCategoryFilter = categoryName;

    // Update active class on chips
    document.querySelectorAll(".category-chip").forEach(chip => {
        chip.classList.remove("active");
    });
    if (chipBtn) chipBtn.classList.add("active");

    const searchInput = document.getElementById("quickProductSearch");
    if (searchInput) {
        handleQuickSearch(searchInput);
        searchInput.focus();
    }
}

function addQuickSearchItem(itemId) {
    const item = catalogItems.find(i => i.item_id === itemId);
    if (!item) return;

    const tr = addRow({
        desc: item.item_name,
        unit: item.unit || 'Pcs.',
        qty: 1,
        rate: item.unit_price || 0.0,
        remarks: ''
    });

    // Close dropdown and clear search
    const resultsContainer = document.getElementById("quickSearchResults");
    if (resultsContainer) resultsContainer.classList.add("d-none");

    const searchInput = document.getElementById("quickProductSearch");
    if (searchInput) searchInput.value = "";
    const clearBtn = document.getElementById("clearQuickSearchBtn");
    if (clearBtn) clearBtn.style.display = "none";

    // Focus quantity on new row
    if (tr) {
        tr.classList.add("item-highlight");
        const qtyInput = tr.querySelector(".item-qty");
        if (qtyInput) {
            qtyInput.focus();
            qtyInput.select();
        }
    }
}

function addCustomRowFromQuery(customName) {
    const tr = addRow({
        desc: customName || '',
        unit: 'Pcs.',
        qty: 1,
        rate: 0.0,
        remarks: ''
    });

    clearQuickSearch();

    if (tr) {
        tr.classList.add("item-highlight");
        const rateInput = tr.querySelector(".item-rate");
        if (rateInput) rateInput.focus();
    }
}

// --- CATALOG MODAL (BROWSE 136 ITEMS) ---
function filterCatalog() {
    const selectedCat = document.getElementById("catalogCategorySelect")?.value || "";
    const searchVal = (document.getElementById("catalogSearchInput")?.value || "").toLowerCase().trim();
    const container = document.getElementById("catalogItemsList");
    const countBadge = document.getElementById("catalogFilteredCount");
    if (!container) return;

    container.innerHTML = "";

    const filtered = catalogItems.filter(item => {
        const matchesCat = !selectedCat || item.category === selectedCat;
        const matchesSearch = !searchVal || 
            item.item_name.toLowerCase().includes(searchVal) ||
            item.category.toLowerCase().includes(searchVal);
        return matchesCat && matchesSearch;
    });

    if (countBadge) countBadge.textContent = `${filtered.length} items`;

    if (filtered.length === 0) {
        container.innerHTML = '<div class="col-12 p-4 text-muted text-center">No scaffolding components matched your criteria.</div>';
        return;
    }

    filtered.forEach(item => {
        const col = document.createElement("div");
        col.className = "col-md-6";
        col.innerHTML = `
            <div class="catalog-item-card d-flex justify-content-between align-items-center shadow-xs">
                <div>
                    <div class="fw-bold text-dark" style="font-size: 0.88rem;">${escapeHtml(item.item_name)}</div>
                    <div class="text-muted small mt-1">
                        <span class="badge bg-secondary-subtle text-secondary me-1">${escapeHtml(item.category)}</span>
                        <span>Unit: <strong>${escapeHtml(item.unit || 'Pcs.')}</strong></span>
                    </div>
                </div>
                <div class="d-flex align-items-center gap-3">
                    <span class="fw-bold text-dark" style="font-size: 0.9rem;">SAR ${parseFloat(item.unit_price).toFixed(2)}</span>
                    <button type="button" class="btn btn-secondary-tmr btn-sm py-1 px-3" onclick="addItemFromCatalog(${item.item_id}, this)">
                        <i class="bi bi-plus-lg"></i> Add
                    </button>
                </div>
            </div>
        `;
        container.appendChild(col);
    });
}

function addItemFromCatalog(itemId, btn) {
    const item = catalogItems.find(i => i.item_id === itemId);
    if (!item) return;

    addRow({
        desc: item.item_name,
        unit: item.unit || 'Pcs.',
        qty: 1,
        rate: item.unit_price || 0.0,
        remarks: ''
    });

    sessionAddedCount++;
    const badge = document.getElementById("modalAddedCountBadge");
    if (badge) {
        badge.className = "badge bg-success text-white px-3 py-2 fw-semibold";
        badge.innerHTML = `<i class="bi bi-check-circle-fill me-1"></i> ${sessionAddedCount} component(s) added to quotation`;
    }

    if (btn) {
        btn.className = "btn btn-success btn-sm py-1 px-3";
        btn.innerHTML = `<i class="bi bi-check2"></i> Added`;
        setTimeout(() => {
            btn.className = "btn btn-secondary-tmr btn-sm py-1 px-3";
            btn.innerHTML = `<i class="bi bi-plus-lg"></i> Add More`;
        }, 1200);
    }
}

// --- TABLE ROW MANAGEMENT ---
function addEmptyRow() {
    const tr = addRow({
        desc: '',
        unit: 'Pcs.',
        qty: 1,
        rate: 0.0,
        remarks: ''
    });
    if (tr) {
        const descInput = tr.querySelector(".item-desc");
        if (descInput) descInput.focus();
    }
}

function addRow(data = {}) {
    const tbody = document.getElementById("itemsTableBody");
    if (!tbody) return null;

    const rowIndex = tbody.children.length + 1;

    const tr = document.createElement("tr");
    tr.className = "item-row";
    tr.innerHTML = `
        <td class="text-center text-muted small row-sl fw-bold">${rowIndex}</td>
        <td>
            <div class="position-relative">
                <input type="text" class="form-control item-desc" value="${escapeHtml(data.desc || '')}" 
                       placeholder="Type component name or select above..." autocomplete="off" oninput="showSuggestions(this)">
                <div class="autocomplete-suggestions d-none"></div>
            </div>
        </td>
        <td>
            <select class="form-select item-unit text-center">
                <option value="Pcs." ${data.unit === 'Pcs.' ? 'selected' : ''}>Pcs.</option>
                <option value="Mtr" ${data.unit === 'Mtr' ? 'selected' : ''}>Mtr</option>
                <option value="Set" ${data.unit === 'Set' ? 'selected' : ''}>Set</option>
                <option value="LS" ${data.unit === 'LS' ? 'selected' : ''}>LS</option>
            </select>
        </td>
        <td>
            <input type="number" class="form-control text-end item-qty" value="${data.qty || 1}" min="1" step="any" oninput="calculateRow(this)">
        </td>
        <td>
            <input type="number" class="form-control text-end item-rate" value="${parseFloat(data.rate || 0).toFixed(2)}" step="0.01" min="0" oninput="calculateRow(this)">
        </td>
        <td>
            <input type="text" class="form-control text-end fw-semibold item-amount bg-light" value="0.00" readonly>
        </td>
        <td>
            <input type="text" class="form-control item-remarks" value="${escapeHtml(data.remarks || '')}" placeholder="Optional note">
        </td>
        <td class="text-center">
            <button type="button" class="icon-action-btn danger" onclick="removeRow(this)" title="Delete line">
                <i class="bi bi-trash"></i>
            </button>
        </td>
    `;

    tbody.appendChild(tr);
    calculateRow(tr.querySelector('.item-qty'));
    updateTableUI();
    return tr;
}

function removeRow(btn) {
    const row = btn.closest("tr");
    if (!row) return;
    row.remove();
    renumberRows();
    recalculateTotals();
    updateTableUI();
}

function clearAllRows() {
    const tbody = document.getElementById("itemsTableBody");
    if (!tbody || tbody.children.length === 0) return;

    if (confirm("Are you sure you want to remove all components from this quotation?")) {
        tbody.innerHTML = "";
        recalculateTotals();
        updateTableUI();
    }
}

function renumberRows() {
    const rows = document.querySelectorAll("#itemsTableBody tr");
    rows.forEach((r, idx) => {
        const slEl = r.querySelector(".row-sl");
        if (slEl) slEl.textContent = idx + 1;
    });
}

function calculateRow(inputElem) {
    const tr = inputElem.closest("tr");
    if (!tr) return;
    const qty = parseFloat(tr.querySelector(".item-qty")?.value) || 0;
    const rate = parseFloat(tr.querySelector(".item-rate")?.value) || 0;
    const amount = qty * rate;

    const amountInput = tr.querySelector(".item-amount");
    if (amountInput) {
        amountInput.value = amount.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    }
    recalculateTotals();
}

function recalculateTotals() {
    let subtotal = 0;
    let totalQty = 0;
    const rows = document.querySelectorAll("#itemsTableBody tr");

    rows.forEach(r => {
        const qty = parseFloat(r.querySelector(".item-qty")?.value) || 0;
        const rate = parseFloat(r.querySelector(".item-rate")?.value) || 0;
        subtotal += (qty * rate);
        totalQty += qty;
    });

    const vat = subtotal * 0.15;
    const grandTotal = subtotal + vat;

    const subtotalElem = document.getElementById("summarySubtotal");
    const vatElem = document.getElementById("summaryVat");
    const grandTotalElem = document.getElementById("summaryGrandTotal");

    if (subtotalElem) subtotalElem.textContent = "SAR " + subtotal.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    if (vatElem) vatElem.textContent = "SAR " + vat.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    if (grandTotalElem) grandTotalElem.textContent = "SAR " + grandTotal.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });

    const hidSubtotal = document.getElementById("hiddenSubtotal");
    const hidVat = document.getElementById("hiddenVat");
    const hidGrand = document.getElementById("hiddenGrandTotal");
    const hidItemsJson = document.getElementById("hiddenItemsJson");

    if (hidSubtotal) hidSubtotal.value = subtotal.toFixed(2);
    if (hidVat) hidVat.value = vat.toFixed(2);
    if (hidGrand) hidGrand.value = grandTotal.toFixed(2);

    const itemsList = [];
    rows.forEach((r, idx) => {
        const desc = r.querySelector(".item-desc")?.value.trim() || "";
        const unit = r.querySelector(".item-unit")?.value || "Pcs.";
        const qty = parseFloat(r.querySelector(".item-qty")?.value) || 0;
        const rate = parseFloat(r.querySelector(".item-rate")?.value) || 0;
        const amount = qty * rate;
        const remarks = r.querySelector(".item-remarks")?.value.trim() || "";

        if (desc) {
            itemsList.push({
                sl: idx + 1,
                desc: desc,
                unit: unit,
                qty: qty,
                rate: rate,
                amount: amount,
                remarks: remarks
            });
        }
    });

    if (hidItemsJson) hidItemsJson.value = JSON.stringify(itemsList);
    updateTableUI();
}

function updateTableUI() {
    const tbody = document.getElementById("itemsTableBody");
    const emptyState = document.getElementById("emptyTableState");
    const countBadge = document.getElementById("itemsCountBadge");
    const count = tbody ? tbody.children.length : 0;

    if (countBadge) {
        countBadge.textContent = `${count} ${count === 1 ? 'component' : 'components'}`;
    }

    if (emptyState) {
        if (count === 0) {
            emptyState.classList.remove("d-none");
        } else {
            emptyState.classList.add("d-none");
        }
    }
}

// --- IN-CELL AUTOCOMPLETE ---
function showSuggestions(input) {
    const val = input.value.toLowerCase().trim();
    const container = input.parentElement.querySelector(".autocomplete-suggestions");
    if (!container) return;

    if (!val || val.length < 1) {
        container.innerHTML = "";
        container.classList.add("d-none");
        return;
    }

    const matches = catalogItems.filter(item => 
        item.item_name.toLowerCase().includes(val) ||
        item.category.toLowerCase().includes(val)
    ).slice(0, 8);

    if (matches.length === 0) {
        container.classList.add("d-none");
        return;
    }

    container.innerHTML = "";
    container.classList.remove("d-none");

    matches.forEach(m => {
        const itemDiv = document.createElement("div");
        itemDiv.className = "autocomplete-suggestion";
        itemDiv.innerHTML = `
            <div>
                <span class="fw-semibold text-dark">${escapeHtml(m.item_name)}</span>
                <span class="badge bg-secondary-subtle text-secondary ms-1" style="font-size: 0.68rem;">${escapeHtml(m.category)}</span>
            </div>
            <span class="text-dark small fw-bold">SAR ${parseFloat(m.unit_price).toFixed(2)}</span>
        `;
        itemDiv.onclick = function () {
            selectCatalogItem(input, m);
        };
        container.appendChild(itemDiv);
    });
}

function selectCatalogItem(input, item) {
    input.value = item.item_name;
    const tr = input.closest("tr");
    if (tr) {
        const unitSelect = tr.querySelector(".item-unit");
        const rateInput = tr.querySelector(".item-rate");
        if (unitSelect) unitSelect.value = item.unit || "Pcs.";
        if (rateInput) rateInput.value = parseFloat(item.unit_price || 0).toFixed(2);
        calculateRow(tr.querySelector(".item-qty"));
    }
    
    const container = input.parentElement.querySelector(".autocomplete-suggestions");
    if (container) {
        container.innerHTML = "";
        container.classList.add("d-none");
    }
}

function attachGlobalListeners() {
    document.addEventListener("click", function (e) {
        // Close in-cell suggestions if click outside
        if (!e.target.closest(".position-relative")) {
            document.querySelectorAll(".autocomplete-suggestions").forEach(el => {
                el.classList.add("d-none");
            });
        }
        // Close top quick search if click outside
        if (!e.target.closest("#quickProductSearch") && !e.target.closest("#quickSearchResults")) {
            const qs = document.getElementById("quickSearchResults");
            if (qs) qs.classList.add("d-none");
        }
    });

    document.addEventListener("keydown", function (e) {
        if (e.key === "Escape") {
            const qs = document.getElementById("quickSearchResults");
            if (qs) qs.classList.add("d-none");
            document.querySelectorAll(".autocomplete-suggestions").forEach(el => {
                el.classList.add("d-none");
            });
        }
    });
}

function escapeHtml(text) {
    if (!text) return "";
    return text.toString()
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#039;");
}
