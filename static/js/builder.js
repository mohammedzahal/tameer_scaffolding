// TAMEER Scaffolding Quotation Builder JavaScript

let catalogItems = [];

document.addEventListener("DOMContentLoaded", function () {
    fetch('/api/items')
        .then(response => response.json())
        .then(data => {
            catalogItems = data;
            initCategoryDropdowns();
        })
        .catch(err => console.error("Error loading scaffolding catalog:", err));

    recalculateTotals();
    attachRowListeners();
});

function initCategoryDropdowns() {
    const categories = [...new Set(catalogItems.map(item => item.category))];
    const catSelect = document.getElementById("catalogCategorySelect");
    if (catSelect) {
        catSelect.innerHTML = '<option value="">All Categories</option>';
        categories.forEach(cat => {
            const opt = document.createElement("option");
            opt.value = cat;
            opt.textContent = cat;
            catSelect.appendChild(opt);
        });
    }
}

function filterCatalog() {
    const selectedCat = document.getElementById("catalogCategorySelect")?.value || "";
    const searchVal = (document.getElementById("catalogSearchInput")?.value || "").toLowerCase().trim();
    const container = document.getElementById("catalogItemsList");
    if (!container) return;

    container.innerHTML = "";

    const filtered = catalogItems.filter(item => {
        const matchesCat = !selectedCat || item.category === selectedCat;
        const matchesSearch = !searchVal || item.item_name.toLowerCase().includes(searchVal);
        return matchesCat && matchesSearch;
    });

    if (filtered.length === 0) {
        container.innerHTML = '<div class="p-3 text-muted text-center small">No components found.</div>';
        return;
    }

    filtered.forEach(item => {
        const div = document.createElement("div");
        div.className = "list-group-item d-flex justify-content-between align-items-center py-3 px-3 border-0 border-bottom";
        div.innerHTML = `
            <div>
                <span class="fw-bold text-dark" style="font-size: 1.05rem;">${item.item_name}</span>
                <span class="text-secondary d-block fw-medium" style="font-size: 0.88rem;">${item.category} • ${item.unit}</span>
            </div>
            <div class="d-flex align-items-center gap-3">
                <span class="text-dark fw-bold" style="font-size: 1.05rem;">SAR ${parseFloat(item.unit_price).toFixed(2)}</span>
                <button type="button" class="btn btn-secondary-tmr btn-sm px-3" onclick="addItemFromCatalog(${item.item_id})">
                    <i class="bi bi-plus-lg"></i> Add
                </button>
            </div>
        `;
        container.appendChild(div);
    });
}

function addItemFromCatalog(itemId) {
    const item = catalogItems.find(i => i.item_id === itemId);
    if (!item) return;

    addRow({
        desc: item.item_name,
        unit: item.unit || 'Pcs.',
        qty: 1,
        rate: item.unit_price || 0.0,
        remarks: ''
    });

    const modalEl = document.getElementById('catalogModal');
    if (modalEl) {
        const modal = bootstrap.Modal.getInstance(modalEl);
        if (modal) modal.hide();
    }
}

function addEmptyRow() {
    addRow({
        desc: '',
        unit: 'Pcs.',
        qty: 1,
        rate: 0.0,
        remarks: ''
    });
}

function addRow(data = {}) {
    const tbody = document.getElementById("itemsTableBody");
    const rowIndex = tbody.children.length + 1;

    const tr = document.createElement("tr");
    tr.className = "item-row";
    tr.innerHTML = `
        <td class="text-center fw-bold row-sl">${rowIndex}</td>
        <td>
            <div class="position-relative">
                <input type="text" class="form-control item-desc" value="${escapeHtml(data.desc || '')}" 
                       placeholder="Search component or type custom..." autocomplete="off" oninput="showSuggestions(this)">
                <div class="autocomplete-suggestions d-none"></div>
            </div>
        </td>
        <td>
            <select class="form-select item-unit text-center fw-semibold">
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
            <input type="text" class="form-control text-end fw-bold item-amount bg-light" value="0.00" readonly>
        </td>
        <td>
            <input type="text" class="form-control item-remarks" value="${escapeHtml(data.remarks || '')}" placeholder="Optional note">
        </td>
        <td class="text-center">
            <button type="button" class="icon-action-btn danger" onclick="removeRow(this)" title="Delete row">
                <i class="bi bi-x-lg"></i>
            </button>
        </td>
    `;

    tbody.appendChild(tr);
    calculateRow(tr.querySelector('.item-qty'));
    attachRowListeners();
}

function removeRow(btn) {
    const row = btn.closest("tr");
    const tbody = document.getElementById("itemsTableBody");
    if (tbody.children.length <= 1) {
        alert("Quotation must have at least one line item.");
        return;
    }
    row.remove();
    renumberRows();
    recalculateTotals();
}

function renumberRows() {
    const rows = document.querySelectorAll("#itemsTableBody tr");
    rows.forEach((r, idx) => {
        r.querySelector(".row-sl").textContent = idx + 1;
    });
}

function calculateRow(inputElem) {
    const tr = inputElem.closest("tr");
    const qty = parseFloat(tr.querySelector(".item-qty").value) || 0;
    const rate = parseFloat(tr.querySelector(".item-rate").value) || 0;
    const amount = qty * rate;

    tr.querySelector(".item-amount").value = amount.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    recalculateTotals();
}

function recalculateTotals() {
    let subtotal = 0;
    const rows = document.querySelectorAll("#itemsTableBody tr");
    rows.forEach(r => {
        const qty = parseFloat(r.querySelector(".item-qty")?.value) || 0;
        const rate = parseFloat(r.querySelector(".item-rate")?.value) || 0;
        subtotal += (qty * rate);
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
}

function showSuggestions(input) {
    const val = input.value.toLowerCase().trim();
    const container = input.parentElement.querySelector(".autocomplete-suggestions");
    if (!container) return;

    if (!val || val.length < 1) {
        container.innerHTML = "";
        container.classList.add("d-none");
        return;
    }

    const matches = catalogItems.filter(item => item.item_name.toLowerCase().includes(val)).slice(0, 7);
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
            <span><strong class="text-dark">${m.item_name}</strong> <span class="text-secondary fw-normal">(${m.category})</span></span>
            <span class="text-dark fw-bold" style="font-size: 1rem;">SAR ${parseFloat(m.unit_price).toFixed(2)}</span>
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
    tr.querySelector(".item-unit").value = item.unit || "Pcs.";
    tr.querySelector(".item-rate").value = parseFloat(item.unit_price || 0).toFixed(2);
    
    const container = input.parentElement.querySelector(".autocomplete-suggestions");
    if (container) {
        container.innerHTML = "";
        container.classList.add("d-none");
    }

    calculateRow(tr.querySelector(".item-qty"));
}

function attachRowListeners() {
    document.addEventListener("click", function (e) {
        if (!e.target.closest(".position-relative")) {
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
