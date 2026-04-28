/* ==============================================================
   AccuBooks — Accounting System Logic
   Vanilla JavaScript · localStorage persistence · CSV export
   ============================================================== */

// ─────────────────────────────────────────────────────────────
// 1. DATA LAYER — localStorage helpers
// ─────────────────────────────────────────────────────────────

/**
 * Read a JSON array from localStorage.
 * Returns [] if the key doesn't exist or is corrupted.
 */
function loadData(key) {
  try {
    return JSON.parse(localStorage.getItem(key)) || [];
  } catch {
    return [];
  }
}

/** Write a JSON-serialisable value to localStorage. */
function saveData(key, data) {
  localStorage.setItem(key, JSON.stringify(data));
}

// Central data stores (loaded once on init)
let accounts = loadData('acc_accounts');
let journalRows = loadData('acc_journal');

// Seed from defaults.js if empty
if (accounts.length === 0 && window.DEFAULT_ACCOUNTS) {
  accounts = window.DEFAULT_ACCOUNTS;
  saveData('acc_accounts', accounts);
}
if (journalRows.length === 0 && window.DEFAULT_JOURNAL) {
  journalRows = window.DEFAULT_JOURNAL;
  saveData('acc_journal', journalRows);
}

// ─────────────────────────────────────────────────────────────
// 1b. CODE-PREFIX HIERARCHY HELPERS
// ─────────────────────────────────────────────────────────────
// Hierarchy is inferred from account codes:
//   "1"        → top-level
//   "101"      → child of "1"
//   "101001"   → child of "101"
//   "10100101" → child of "101001"
// A is parent of B when B.id starts with A.id AND A.id !== B.id,
// and no other account C exists where B starts with C and C starts with A
// (i.e., A must be the *immediate* prefix-parent).

// Track which tree nodes are expanded (by account id)
let expandedNodes = new Set(loadData('acc_expanded'));
function saveExpanded() { saveData('acc_expanded', [...expandedNodes]); }

/**
 * Find the immediate parent of an account by code prefix.
 * Returns the account object whose id is the longest prefix of `code`,
 * or null if none exists (top-level account).
 */
function findCodeParent(code) {
  let bestParent = null;
  let bestLen = 0;
  accounts.forEach(a => {
    // a.id must be a strict prefix of code (shorter and code starts with it)
    if (a.id !== code && code.startsWith(a.id) && a.id.length > bestLen) {
      bestLen = a.id.length;
      bestParent = a;
    }
  });
  return bestParent;
}

/**
 * Get the parent account ID for a given code (prefix-based).
 * Returns the id string or null.
 */
function getParentId(code) {
  const p = findCodeParent(code);
  return p ? p.id : null;
}

/** Get direct children of a given account code (prefix-based). */
function getChildren(parentCode) {
  if (parentCode === null) {
    // Root-level: accounts that have no prefix-parent
    return accounts.filter(a => findCodeParent(a.id) === null);
  }
  return accounts.filter(a => {
    const p = findCodeParent(a.id);
    return p && p.id === parentCode;
  });
}

/** Check if an account code has any children (prefix-based). */
function hasChildren(code) {
  return accounts.some(a => {
    const p = findCodeParent(a.id);
    return p && p.id === code;
  });
}

/** Compute depth level of an account in the tree (prefix-based). */
function getDepth(code) {
  let depth = 0;
  let current = code;
  let steps = 0;
  while (steps < 20) {
    const parent = findCodeParent(current);
    if (!parent) break;
    depth++;
    current = parent.id;
    steps++;
  }
  return depth;
}

/** Get all descendant IDs (recursive, prefix-based). */
function getDescendantIds(code) {
  // All accounts whose code starts with this code (and isn't equal)
  return accounts.filter(a => a.id !== code && a.id.startsWith(code)).map(a => a.id);
}

/**
 * Build a flat ordered list from the code-prefix tree for rendering.
 * Each entry: { account, depth, hasKids }
 * @param {Function} [filterFn] - optional filter applied to each account
 */
function buildFlatTree(filterFn) {
  const result = [];
  const allAccounts = filterFn ? accounts.filter(filterFn) : [...accounts];

  function walk(parentCode, depth) {
    // Find children of parentCode among allAccounts
    let children;
    if (parentCode === null) {
      // Root: accounts in allAccounts that have no prefix-parent in allAccounts
      children = allAccounts.filter(a => {
        // check if any OTHER account in allAccounts is a prefix of a.id
        return !allAccounts.some(p => p.id !== a.id && a.id.startsWith(p.id));
      });
    } else {
      // Children: accounts whose immediate prefix-parent (among allAccounts) is parentCode
      children = allAccounts.filter(a => {
        if (a.id === parentCode || !a.id.startsWith(parentCode)) return false;
        // Check that parentCode is the LONGEST prefix in allAccounts for a.id
        let bestLen = 0;
        allAccounts.forEach(p => {
          if (p.id !== a.id && a.id.startsWith(p.id) && p.id.length > bestLen) {
            bestLen = p.id.length;
          }
        });
        return bestLen === parentCode.length;
      });
    }

    children.sort((a, b) => parseInt(a.id, 10) - parseInt(b.id, 10));

    children.forEach(acc => {
      const kids = allAccounts.some(a => a.id !== acc.id && a.id.startsWith(acc.id));
      result.push({ account: acc, depth, hasKids: kids });
      walk(acc.id, depth + 1);
    });
  }

  walk(null, 0);
  return result;
}

/**
 * Determine the visual group for an account based on its code prefix.
 * Uses the first character of the account's topmost ancestor code.
 * This ensures correct grouping regardless of the stored type field.
 */
function getCodeGroup(acc) {
  // Walk to the root ancestor in the full accounts list
  let rootCode = acc.id;
  let steps = 0;
  while (steps < 20) {
    const parent = findCodeParent(rootCode);
    if (!parent) break;
    rootCode = parent.id;
    steps++;
  }
  const first = String(rootCode).charAt(0);
  switch (first) {
    case '1': return 'Asset';
    case '2': return 'LiabEquity';
    case '3': return 'Expense';
    case '4': return 'Revenue';
    default:  return acc.type || 'Asset'; // fallback to stored type
  }
}

/**
 * Build a flat tree that groups accounts into the standard categories:
 *   1. Assets
 *   2. Liabilities & Equity (virtual combined group)
 *   3. Revenue
 *   4. Expense
 * Grouping is based on code prefix, NOT the stored type field.
 * Used for the Accounts Codes tree view.
 */
function buildGroupedFlatTree(filterFn) {
  const result = [];
  const allAccounts = filterFn ? accounts.filter(filterFn) : [...accounts];

  // Separate accounts into groups based on code prefix
  const assetAccounts    = allAccounts.filter(a => getCodeGroup(a) === 'Asset');
  const liabEqAccounts   = allAccounts.filter(a => getCodeGroup(a) === 'LiabEquity');
  const revenueAccounts  = allAccounts.filter(a => getCodeGroup(a) === 'Revenue');
  const expenseAccounts  = allAccounts.filter(a => getCodeGroup(a) === 'Expense');

  function walkSubset(subset, parentCode, depth) {
    let children;
    if (parentCode === null) {
      children = subset.filter(a => {
        return !subset.some(p => p.id !== a.id && a.id.startsWith(p.id));
      });
    } else {
      children = subset.filter(a => {
        if (a.id === parentCode || !a.id.startsWith(parentCode)) return false;
        let bestLen = 0;
        subset.forEach(p => {
          if (p.id !== a.id && a.id.startsWith(p.id) && p.id.length > bestLen) {
            bestLen = p.id.length;
          }
        });
        return bestLen === parentCode.length;
      });
    }

    children.sort((a, b) => parseInt(a.id, 10) - parseInt(b.id, 10));

    children.forEach(acc => {
      const kids = subset.some(a => a.id !== acc.id && a.id.startsWith(acc.id));
      result.push({ account: acc, depth, hasKids: kids });
      walkSubset(subset, acc.id, depth + 1);
    });
  }

  // 1. Assets (1xxx)
  walkSubset(assetAccounts, null, 0);

  // 2. Liabilities & Equity (2xxx)
  walkSubset(liabEqAccounts, null, 0);

  // 3. Expenses (3xxx)
  walkSubset(expenseAccounts, null, 0);

  // 4. Revenue (4xxx)
  walkSubset(revenueAccounts, null, 0);

  return result;
}

// ─────────────────────────────────────────────────────────────
// 2. NAVIGATION (SPA page switching)
// ─────────────────────────────────────────────────────────────

const navLinks = document.querySelectorAll('.nav-link');
const pages = document.querySelectorAll('.page');
const mobileBtn = document.getElementById('mobile-menu-btn');
const navLinksUl = document.querySelector('.nav-links');

/** Switch visible page and update the active nav link. */
function switchPage(pageId) {
  pages.forEach(p => p.classList.remove('active'));
  navLinks.forEach(l => l.classList.remove('active'));

  document.getElementById('page-' + pageId).classList.add('active');
  document.querySelector(`[data-page="${pageId}"]`).classList.add('active');

  // Close mobile menu if open
  navLinksUl.classList.remove('open');

  // Refresh data-dependent pages each time they are viewed
  if (pageId === 'trial') renderTrialBalance();
  if (pageId === 'income') renderIncomeStatement();
  if (pageId === 'balance') renderBalanceSheet();
  if (pageId === 'journal') renderJournal();
}

navLinks.forEach(link => {
  link.addEventListener('click', e => {
    e.preventDefault();
    switchPage(link.dataset.page);
  });
});

// Mobile hamburger toggle
mobileBtn.addEventListener('click', () => {
  navLinksUl.classList.toggle('open');
});

// ─────────────────────────────────────────────────────────────
// 3. TOAST NOTIFICATIONS
// ─────────────────────────────────────────────────────────────

/**
 * Show a short toast notification.
 * @param {string} msg  Text to display.
 * @param {'success'|'error'|'info'} type  Visual style.
 */
function toast(msg, type = 'info') {
  const container = document.getElementById('toast-container');
  const el = document.createElement('div');
  el.className = 'toast ' + type;
  el.textContent = msg;
  container.appendChild(el);
  setTimeout(() => el.remove(), 3000);
}

// ─────────────────────────────────────────────────────────────
// 4. ACCOUNTS CODES PAGE
// ─────────────────────────────────────────────────────────────

const accountsTbody = document.getElementById('accounts-tbody');
const accountFormWrap = document.getElementById('account-form-wrapper');
const accountForm = document.getElementById('account-form');
const accountFormTitle = document.getElementById('account-form-title');
const accIdInput = document.getElementById('acc-id');
const accNameInput = document.getElementById('acc-name');
const accTypeInput = document.getElementById('acc-type');
const accountsEmpty = document.getElementById('accounts-empty');
const accountsSearch = document.getElementById('accounts-search');

let editingAccountIdx = -1; // -1 = adding new, >=0 = editing

/**
 * Check if a node (or any ancestor) is collapsed, meaning it should be hidden.
 * Uses code-prefix tree: walk up from the account's prefix-parent.
 */
function isAncestorCollapsed(code) {
  let current = code;
  let steps = 0;
  while (steps < 20) {
    const parent = findCodeParent(current);
    if (!parent) break;
    if (!expandedNodes.has(parent.id)) return true;
    current = parent.id;
    steps++;
  }
  return false;
}

/** Render the accounts table as a tree, applying optional search filter. */
function renderAccounts(filter = '') {
  accountsTbody.innerHTML = '';
  const lower = filter.toLowerCase();
  const isSearching = lower.length > 0;

  let tree;
  if (isSearching) {
    // When searching, show flat filtered list (no tree hierarchy)
    const filtered = accounts.filter(a =>
      a.id.toLowerCase().includes(lower) ||
      a.name.toLowerCase().includes(lower) ||
      a.type.toLowerCase().includes(lower)
    );
    tree = filtered.map(acc => ({
      account: acc,
      depth: getDepth(acc.id),
      hasKids: hasChildren(acc.id)
    }));
  } else {
    tree = buildGroupedFlatTree();
  }

  if (tree.length === 0) {
    accountsEmpty.classList.remove('hidden');
  } else {
    accountsEmpty.classList.add('hidden');
  }

  let rowNum = 0;
  tree.forEach(node => {
    const acc = node.account;
    const depth = node.depth;
    const kids = node.hasKids;
    const realIdx = accounts.indexOf(acc);

    // When not searching, hide children of collapsed parents
    if (!isSearching && depth > 0 && isAncestorCollapsed(acc.id)) return;

    rowNum++;
    const isExpanded = expandedNodes.has(acc.id);
    const isLeaf = !kids;

    // Build toggle or spacer
    let toggleHtml = '';
    if (kids) {
      toggleHtml = `<button class="tree-toggle ${isExpanded ? 'expanded' : ''}" onclick="toggleTreeNode('${acc.id}')">▶</button>`;
    } else if (depth > 0) {
      toggleHtml = '<span class="tree-connector"></span>';
    } else {
      toggleHtml = '<span class="tree-spacer"></span>';
    }

    // Depth-level class for Excel-like level shading
    const depthClass = `tree-depth-${Math.min(depth, 5)}`;

    const tr = document.createElement('tr');
    tr.className = `tree-level-${Math.min(depth, 5)} ${depthClass}${kids ? ' tree-parent-row' : ''}${isLeaf ? ' tree-leaf-row' : ''}`;
    tr.innerHTML = `
      <td class="row-number">${rowNum}</td>
      <td class="tree-code-cell" style="padding:10px 14px;font-weight:600">${esc(acc.id)}</td>
      <td class="tree-cell">
        ${toggleHtml}
        <span class="tree-acc-name">${esc(acc.name)}</span>
      </td>
      <td style="padding:10px 14px">
        <span class="badge badge-${getCodeGroup(acc)}">${getCodeGroup(acc) === 'LiabEquity' ? 'Liab & Equity' : getCodeGroup(acc)}</span>
      </td>
      <td>
        <div class="cell-actions">
          <button class="btn-icon edit" title="Edit" onclick="editAccount(${realIdx})">✏️</button>
          <button class="btn-icon danger" title="Delete" onclick="deleteAccount(${realIdx})">🗑️</button>
        </div>
      </td>`;
    accountsTbody.appendChild(tr);
  });
}

/** Toggle expand/collapse for a tree node. */
window.toggleTreeNode = function (accountId) {
  if (expandedNodes.has(accountId)) {
    expandedNodes.delete(accountId);
  } else {
    expandedNodes.add(accountId);
  }
  saveExpanded();
  renderAccounts(accountsSearch.value);
};

// Expand All / Collapse All buttons
document.getElementById('btn-expand-all').addEventListener('click', () => {
  accounts.forEach(a => { if (hasChildren(a.id)) expandedNodes.add(a.id); });
  saveExpanded();
  renderAccounts(accountsSearch.value);
});

document.getElementById('btn-collapse-all').addEventListener('click', () => {
  expandedNodes.clear();
  saveExpanded();
  renderAccounts(accountsSearch.value);
});

// Show / hide the form
document.getElementById('btn-add-account').addEventListener('click', () => {
  editingAccountIdx = -1;
  accountFormTitle.textContent = 'Add New Account';
  accountForm.reset();
  accIdInput.disabled = false;
  accountFormWrap.classList.remove('hidden');
  accIdInput.focus();
});

document.getElementById('btn-cancel-account').addEventListener('click', () => {
  accountFormWrap.classList.add('hidden');
  accountForm.reset();
});

/**
 * Resolve the UI type 'LiabEquity' to an actual internal type
 * ('Liability' or 'Equity') by inspecting the parent account's type.
 * If no parent is found or parent is also mixed, defaults to 'Liability'.
 */
function resolveAccountType(uiType, accountId) {
  if (uiType !== 'LiabEquity') return uiType;

  // Walk up the code-prefix hierarchy looking for a parent with a definitive type
  const parent = findCodeParent(accountId);
  if (parent) {
    if (parent.type === 'Equity') return 'Equity';
    if (parent.type === 'Liability') return 'Liability';
  }

  // Heuristic: check if accountId starts with common equity prefixes (e.g. '3')
  // based on existing accounts in the system
  const equityRoots = accounts.filter(a => a.type === 'Equity' && getDepth(a.id) === 0).map(a => a.id);
  const liabRoots = accounts.filter(a => a.type === 'Liability' && getDepth(a.id) === 0).map(a => a.id);

  for (const root of equityRoots) {
    if (accountId.startsWith(root)) return 'Equity';
  }
  for (const root of liabRoots) {
    if (accountId.startsWith(root)) return 'Liability';
  }

  // Default to Liability if no parent context can be determined
  return 'Liability';
}

// Save (add or update)
accountForm.addEventListener('submit', e => {
  e.preventDefault();
  const id = accIdInput.value.trim();
  const name = accNameInput.value.trim();
  const uiType = accTypeInput.value;

  if (!id || !name || !uiType) return toast('Please fill all fields.', 'error');

  // Resolve 'LiabEquity' → 'Liability' or 'Equity'
  const type = resolveAccountType(uiType, id);

  if (editingAccountIdx === -1) {
    // Check duplicate ID
    if (accounts.some(a => a.id === id)) {
      return toast('Account ID already exists.', 'error');
    }
    accounts.push({ id, name, type });
    toast('Account added!', 'success');
  } else {
    accounts[editingAccountIdx] = { id, name, type };
    toast('Account updated!', 'success');
  }

  saveData('acc_accounts', accounts);
  accountFormWrap.classList.add('hidden');
  accountForm.reset();
  renderAccounts(accountsSearch.value);
});

/** Populate form for editing an existing account. */
window.editAccount = function (idx) {
  editingAccountIdx = idx;
  const acc = accounts[idx];
  accountFormTitle.textContent = 'Edit Account';
  accIdInput.value = acc.id;
  accIdInput.disabled = true;
  accNameInput.value = acc.name;
  // Map internal Liability/Equity types back to the merged UI option
  accTypeInput.value = (acc.type === 'Liability' || acc.type === 'Equity') ? 'LiabEquity' : acc.type;
  accountFormWrap.classList.remove('hidden');
  accNameInput.focus();
};

/** Delete an account after confirmation. */
window.deleteAccount = function (idx) {
  const acc = accounts[idx];
  const childCount = getDescendantIds(acc.id).length;
  const msg = childCount > 0
    ? `Delete account "${acc.name}"? It has ${childCount} child account(s) that will become standalone.`
    : `Delete account "${acc.name}"?`;
  if (!confirm(msg)) return;

  accounts.splice(idx, 1);
  saveData('acc_accounts', accounts);
  renderAccounts(accountsSearch.value);
  toast('Account deleted.', 'info');
};

// Live search
accountsSearch.addEventListener('input', () => {
  renderAccounts(accountsSearch.value);
});

// ─────────────────────────────────────────────────────────────
// 4b. EXCEL IMPORT
// ─────────────────────────────────────────────────────────────

const excelFileInput = document.getElementById('excel-file-input');

/**
 * Infer the account type from its code using common chart-of-accounts
 * prefix conventions:
 *   1xxx = Asset
 *   2xxx = Liability
 *   3xxx = Equity
 *   4xxx = Revenue
 *   5xxx = Expense
 * Falls back to 'Asset' if unrecognised.
 */
function inferTypeFromCode(code) {
  const first = String(code).charAt(0);
  switch (first) {
    case '1': return 'Asset';
    case '2': return 'LiabEquity';
    case '3': return 'Expense';
    case '4': return 'Revenue';
    default:  return 'Asset';
  }
}

/**
 * Try to find the Account Code and Account Name columns in a header row.
 * Returns { codeCol, nameCol } indices, or null if not found.
 * Supports various common header labels.
 */
function detectColumns(headerRow) {
  if (!headerRow || headerRow.length === 0) return null;

  let codeCol = -1;
  let nameCol = -1;

  headerRow.forEach((cell, idx) => {
    const val = String(cell || '').toLowerCase().trim();
    // Detect code column
    if (codeCol === -1 && (
      val === 'account code' || val === 'code' || val === 'acc code' ||
      val === 'account_code' || val === 'accountcode' || val === 'acct code' ||
      val === 'account id' || val === 'acc id' || val === 'id' ||
      val === 'account no' || val === 'account number' || val === 'no' || val === 'no.'
    )) {
      codeCol = idx;
    }
    // Detect name column
    if (nameCol === -1 && (
      val === 'account name' || val === 'name' || val === 'acc name' ||
      val === 'account_name' || val === 'accountname' || val === 'acct name' ||
      val === 'description' || val === 'account description' || val === 'account'
    )) {
      nameCol = idx;
    }
  });

  if (codeCol !== -1 && nameCol !== -1) return { codeCol, nameCol };
  return null;
}

// Import Excel button → trigger file picker
document.getElementById('btn-import-excel').addEventListener('click', () => {
  excelFileInput.value = ''; // reset so same file can be re-selected
  excelFileInput.click();
});

// Handle file selection
excelFileInput.addEventListener('change', (e) => {
  const file = e.target.files[0];
  if (!file) return;

  // Validate file type
  const validExts = ['.xlsx', '.xls'];
  const ext = file.name.substring(file.name.lastIndexOf('.')).toLowerCase();
  if (!validExts.includes(ext)) {
    toast('Please select an Excel file (.xlsx or .xls)', 'error');
    return;
  }

  const reader = new FileReader();
  reader.onload = function (evt) {
    try {
      const data = new Uint8Array(evt.target.result);
      const workbook = XLSX.read(data, { type: 'array' });

      // Use the first sheet
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];

      // Convert to array of arrays
      const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

      if (rows.length < 2) {
        toast('Excel file appears to be empty.', 'error');
        return;
      }

      // Try to auto-detect column positions from the header row
      let codeCol = 0;
      let nameCol = 1;
      const detected = detectColumns(rows[0]);
      let startRow = 0;

      if (detected) {
        codeCol = detected.codeCol;
        nameCol = detected.nameCol;
        startRow = 1; // skip header row
      } else {
        // No recognised header — check if row 0 looks like data
        // (first cell is a number-like string)
        const firstCell = String(rows[0][0] || '').trim();
        if (/^\d+$/.test(firstCell)) {
          startRow = 0; // no header, start from row 0
        } else {
          startRow = 1; // assume first row is a header even if unrecognised
        }
      }

      // Parse rows into accounts
      const imported = [];
      const seenCodes = new Set();
      let skipped = 0;

      for (let i = startRow; i < rows.length; i++) {
        const row = rows[i];
        const rawCode = row[codeCol];
        const rawName = row[nameCol];

        // Normalise code: remove whitespace, convert numbers to string
        const code = String(rawCode ?? '').trim();
        const name = String(rawName ?? '').trim();

        // Skip empty rows
        if (!code || !name) {
          skipped++;
          continue;
        }

        // Skip duplicates
        if (seenCodes.has(code)) {
          skipped++;
          continue;
        }

        seenCodes.add(code);
        imported.push({
          id: code,
          name: name,
          type: inferTypeFromCode(code)
        });
      }

      if (imported.length === 0) {
        toast('No valid accounts found in the file.', 'error');
        return;
      }

      // Confirm replacement
      const msg = `Import ${imported.length} accounts from "${file.name}"?\n` +
        (skipped > 0 ? `(${skipped} empty/duplicate rows skipped)\n` : '') +
        'This will REPLACE all existing accounts.';
      if (!confirm(msg)) return;

      // Replace accounts
      accounts = imported;
      saveData('acc_accounts', accounts);

      // Auto-expand top-level nodes
      expandedNodes.clear();
      expandedNodes.add('__liab_equity__');
      accounts.forEach(a => {
        if (getDepth(a.id) === 0 && hasChildren(a.id)) {
          expandedNodes.add(a.id);
        }
      });
      saveExpanded();

      // Re-render
      renderAccounts();
      toast(`✓ Imported ${imported.length} accounts successfully!`, 'success');

    } catch (err) {
      console.error('Excel import error:', err);
      toast('Failed to parse Excel file. Please check the format.', 'error');
    }
  };

  reader.readAsArrayBuffer(file);
});


// ─────────────────────────────────────────────────────────────
// 5. JOURNAL ENTRY PAGE
// ─────────────────────────────────────────────────────────────

const journalTbody = document.getElementById('journal-tbody');
const totalDebitEl = document.getElementById('total-debit');
const totalCreditEl = document.getElementById('total-credit');
const balanceInd = document.getElementById('balance-indicator');
const balanceIcon = document.getElementById('balance-icon');
const balanceText = document.getElementById('balance-text');
const journalEmpty = document.getElementById('journal-empty');
const journalSearch = document.getElementById('journal-search');

/**
 * Build a grouped <optgroup> + <option> list from the accounts array.
 * Groups: Assets | Liabilities & Equity (combined) | Revenue | Expenses
 * Hierarchy indentation is preserved within each group.
 * Used inside each journal row's Account <select>.
 */
function accountOptions(selectedId) {
  let html = '<option value="">Select…</option>';

  // Define the display groups and which account types belong to each
  const groups = [
    { label: 'Assets', types: ['Asset'] },
    { label: 'Liabilities & Equity', types: ['Liability', 'Equity'] },
    { label: 'Revenue', types: ['Revenue'] },
    { label: 'Expenses', types: ['Expense'] },
  ];

  groups.forEach(group => {
    // Build a subtree for only the account types in this group
    const groupTree = buildFlatTree(a => group.types.includes(a.type));
    if (groupTree.length === 0) return; // skip empty groups

    html += `<optgroup label="${group.label}">`;
    groupTree.forEach(node => {
      const a = node.account;
      const indent = '\u00A0\u00A0'.repeat(node.depth);
      const sel = a.id === selectedId ? 'selected' : '';
      html += `<option value="${a.id}" ${sel}>${indent}${a.id} — ${esc(a.name)}</option>`;
    });
    html += '</optgroup>';
  });

  return html;
}

/**
 * Get a human-readable account type label from an account code.
 * Uses the first character of the root ancestor to determine the type.
 *   1 → Asset
 *   2 → Liabilities & Equity
 *   3 → Expense
 *   4 → Revenue
 */
function getAccountTypeLabel(accountId) {
  if (!accountId) return '';
  const acc = accounts.find(a => a.id === accountId);
  if (!acc) return '';
  const group = getCodeGroup(acc);
  switch (group) {
    case 'Asset': return 'Asset';
    case 'LiabEquity': return 'Liabilities & Equity';
    case 'Revenue': return 'Revenue';
    case 'Expense': return 'Expense';
    default: return group;
  }
}

// ─────────────────────────────────────────────────────────────
// 5b. JOURNAL ENTRY — Auto-numbering & Grouping helpers
// ─────────────────────────────────────────────────────────────

/** Load the next journal number counter from localStorage. */
let _journalCounter = parseInt(localStorage.getItem('acc_journal_counter') || '0', 10);

function saveJournalCounter() {
  localStorage.setItem('acc_journal_counter', String(_journalCounter));
}

/**
 * Format a journal number as JRN-001, JRN-002, etc.
 */
function formatJournalNo(n) {
  return 'JRN-' + String(n).padStart(3, '0');
}

/**
 * Get the next available journal number.
 * Looks at all existing rows to find the highest used number,
 * then returns max(counter, highest) + 1.
 */
function getNextJournalNo() {
  let maxUsed = _journalCounter;
  journalRows.forEach(r => {
    if (r.journalNo) {
      const m = r.journalNo.match(/^JRN-(\d+)$/);
      if (m) {
        const n = parseInt(m[1], 10);
        if (n > maxUsed) maxUsed = n;
      }
    }
  });
  return maxUsed + 1;
}

/**
 * Dynamically renumber all journal entries so they start from 1
 * and remain contiguous even after deletions.
 */
function renumberJournalEntries() {
  const seen = [];
  journalRows.forEach(r => {
    if (r.journalNo && !seen.includes(r.journalNo)) {
      seen.push(r.journalNo);
    }
  });

  const map = {};
  seen.forEach((oldNo, idx) => {
    map[oldNo] = formatJournalNo(idx + 1);
  });

  let changed = false;
  journalRows.forEach(r => {
    if (r.journalNo && r.journalNo !== map[r.journalNo]) {
      r.journalNo = map[r.journalNo];
      changed = true;
    }
  });

  _journalCounter = seen.length;
  saveJournalCounter();

  return changed;
}

/**
 * Compute per-entry totals grouped by journalNo.
 * Returns a Map: journalNo -> { totalDebit, totalCredit, isComplete }
 */
function computeEntryTotals() {
  const map = new Map();
  journalRows.forEach(r => {
    const jno = r.journalNo || '';
    if (!jno) return;
    if (!map.has(jno)) map.set(jno, { totalDebit: 0, totalCredit: 0 });
    const entry = map.get(jno);
    entry.totalDebit += parseFloat(r.debit) || 0;
    entry.totalCredit += parseFloat(r.credit) || 0;
  });
  // Mark completeness
  map.forEach((entry, jno) => {
    const hasValues = entry.totalDebit > 0 || entry.totalCredit > 0;
    entry.isComplete = hasValues && Math.abs(entry.totalDebit - entry.totalCredit) < 0.005;
  });
  return map;
}

/**
 * Determine the "active" (current) journal number.
 * This is the last entry that is NOT yet complete, or a new number
 * if all existing entries are complete.
 */
function getActiveJournalNo() {
  const entryTotals = computeEntryTotals();

  // Collect unique journal numbers in order of first appearance
  const seen = [];
  journalRows.forEach(r => {
    if (r.journalNo && !seen.includes(r.journalNo)) seen.push(r.journalNo);
  });

  // Find the last incomplete entry
  for (let i = seen.length - 1; i >= 0; i--) {
    const entry = entryTotals.get(seen[i]);
    if (entry && !entry.isComplete) return seen[i];
  }

  // All complete (or no rows) → generate next number
  const nextNum = getNextJournalNo();
  _journalCounter = nextNum;
  saveJournalCounter();
  return formatJournalNo(nextNum);
}

/**
 * Ensure all rows without a journalNo get assigned one.
 * Called on load and when migrating old data.
 */
function ensureJournalNumbers() {
  let changed = false;
  journalRows.forEach(r => {
    if (!r.journalNo) {
      r.journalNo = getActiveJournalNo();
      changed = true;
    }
  });
  
  if (renumberJournalEntries()) {
    changed = true;
  }
  
  if (changed) saveData('acc_journal', journalRows);
}

// Run on init — assign journal numbers to any legacy rows and renumber
ensureJournalNumbers();

// ─────────────────────────────────────────────────────────────
// 5c. JOURNAL ENTRY — Render with grouping & visual separators
// ─────────────────────────────────────────────────────────────

/** Render all journal rows and recalculate totals. */
function renderJournal(filter = '') {
  journalTbody.innerHTML = '';
  const lower = filter.toLowerCase();

  // Build filtered view but keep original indices for mutations
  const entries = journalRows.map((r, i) => ({ ...r, _idx: i }));
  const filtered = entries.filter(r =>
    (r.account || '').toLowerCase().includes(lower) ||
    (r.description || '').toLowerCase().includes(lower) ||
    (r.date || '').includes(lower) ||
    (r.journalNo || '').toLowerCase().includes(lower) ||
    (r.costCenter || '').toLowerCase().includes(lower) ||
    (r.numerical || '').toLowerCase().includes(lower)
  );

  if (journalRows.length === 0) {
    journalEmpty.classList.remove('hidden');
  } else {
    journalEmpty.classList.add('hidden');
  }

  // Compute per-entry totals for visual separators
  const entryTotals = computeEntryTotals();

  filtered.forEach((row, vi) => {
    const idx = row._idx;
    const typeLabel = getAccountTypeLabel(row.account);
    const typeBadgeClass = row.account ? getCodeGroup(accounts.find(a => a.id === row.account) || {}) : '';
    const jno = row.journalNo || '';
    const entryInfo = entryTotals.get(jno);
    const entryComplete = entryInfo ? entryInfo.isComplete : false;

    // Determine if this is the last row of its entry in the filtered list
    const isLastOfEntry = (() => {
      for (let k = vi + 1; k < filtered.length; k++) {
        if ((filtered[k].journalNo || '') === jno) return false;
      }
      return true;
    })();

    const tr = document.createElement('tr');

    // Add entry-separator class on the last row of a complete entry
    if (isLastOfEntry && entryComplete) {
      tr.classList.add('jrn-entry-separator');
    }

    // Alternate entry background for readability
    // Color rows by even/odd journal number grouping
    const jnoIdx = (() => {
      const unique = [];
      for (const r of filtered) {
        const j = r.journalNo || '';
        if (j && !unique.includes(j)) unique.push(j);
      }
      return unique.indexOf(jno);
    })();
    if (jnoIdx % 2 === 1) {
      tr.classList.add('jrn-entry-alt');
    }

    tr.innerHTML = `
      <td class="row-number">${vi + 1}</td>
      <td class="jrn-no-cell">
        <span class="jrn-no-badge${entryComplete ? ' jrn-no-complete' : ''}">${esc(jno)}</span>
      </td>
      <td><input type="date" value="${row.date || ''}" data-idx="${idx}" data-field="date" /></td>
      <td><select data-idx="${idx}" data-field="account">${accountOptions(row.account)}</select></td>
      <td class="jrn-type-cell">
        <span class="jrn-type-badge ${typeBadgeClass ? 'jrn-type-' + typeBadgeClass : ''}" data-idx="${idx}" data-field="accountType">${typeLabel || '—'}</span>
      </td>
      <td><input type="text" value="${esc(row.costCenter || '')}" data-idx="${idx}" data-field="costCenter" placeholder="CC…" /></td>
      <td><input type="text" value="${esc(row.numerical || '')}" data-idx="${idx}" data-field="numerical" placeholder="Num…" /></td>
      <td><input type="text" value="${esc(row.description || '')}" data-idx="${idx}" data-field="description" placeholder="Description…" /></td>
      <td><input type="number" min="0" step="0.01" value="${row.debit || ''}" data-idx="${idx}" data-field="debit" placeholder="0.00" /></td>
      <td><input type="number" min="0" step="0.01" value="${row.credit || ''}" data-idx="${idx}" data-field="credit" placeholder="0.00" /></td>
      <td>
        <div class="cell-actions">
          <button class="btn-icon danger btn-delete-jrn" title="Delete row" data-delete-idx="${idx}">
            <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" fill="currentColor" viewBox="0 0 16 16">
              <path d="M5.5 5.5A.5.5 0 0 1 6 6v6a.5.5 0 0 1-1 0V6a.5.5 0 0 1 .5-.5m2.5 0a.5.5 0 0 1 .5.5v6a.5.5 0 0 1-1 0V6a.5.5 0 0 1 .5-.5m3 .5a.5.5 0 0 0-1 0v6a.5.5 0 0 0 1 0z"/>
              <path d="M14.5 3a1 1 0 0 1-1 1H13v9a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V4h-.5a1 1 0 0 1-1-1V2a1 1 0 0 1 1-1H6a1 1 0 0 1 1-1h2a1 1 0 0 1 1 1h3.5a1 1 0 0 1 1 1zM4.118 4 4 4.059V13a1 1 0 0 0 1 1h6a1 1 0 0 0 1-1V4.059L11.882 4zM2.5 3h11V2h-11z"/>
            </svg>
          </button>
        </div>
      </td>`;
    journalTbody.appendChild(tr);

    // If this is the last row of a complete entry, insert a per-entry subtotal row
    if (isLastOfEntry && entryComplete && entryInfo) {
      const subTr = document.createElement('tr');
      subTr.className = 'jrn-entry-subtotal';
      subTr.innerHTML = `
        <td colspan="8" class="jrn-subtotal-label">
          <span class="jrn-subtotal-badge">✓ ${esc(jno)}</span> Entry balanced
        </td>
        <td class="col-money jrn-subtotal-val">${formatMoney(entryInfo.totalDebit)}</td>
        <td class="col-money jrn-subtotal-val">${formatMoney(entryInfo.totalCredit)}</td>
        <td></td>`;
      journalTbody.appendChild(subTr);
    }
  });

  // Attach real-time change listeners
  journalTbody.querySelectorAll('input, select').forEach(el => {
    el.addEventListener('change', onJournalCellChange);
    // For number inputs, also listen on 'input' for real-time calc
    if (el.type === 'number') {
      el.addEventListener('input', recalcJournalTotals);
    }
  });

  recalcJournalTotals();
}

/** Handle cell change — persist immediately. */
function onJournalCellChange(e) {
  const idx = Number(e.target.dataset.idx);
  const field = e.target.dataset.field;
  let val = e.target.value;

  if (field === 'debit' || field === 'credit') {
    val = parseFloat(val) || 0;
  }

  journalRows[idx][field] = val;
  saveData('acc_journal', journalRows);

  // Auto-fill account type when account changes
  if (field === 'account') {
    const row = e.target.closest('tr');
    if (row) {
      const typeSpan = row.querySelector('[data-field="accountType"]');
      if (typeSpan) {
        const typeLabel = getAccountTypeLabel(val);
        const typeBadgeClass = val ? getCodeGroup(accounts.find(a => a.id === val) || {}) : '';
        typeSpan.textContent = typeLabel || '—';
        typeSpan.className = 'jrn-type-badge' + (typeBadgeClass ? ' jrn-type-' + typeBadgeClass : '');
      }
    }
  }

  // Re-render to update visual separators & entry status when debit/credit change
  if (field === 'debit' || field === 'credit') {
    renderJournal(journalSearch.value);
  } else {
    recalcJournalTotals();
  }
}

/** Recalculate total debit & credit, update balance indicator. */
function recalcJournalTotals() {
  let totalD = 0;
  let totalC = 0;

  // Read directly from inputs to capture unsaved typing
  journalTbody.querySelectorAll('input[data-field="debit"]').forEach(el => {
    totalD += parseFloat(el.value) || 0;
  });
  journalTbody.querySelectorAll('input[data-field="credit"]').forEach(el => {
    totalC += parseFloat(el.value) || 0;
  });

  totalDebitEl.textContent = formatMoney(totalD);
  totalCreditEl.textContent = formatMoney(totalC);

  const balanced = Math.abs(totalD - totalC) < 0.005;
  if (balanced) {
    balanceInd.className = 'balance-indicator balanced';
    balanceIcon.textContent = '✓';
    balanceText.textContent = 'Balanced — Debit equals Credit';
  } else {
    balanceInd.className = 'balance-indicator unbalanced';
    balanceIcon.textContent = '✗';
    const diff = Math.abs(totalD - totalC);
    balanceText.textContent = `Unbalanced — Difference: ${formatMoney(diff)}`;
  }
}

// Add Row button
document.getElementById('btn-add-row').addEventListener('click', () => {
  // Determine journal number: use active (incomplete) entry, or start new
  const activeJno = getActiveJournalNo();

  journalRows.push({
    journalNo: activeJno,
    date: new Date().toISOString().slice(0, 10),
    account: '',
    costCenter: '',
    numerical: '',
    description: '',
    debit: 0,
    credit: 0
  });
  saveData('acc_journal', journalRows);
  renderJournal(journalSearch.value);
  toast('Row added.', 'info');

  // Scroll to bottom of table and focus the date input in new row
  setTimeout(() => {
    // Find the last data row (skip subtotal rows)
    const allRows = journalTbody.querySelectorAll('tr:not(.jrn-entry-subtotal)');
    const lastRow = allRows[allRows.length - 1];
    if (lastRow) {
      lastRow.scrollIntoView({ behavior: 'smooth', block: 'center' });
      const firstInput = lastRow.querySelector('input[data-field="date"]');
      if (firstInput) firstInput.focus();
    }
  }, 50);
});

/**
 * Delete a journal row — event delegation on document.
 * Uses data-delete-idx attribute on the button to identify the row index.
 * Works reliably for dynamically added rows.
 */
document.addEventListener('click', function (e) {
  const btn = e.target.closest('.btn-delete-jrn');
  if (!btn) return;

  const idx = Number(btn.dataset.deleteIdx);
  if (isNaN(idx) || idx < 0 || idx >= journalRows.length) return;

  const row = journalRows[idx];
  if (!row) return;

  // Confirm when deleting the last row of a journal entry
  const jno = row.journalNo || '';
  const siblingsCount = journalRows.filter(r => r.journalNo === jno).length;
  if (siblingsCount === 1 && jno) {
    if (!confirm(`This is the only row in entry "${jno}". Delete this entire entry?`)) return;
  }

  // Find the parent <tr> for the fade-out animation
  const targetTr = btn.closest('tr');

  if (targetTr) {
    targetTr.classList.add('jrn-row-deleting');
    // Also fade the subtotal row if it sits right below
    const nextSibling = targetTr.nextElementSibling;
    if (nextSibling && nextSibling.classList.contains('jrn-entry-subtotal')) {
      nextSibling.classList.add('jrn-row-deleting');
    }

    setTimeout(() => {
      journalRows.splice(idx, 1);
      renumberJournalEntries();
      saveData('acc_journal', journalRows);
      renderJournal(journalSearch.value);
      toast('Row deleted.', 'info');
    }, 250);
  } else {
    // Fallback: immediate delete (no animation)
    journalRows.splice(idx, 1);
    renumberJournalEntries();
    saveData('acc_journal', journalRows);
    renderJournal(journalSearch.value);
    toast('Row deleted.', 'info');
  }
});

// Live search for journal
journalSearch.addEventListener('input', () => {
  renderJournal(journalSearch.value);
});

/**
 * Aggregate journal entries by account to produce movement figures.
 * Groups by account ID, sums debit and credit.
 * This is the RAW (leaf-level) aggregation — unchanged from original.
 */
function computeTrialBalance() {
  const map = {}; // accountId -> { debit, credit }

  journalRows.forEach(row => {
    if (!row.account) return;
    if (!map[row.account]) map[row.account] = { debit: 0, credit: 0 };
    map[row.account].debit += parseFloat(row.debit) || 0;
    map[row.account].credit += parseFloat(row.credit) || 0;
  });

  // Convert to array and attach account name
  return Object.keys(map).map(id => {
    const acc = accounts.find(a => a.id === id);
    return {
      id,
      name: acc ? acc.name : id,
      type: acc ? acc.type : '',
      debit: map[id].debit,
      credit: map[id].credit
    };
  });
}

/**
 * Enhanced: Compute hierarchical trial balance (code-prefix-based).
 * Child account values roll up (accumulate) into parent accounts.
 * Returns a map: accountId -> { debit, credit } (including rolled-up totals).
 */
function computeHierarchicalTB() {
  const rawTB = computeTrialBalance();
  const map = {}; // accountId -> { debit, credit }

  // Seed with raw values
  rawTB.forEach(r => {
    map[r.id] = { debit: r.debit, credit: r.credit };
  });

  // For each account with a raw balance, roll up through prefix-ancestors
  rawTB.forEach(r => {
    let current = r.id;
    let steps = 0;
    while (steps < 20) {
      const parent = findCodeParent(current);
      if (!parent) break;
      if (!map[parent.id]) map[parent.id] = { debit: 0, credit: 0 };
      map[parent.id].debit += r.debit;
      map[parent.id].credit += r.credit;
      current = parent.id;
      steps++;
    }
  });

  return map;
}

// ─────────────────────────────────────────────────────────────
// 6b. OPENING BALANCES — localStorage persistence
// ─────────────────────────────────────────────────────────────

/** Opening balances: { accountId: { debit: number, credit: number } } */
let openingBalances = {};
(function loadOpeningBalances() {
  const data = localStorage.getItem('acc_opening_balances');
  if (data) {
    try {
      openingBalances = JSON.parse(data);
    } catch { openingBalances = {}; }
  } else if (window.DEFAULT_OPENING_BALANCES) {
    openingBalances = window.DEFAULT_OPENING_BALANCES;
    localStorage.setItem('acc_opening_balances', JSON.stringify(openingBalances));
  }
})();

function saveOpeningBalances() {
  localStorage.setItem('acc_opening_balances', JSON.stringify(openingBalances));
}

/**
 * Compute hierarchical opening balances (roll up leaves → parents).
 * Returns a map: accountId -> { debit, credit }
 */
function computeHierarchicalOB() {
  const map = {};

  // Seed with stored opening balances
  Object.keys(openingBalances).forEach(id => {
    const ob = openingBalances[id];
    const d = parseFloat(ob.debit) || 0;
    const c = parseFloat(ob.credit) || 0;
    if (d !== 0 || c !== 0) {
      map[id] = { debit: d, credit: c };
    }
  });

  // Roll up to parents
  Object.keys(openingBalances).forEach(id => {
    const ob = openingBalances[id];
    const d = parseFloat(ob.debit) || 0;
    const c = parseFloat(ob.credit) || 0;
    if (d === 0 && c === 0) return;

    let current = id;
    let steps = 0;
    while (steps < 20) {
      const parent = findCodeParent(current);
      if (!parent) break;
      if (!map[parent.id]) map[parent.id] = { debit: 0, credit: 0 };
      map[parent.id].debit += d;
      map[parent.id].credit += c;
      current = parent.id;
      steps++;
    }
  });

  return map;
}

// ─────────────────────────────────────────────────────────────
// 6c. OPENING BALANCES — UI Panel
// ─────────────────────────────────────────────────────────────

document.getElementById('btn-manage-opening').addEventListener('click', () => {
  const panel = document.getElementById('opening-balance-panel');
  panel.classList.toggle('hidden');
  if (!panel.classList.contains('hidden')) {
    renderOpeningBalancesTable();
  }
});

document.getElementById('btn-close-opening').addEventListener('click', () => {
  document.getElementById('opening-balance-panel').classList.add('hidden');
});

document.getElementById('btn-save-opening').addEventListener('click', () => {
  // Read values from the OB table inputs
  const tbody = document.getElementById('ob-tbody');
  tbody.querySelectorAll('tr').forEach(tr => {
    const accId = tr.dataset.accId;
    if (!accId) return;
    const dInput = tr.querySelector('input[data-field="ob-debit"]');
    const cInput = tr.querySelector('input[data-field="ob-credit"]');
    const d = parseFloat(dInput?.value) || 0;
    const c = parseFloat(cInput?.value) || 0;
    if (d !== 0 || c !== 0) {
      openingBalances[accId] = { debit: d, credit: c };
    } else {
      delete openingBalances[accId];
    }
  });

  saveOpeningBalances();
  renderTrialBalance();
  toast('Opening balances saved!', 'success');
});

document.getElementById('ob-search').addEventListener('input', () => {
  renderOpeningBalancesTable(document.getElementById('ob-search').value);
});

function renderOpeningBalancesTable(filter = '') {
  const tbody = document.getElementById('ob-tbody');
  tbody.innerHTML = '';
  const lower = filter.toLowerCase();

  // Show leaf accounts (accounts without children) for entering opening balances
  const leafAccounts = accounts.filter(a => {
    const isLeaf = !hasChildren(a.id);
    if (!isLeaf) return false;
    if (lower) {
      return a.id.toLowerCase().includes(lower) || a.name.toLowerCase().includes(lower);
    }
    return true;
  }).sort((a, b) => parseInt(a.id, 10) - parseInt(b.id, 10));

  leafAccounts.forEach(acc => {
    const ob = openingBalances[acc.id] || { debit: 0, credit: 0 };
    const tr = document.createElement('tr');
    tr.dataset.accId = acc.id;
    tr.innerHTML = `
      <td style="padding:8px 14px;font-weight:600;font-variant-numeric:tabular-nums">${esc(acc.id)}</td>
      <td style="padding:8px 14px">${esc(acc.name)}</td>
      <td><input type="number" min="0" step="0.01" value="${ob.debit || ''}" data-field="ob-debit" placeholder="0.00" style="text-align:right;padding:8px 14px" /></td>
      <td><input type="number" min="0" step="0.01" value="${ob.credit || ''}" data-field="ob-credit" placeholder="0.00" style="text-align:right;padding:8px 14px" /></td>`;
    tbody.appendChild(tr);
  });
}

// ─────────────────────────────────────────────────────────────
// 6d. FULL TRIAL BALANCE — Render with 4 column groups
// ─────────────────────────────────────────────────────────────

function renderTrialBalance() {
  const tbody = document.getElementById('trial-tbody');
  tbody.innerHTML = '';

  const hierMovement = computeHierarchicalTB();   // Movement (from journal entries)
  const hierBeginning = computeHierarchicalOB();   // Beginning (opening balances)

  // Determine which accounts have ANY data (beginning or movement)
  const allIds = new Set([...Object.keys(hierMovement), ...Object.keys(hierBeginning)]);

  // Also include accounts that are ancestors of those with data
  accounts.forEach(a => {
    if (allIds.has(a.id)) {
      // Walk up to ancestors and add them
      let current = a.id;
      let steps = 0;
      while (steps < 20) {
        const parent = findCodeParent(current);
        if (!parent) break;
        allIds.add(parent.id);
        current = parent.id;
        steps++;
      }
    }
  });

  const emptyEl = document.getElementById('trial-empty');
  if (allIds.size === 0) {
    emptyEl.classList.remove('hidden');
  } else {
    emptyEl.classList.add('hidden');
  }

  // Build a full tree of all accounts that have data
  const accountsWithData = new Set(allIds);

  // Build tree using existing flat-tree builder, filtering to only accounts with data
  const fullTree = buildFlatTree(a => accountsWithData.has(a.id));

  // Grand totals (computed from leaf accounts only to avoid double-counting)
  let gtBegD = 0, gtBegC = 0, gtMovD = 0, gtMovC = 0;
  let gtTotD = 0, gtTotC = 0, gtBalD = 0, gtBalC = 0;

  let rowNum = 0;

  fullTree.forEach(node => {
    const acc = node.account;
    const beg = hierBeginning[acc.id] || { debit: 0, credit: 0 };
    const mov = hierMovement[acc.id] || { debit: 0, credit: 0 };

    // Total = Beginning + Movement
    const totD = beg.debit + mov.debit;
    const totC = beg.credit + mov.credit;

    // Balance: net = totalDebit - totalCredit
    const net = totD - totC;
    const balD = net > 0 ? net : 0;
    const balC = net < 0 ? Math.abs(net) : 0;

    // For grand totals, only count leaf-level to avoid double-counting
    const isLeaf = !node.hasKids;
    if (isLeaf) {
      gtBegD += beg.debit;
      gtBegC += beg.credit;
      gtMovD += mov.debit;
      gtMovC += mov.credit;
      gtTotD += totD;
      gtTotC += totC;
      gtBalD += balD;
      gtBalC += balC;
    }

    rowNum++;
    const isParent = node.hasKids;
    const depth = node.depth;

    const tr = document.createElement('tr');
    if (isParent) tr.className = 'report-tree-parent';

    tr.innerHTML = `
      <td class="row-number">${rowNum}</td>
      <td class="tb-cell-code" style="padding:8px 12px;font-weight:600">${esc(acc.id)}</td>
      <td class="tb-cell-name" style="padding:8px 12px;${depth > 0 ? 'padding-left:' + (12 + depth * 20) + 'px' : ''}${isParent ? ';font-weight:700' : ''}">${esc(acc.name)}</td>
      <td class="col-money tb-beg-col" style="padding:8px 10px">${formatMoney(beg.debit)}</td>
      <td class="col-money tb-beg-col" style="padding:8px 10px">${formatMoney(beg.credit)}</td>
      <td class="col-money tb-mov-col" style="padding:8px 10px">${formatMoney(mov.debit)}</td>
      <td class="col-money tb-mov-col" style="padding:8px 10px">${formatMoney(mov.credit)}</td>
      <td class="col-money tb-tot-col" style="padding:8px 10px">${formatMoney(totD)}</td>
      <td class="col-money tb-tot-col" style="padding:8px 10px">${formatMoney(totC)}</td>
      <td class="col-money tb-bal-col" style="padding:8px 10px;${balD > 0 ? 'color:#16a34a;font-weight:600' : ''}">${balD > 0 ? formatMoney(balD) : ''}</td>
      <td class="col-money tb-bal-col" style="padding:8px 10px;${balC > 0 ? 'color:#dc2626;font-weight:600' : ''}">${balC > 0 ? formatMoney(balC) : ''}</td>`;
    tbody.appendChild(tr);
  });

  // Update grand totals footer
  document.getElementById('tb-gt-beg-d').textContent = formatMoney(gtBegD);
  document.getElementById('tb-gt-beg-c').textContent = formatMoney(gtBegC);
  document.getElementById('tb-gt-mov-d').textContent = formatMoney(gtMovD);
  document.getElementById('tb-gt-mov-c').textContent = formatMoney(gtMovC);
  document.getElementById('tb-gt-tot-d').textContent = formatMoney(gtTotD);
  document.getElementById('tb-gt-tot-c').textContent = formatMoney(gtTotC);
  document.getElementById('tb-gt-bal-d').textContent = formatMoney(gtBalD);
  document.getElementById('tb-gt-bal-c').textContent = formatMoney(gtBalC);

  // Balance check card
  document.getElementById('tb-eq-debit').textContent = formatMoney(gtTotD);
  document.getElementById('tb-eq-credit').textContent = formatMoney(gtTotC);
  const eqStatus = document.getElementById('tb-eq-status');
  const balanced = Math.abs(gtTotD - gtTotC) < 0.005;
  if (balanced) {
    eqStatus.textContent = '✓ Balanced';
    eqStatus.className = 'eq-status balanced';
  } else {
    eqStatus.textContent = '✗ Unbalanced';
    eqStatus.className = 'eq-status unbalanced';
  }
}

// ─────────────────────────────────────────────────────────────
// 7. INCOME STATEMENT PAGE
// ─────────────────────────────────────────────────────────────

/**
 * Build Income Statement from trial balance data.
 * Enhanced with hierarchy: shows tree structure with rolled-up amounts.
 * Revenue accounts: net amount = credit - debit
 * Expense accounts: net amount = debit - credit
 */
function renderIncomeStatement() {
  const hierMap = computeHierarchicalTB();
  const rawTB = computeTrialBalance();

  // Get set of accounts with any balance
  const hasBalance = new Set(Object.keys(hierMap).filter(id => {
    const v = hierMap[id];
    return v.debit !== 0 || v.credit !== 0;
  }));

  const revTree = buildFlatTree(a => getCodeGroup(a) === 'Revenue' && hasBalance.has(a.id));
  const expTree = buildFlatTree(a => getCodeGroup(a) === 'Expense' && hasBalance.has(a.id));

  const revBody = document.getElementById('income-revenue-tbody');
  const expBody = document.getElementById('income-expense-tbody');
  revBody.innerHTML = '';
  expBody.innerHTML = '';

  function renderReportTree(tree, tbody, calcFn) {
    let leafTotal = 0;
    tree.forEach(node => {
      const vals = hierMap[node.account.id] || { debit: 0, credit: 0 };
      const amt = calcFn(vals);
      const isParent = node.hasKids;
      const indent = '\u00A0\u00A0\u00A0'.repeat(node.depth);

      // Only sum leaf-level raw entries for totals
      const raw = rawTB.find(r => r.id === node.account.id);
      if (raw) leafTotal += calcFn(raw);

      const tr = document.createElement('tr');
      if (isParent) tr.className = 'report-tree-parent';
      tr.innerHTML = `
        <td style="padding:10px 14px;${node.depth > 0 ? 'padding-left:' + (14 + node.depth * 18) + 'px' : ''}">${indent}${esc(node.account.name)}</td>
        <td class="col-money" style="padding:10px 14px">${formatMoney(amt)}</td>`;
      tbody.appendChild(tr);
    });
    return leafTotal;
  }

  const totalRev = renderReportTree(revTree, revBody, v => v.credit - v.debit);
  const totalExp = renderReportTree(expTree, expBody, v => v.debit - v.credit);

  document.getElementById('total-revenue').textContent = formatMoney(totalRev);
  document.getElementById('total-expenses').textContent = formatMoney(totalExp);

  const netIncome = totalRev - totalExp;
  const niCard = document.getElementById('net-income-card');
  document.getElementById('net-income-value').textContent = formatMoney(netIncome);
  niCard.classList.toggle('negative', netIncome < 0);
}

// ─────────────────────────────────────────────────────────────
// 8. BALANCE SHEET PAGE
// ─────────────────────────────────────────────────────────────

/**
 * Build Balance Sheet from trial balance.
 * Enhanced with hierarchy: shows tree structure with rolled-up amounts.
 * Assets:      net = debit - credit
 * Liabilities: net = credit - debit
 * Equity:      net = credit - debit  (+ retained earnings / net income)
 */
function renderBalanceSheet() {
  const hierMovement = computeHierarchicalTB();
  const hierBeginning = computeHierarchicalOB();

  // 1. Calculate Combined Balances (Trial Balance by balances)
  const allIds = new Set([...Object.keys(hierMovement), ...Object.keys(hierBeginning)]);
  const combinedMap = {};
  allIds.forEach(id => {
    const beg = hierBeginning[id] || { debit: 0, credit: 0 };
    const mov = hierMovement[id] || { debit: 0, credit: 0 };
    combinedMap[id] = { debit: beg.debit + mov.debit, credit: beg.credit + mov.credit };
  });

  const hasBalance = new Set(Object.keys(combinedMap).filter(id => {
    const v = combinedMap[id];
    return v.debit !== 0 || v.credit !== 0;
  }));

  // Net Income = Revenue - Expense
  let totalRev = 0;
  let totalExp = 0;
  accounts.forEach(a => {
    if (!hasChildren(a.id)) {
      const g = getCodeGroup(a);
      if (g === 'Revenue' || g === 'Expense') {
        const vals = combinedMap[a.id] || { debit: 0, credit: 0 };
        if (g === 'Revenue') totalRev += (vals.credit - vals.debit);
        if (g === 'Expense') totalExp += (vals.debit - vals.credit);
      }
    }
  });
  const netIncome = totalRev - totalExp;

  function fillSection(tbodyId, type, calcFn) {
    const tbody = document.getElementById(tbodyId);
    if (!tbody) return 0; // Guard in case element is missing
    tbody.innerHTML = '';
    const tree = buildFlatTree(a => getCodeGroup(a) === type && hasBalance.has(a.id));
    let leafTotal = 0;

    tree.forEach(node => {
      const vals = combinedMap[node.account.id] || { debit: 0, credit: 0 };
      const amt = calcFn(vals);
      const isParent = node.hasKids;
      const indent = '\u00A0\u00A0\u00A0'.repeat(node.depth);

      // Only sum leaf nodes to avoid double-counting
      if (!isParent) {
        leafTotal += amt;
      }

      const tr = document.createElement('tr');
      if (isParent) tr.className = 'report-tree-parent';
      tr.innerHTML = `
        <td style="padding:10px 14px;${node.depth > 0 ? 'padding-left:' + (14 + node.depth * 18) + 'px' : ''}">${indent}${esc(node.account.name)}</td>
        <td class="col-money" style="padding:10px 14px">${formatMoney(amt)}</td>`;
      tbody.appendChild(tr);
    });
    return leafTotal;
  }

  const totalAssets = fillSection('bs-assets-tbody', 'Asset', v => v.debit - v.credit);
  const totalLiabEq = fillSection('bs-liab-equity-tbody', 'LiabEquity', v => v.credit - v.debit);

  document.getElementById('total-assets').textContent = formatMoney(totalAssets);
  
  const elTotalLiabEq = document.getElementById('total-liab-equity');
  if (elTotalLiabEq) elTotalLiabEq.textContent = formatMoney(totalLiabEq);
  
  const elNetIncomeVal = document.getElementById('bs-net-income-val');
  if (elNetIncomeVal) elNetIncomeVal.textContent = formatMoney(netIncome);

  const totalRightSide = totalLiabEq + netIncome;

  // Combined Liabilities & Equity + Net Income total
  document.getElementById('total-liab-equity-combined').textContent = formatMoney(totalRightSide);

  // Equation card
  document.getElementById('eq-assets').textContent = formatMoney(totalAssets);
  document.getElementById('eq-liab-equity').textContent = formatMoney(totalRightSide);

  const eqStatus = document.getElementById('eq-status');
  const balanced = Math.abs(totalAssets - totalRightSide) < 0.005;
  if (balanced) {
    eqStatus.textContent = '✓ Balanced';
    eqStatus.className = 'eq-status balanced';
  } else {
    eqStatus.textContent = '✗ Unbalanced';
    eqStatus.className = 'eq-status unbalanced';
  }
}

// ─────────────────────────────────────────────────────────────
// 9. EXPORT TO CSV
// ─────────────────────────────────────────────────────────────

/**
 * Convert a 2D array to CSV string and trigger a file download.
 * @param {string[][]} rows  Array of arrays (each inner array = one row).
 * @param {string}     name  File name (without .csv).
 */
function downloadCSV(rows, name) {
  const csv = rows.map(r =>
    r.map(cell => `"${String(cell).replace(/"/g, '""')}"`).join(',')
  ).join('\n');

  const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = name + '.csv';
  a.click();
  URL.revokeObjectURL(url);
  toast('CSV exported!', 'success');
}

// Export buttons
document.getElementById('btn-export-accounts').addEventListener('click', () => {
  const rows = [['Account Code', 'Account Name', 'Type', 'Parent Code', 'Parent Name', 'Depth']];
  const tree = buildFlatTree();
  tree.forEach(node => {
    const a = node.account;
    const parent = findCodeParent(a.id);
    rows.push([a.id, a.name, a.type, parent ? parent.id : '', parent ? parent.name : '', node.depth]);
  });
  downloadCSV(rows, 'chart_of_accounts');
});

document.getElementById('btn-export-journal').addEventListener('click', () => {
  const rows = [['Journal No', 'Date', 'Account', 'Account Type', 'Cost Center', 'Numerical', 'Description', 'Debit', 'Credit']];
  journalRows.forEach(r => {
    const accName = accounts.find(a => a.id === r.account)?.name || r.account;
    const typeLabel = getAccountTypeLabel(r.account);
    rows.push([r.journalNo || '', r.date, accName, typeLabel, r.costCenter || '', r.numerical || '', r.description, r.debit, r.credit]);
  });
  downloadCSV(rows, 'journal_entries');
});

document.getElementById('btn-export-trial').addEventListener('click', () => {
  const hierMovement = computeHierarchicalTB();
  const hierBeginning = computeHierarchicalOB();
  const allIds = new Set([...Object.keys(hierMovement), ...Object.keys(hierBeginning)]);
  const fullTree = buildFlatTree(a => allIds.has(a.id));

  const rows = [['Account Code', 'Account Name',
    'Beginning Debit', 'Beginning Credit',
    'Movement Debit', 'Movement Credit',
    'Total Debit', 'Total Credit',
    'Balance Debit', 'Balance Credit']];

  fullTree.forEach(node => {
    const acc = node.account;
    const beg = hierBeginning[acc.id] || { debit: 0, credit: 0 };
    const mov = hierMovement[acc.id] || { debit: 0, credit: 0 };
    const totD = beg.debit + mov.debit;
    const totC = beg.credit + mov.credit;
    const net = totD - totC;
    const balD = net > 0 ? net : 0;
    const balC = net < 0 ? Math.abs(net) : 0;
    rows.push([acc.id, acc.name, beg.debit, beg.credit, mov.debit, mov.credit, totD, totC, balD, balC]);
  });
  downloadCSV(rows, 'trial_balance');
});

document.getElementById('btn-export-income').addEventListener('click', () => {
  const tb = computeTrialBalance();
  const rows = [['Category', 'Account Name', 'Amount']];
  tb.filter(r => getCodeGroup(r) === 'Revenue').forEach(r => rows.push(['Revenue', r.name, r.credit - r.debit]));
  tb.filter(r => getCodeGroup(r) === 'Expense').forEach(r => rows.push(['Expense', r.name, r.debit - r.credit]));
  const totalRev = tb.filter(r => getCodeGroup(r) === 'Revenue').reduce((s, r) => s + (r.credit - r.debit), 0);
  const totalExp = tb.filter(r => getCodeGroup(r) === 'Expense').reduce((s, r) => s + (r.debit - r.credit), 0);
  rows.push(['', 'Net Income', totalRev - totalExp]);
  downloadCSV(rows, 'income_statement');
});

document.getElementById('btn-export-balance').addEventListener('click', () => {
  const hierMovement = computeHierarchicalTB();
  const hierBeginning = computeHierarchicalOB();
  
  const allIds = new Set([...Object.keys(hierMovement), ...Object.keys(hierBeginning)]);
  const combinedMap = {};
  allIds.forEach(id => {
    const beg = hierBeginning[id] || { debit: 0, credit: 0 };
    const mov = hierMovement[id] || { debit: 0, credit: 0 };
    combinedMap[id] = { debit: beg.debit + mov.debit, credit: beg.credit + mov.credit };
  });

  const hasBalance = new Set(Object.keys(combinedMap).filter(id => {
    const v = combinedMap[id];
    return v.debit !== 0 || v.credit !== 0;
  }));

  const rows = [['Category', 'Account Name', 'Amount']];

  // Export Assets
  const assetTree = buildFlatTree(a => getCodeGroup(a) === 'Asset' && hasBalance.has(a.id));
  assetTree.forEach(node => {
    const vals = combinedMap[node.account.id] || { debit: 0, credit: 0 };
    rows.push(['Asset', node.account.name, vals.debit - vals.credit]);
  });

  // Export Liabilities & Equity
  const liabEqTree = buildFlatTree(a => getCodeGroup(a) === 'LiabEquity' && hasBalance.has(a.id));
  liabEqTree.forEach(node => {
    const vals = combinedMap[node.account.id] || { debit: 0, credit: 0 };
    rows.push(['Liabilities & Equity', node.account.name, vals.credit - vals.debit]);
  });
  
  let totalRev = 0;
  let totalExp = 0;
  accounts.forEach(a => {
    if (!hasChildren(a.id)) {
      const g = getCodeGroup(a);
      if (g === 'Revenue' || g === 'Expense') {
        const vals = combinedMap[a.id] || { debit: 0, credit: 0 };
        if (g === 'Revenue') totalRev += (vals.credit - vals.debit);
        if (g === 'Expense') totalExp += (vals.debit - vals.credit);
      }
    }
  });
  const netIncome = totalRev - totalExp;
  
  rows.push(['Equity', 'Net Income', netIncome]);

  downloadCSV(rows, 'balance_sheet');
});

document.getElementById('btn-export-defaults').addEventListener('click', () => {
  let content = "// This file contains the default data for the application when hosted.\n\n";
  content += "window.DEFAULT_ACCOUNTS = " + JSON.stringify(accounts, null, 2) + ";\n\n";
  content += "window.DEFAULT_OPENING_BALANCES = " + JSON.stringify(openingBalances, null, 2) + ";\n\n";
  content += "window.DEFAULT_JOURNAL = " + JSON.stringify(journalRows, null, 2) + ";\n";
  
  const blob = new Blob([content], { type: 'text/javascript' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = 'defaults.js';
  a.click();
  URL.revokeObjectURL(url);
  toast('defaults.js generated!', 'success');
});

// ─────────────────────────────────────────────────────────────
// 10. UTILITY FUNCTIONS
// ─────────────────────────────────────────────────────────────

/** Format a number as money with 2 decimal places and comma separators. */
function formatMoney(n) {
  return Number(n).toLocaleString('en-US', {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  });
}

/** Escape HTML special characters to prevent XSS. */
function esc(str) {
  const div = document.createElement('div');
  div.textContent = str;
  return div.innerHTML;
}

// ─────────────────────────────────────────────────────────────
// 11. INITIALISATION
// ─────────────────────────────────────────────────────────────

// Seed some default accounts if the user has none yet
// Hierarchy is inferred from code prefixes:
//   "1" is parent of "101", "101" is parent of "10101", etc.
function seedDefaults() {
  if (accounts.length === 0) {
    accounts = [
      // Assets (code prefix: 1 → 11 → 1101/1102…, 12 → …)
      { id: '1', name: 'Assets', type: 'Asset' },
      { id: '11', name: 'Current Assets', type: 'Asset' },
      { id: '1101', name: 'Cash', type: 'Asset' },
      { id: '1102', name: 'Accounts Receivable', type: 'Asset' },
      { id: '1103', name: 'Inventory', type: 'Asset' },
      { id: '12', name: 'Fixed Assets', type: 'Asset' },
      { id: '1201', name: 'Equipment', type: 'Asset' },
      { id: '1202', name: 'Furniture', type: 'Asset' },
      // Liabilities (code prefix: 2 → 21 → …)
      { id: '2', name: 'Liabilities', type: 'Liability' },
      { id: '21', name: 'Current Liabilities', type: 'Liability' },
      { id: '2101', name: 'Accounts Payable', type: 'Liability' },
      { id: '2102', name: 'Notes Payable', type: 'Liability' },
      // Equity (code prefix: 3 → 31…)
      { id: '3', name: 'Equity', type: 'Equity' },
      { id: '31', name: 'Owner\'s Capital', type: 'Equity' },
      { id: '32', name: 'Retained Earnings', type: 'Equity' },
      // Revenue (code prefix: 4 → 41…)
      { id: '4', name: 'Revenues', type: 'Revenue' },
      { id: '41', name: 'Sales Revenue', type: 'Revenue' },
      { id: '42', name: 'Service Revenue', type: 'Revenue' },
      // Expenses (code prefix: 5 → 51 → …)
      { id: '5', name: 'Expenses', type: 'Expense' },
      { id: '51', name: 'Operating Expenses', type: 'Expense' },
      { id: '5101', name: 'Cost of Goods Sold', type: 'Expense' },
      { id: '5102', name: 'Rent Expense', type: 'Expense' },
      { id: '5103', name: 'Salaries Expense', type: 'Expense' },
      { id: '5104', name: 'Utilities Expense', type: 'Expense' },
    ];
    saveData('acc_accounts', accounts);
    // Expand top-level nodes by default
    ['1', '2', '3', '4', '5', '__liab_equity__'].forEach(id => expandedNodes.add(id));
    saveExpanded();
  }
}

seedDefaults();
renderAccounts();
renderJournal();
