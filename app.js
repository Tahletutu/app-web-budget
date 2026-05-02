const storageKeys = {
  apiUrl: 'budget-pwa-api-url',
};

const state = {
  apiUrl: localStorage.getItem(storageKeys.apiUrl) || '',
  deferredInstallPrompt: null,
};

const elements = {
  apiUrl: document.getElementById('apiUrl'),
  saveApiBtn: document.getElementById('saveApiBtn'),
  reloadBtn: document.getElementById('reloadBtn'),
  statusText: document.getElementById('statusText'),
  installBtn: document.getElementById('installBtn'),
  expenseTab: document.getElementById('expenseTab'),
  incomeTab: document.getElementById('incomeTab'),
  type: document.getElementById('type'),
  form: document.getElementById('transactionForm'),
  submitBtn: document.getElementById('submitBtn'),
  amount: document.getElementById('amount'),
  category: document.getElementById('category'),
  label: document.getElementById('label'),
  date: document.getElementById('date'),
  notes: document.getElementById('notes'),
  totalBalance: document.getElementById('totalBalance'),
  monthIncome: document.getElementById('monthIncome'),
  monthExpense: document.getElementById('monthExpense'),
  transactionCount: document.getElementById('transactionCount'),
  transactionsList: document.getElementById('transactionsList'),
  categoriesList: document.getElementById('categoriesList'),
};

init();

function init() {
  elements.apiUrl.value = state.apiUrl;
  elements.date.value = todayISO();

  elements.saveApiBtn.addEventListener('click', saveApiUrl);
  elements.reloadBtn.addEventListener('click', loadDashboard);
  elements.form.addEventListener('submit', handleSubmit);
  elements.expenseTab.addEventListener('click', () => selectType('depense'));
  elements.incomeTab.addEventListener('click', () => selectType('revenu'));

  window.addEventListener('beforeinstallprompt', event => {
    event.preventDefault();
    state.deferredInstallPrompt = event;
    elements.installBtn.classList.remove('hidden');
  });

  elements.installBtn.addEventListener('click', async () => {
    if (!state.deferredInstallPrompt) return;
    state.deferredInstallPrompt.prompt();
    await state.deferredInstallPrompt.userChoice;
    state.deferredInstallPrompt = null;
    elements.installBtn.classList.add('hidden');
  });

  if ('serviceWorker' in navigator) {
    window.addEventListener('load', () => {
      navigator.serviceWorker.register('./sw.js').catch(console.error);
    });
  }

  if (state.apiUrl) {
    loadDashboard();
  }
}

function saveApiUrl() {
  const value = elements.apiUrl.value.trim();
  if (!value) {
    setStatus('Colle une URL Apps Script valide.', true);
    return;
  }
  state.apiUrl = value;
  localStorage.setItem(storageKeys.apiUrl, value);
  setStatus('URL enregistrée. Synchronisation en cours...');
  loadDashboard();
}

function selectType(type) {
  elements.type.value = type;
  elements.expenseTab.classList.toggle('is-active', type === 'depense');
  elements.incomeTab.classList.toggle('is-active', type === 'revenu');
  elements.submitBtn.textContent = type === 'depense' ? 'Ajouter la dépense' : 'Ajouter le revenu';
}

async function loadDashboard() {
  if (!state.apiUrl) {
    setStatus('Configure l’URL pour commencer.', true);
    return;
  }

  setStatus('Chargement des données...');

  try {
    const response = await fetch(`${state.apiUrl}?action=bootstrap`, {
      method: 'GET',
      cache: 'no-store',
    });
    const payload = await response.json();

    if (!payload.ok) {
      throw new Error(payload.error || 'Erreur inconnue côté serveur.');
    }

    renderData(payload.data);
    setStatus('Synchronisé avec Google Sheets.');
  } catch (error) {
    console.error(error);
    setStatus(`Connexion impossible : ${error.message}`, true);
  }
}

async function handleSubmit(event) {
  event.preventDefault();

  if (!state.apiUrl) {
    setStatus('Enregistre d’abord l’URL Apps Script.', true);
    return;
  }

  const payload = {
    action: 'addTransaction',
    transaction: {
      type: elements.type.value,
      amount: elements.amount.value,
      category: elements.category.value.trim(),
      label: elements.label.value.trim(),
      date: elements.date.value,
      notes: elements.notes.value.trim(),
    },
  };

  elements.submitBtn.disabled = true;
  setStatus('Envoi en cours...');

  try {
    const response = await fetch(state.apiUrl, {
      method: 'POST',
      headers: {
        'Content-Type': 'text/plain;charset=utf-8',
      },
      body: JSON.stringify(payload),
    });

    const result = await response.json();

    if (!result.ok) {
      throw new Error(result.error || 'Impossible d’ajouter la transaction.');
    }

    renderData(result.data);
    elements.form.reset();
    elements.date.value = todayISO();
    selectType('depense');
    setStatus('Opération ajoutée et Google Sheet mis à jour.');
  } catch (error) {
    console.error(error);
    setStatus(`Erreur : ${error.message}`, true);
  } finally {
    elements.submitBtn.disabled = false;
  }
}

function renderData(data) {
  if (data.summary) {
    elements.totalBalance.textContent = formatCurrency(data.summary.totalBalance);
    elements.monthIncome.textContent = formatCurrency(data.summary.currentMonthIncome);
    elements.monthExpense.textContent = formatCurrency(data.summary.currentMonthExpense);
    elements.transactionCount.textContent = String(data.summary.transactionCount || 0);
  }

  if (Array.isArray(data.categories)) {
    renderCategories(data.categories);
  }

  if (Array.isArray(data.recentTransactions)) {
    renderTransactions(data.recentTransactions);
  }
}

function renderCategories(categories) {
  elements.categoriesList.innerHTML = categories
    .map(category => `<option value="${escapeHtml(category)}"></option>`)
    .join('');
}

function renderTransactions(items) {
  if (!items.length) {
    elements.transactionsList.className = 'transactions-list empty-state';
    elements.transactionsList.textContent = 'Aucune donnée pour l’instant.';
    return;
  }

  elements.transactionsList.className = 'transactions-list';
  elements.transactionsList.innerHTML = items
    .map(item => {
      const isIncome = item.type === 'revenu';
      const title = item.label || item.category || (isIncome ? 'Revenu' : 'Dépense');
      const subtitle = [item.category, item.date, item.notes].filter(Boolean).join(' · ');
      return `
        <article class="tx-row">
          <div class="tx-meta">
            <span class="tx-badge ${isIncome ? 'revenu' : 'depense'}"></span>
            <div class="tx-info">
              <div class="tx-title">${escapeHtml(title)}</div>
              <div class="tx-subtitle">${escapeHtml(subtitle)}</div>
            </div>
          </div>
          <div class="tx-amount">
            <strong class="${isIncome ? 'positive' : 'negative'}">${isIncome ? '+' : '-'}${formatCurrency(item.amount)}</strong>
            <span>${isIncome ? 'Revenu' : 'Dépense'}</span>
          </div>
        </article>
      `;
    })
    .join('');
}

function setStatus(message, isError = false) {
  elements.statusText.textContent = message;
  elements.statusText.style.color = isError ? '#b91c1c' : '#6b7280';
}

function formatCurrency(value) {
  return new Intl.NumberFormat('fr-FR', {
    style: 'currency',
    currency: 'EUR',
  }).format(Number(value || 0));
}

function todayISO() {
  const date = new Date();
  const offsetMs = date.getTimezoneOffset() * 60000;
  return new Date(date.getTime() - offsetMs).toISOString().split('T')[0];
}

function escapeHtml(text) {
  return String(text || '')
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#039;');
}
