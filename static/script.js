function isVisible(el) {
  return !el || el.offsetParent !== null || el.style.display !== 'none';
}

function toggleChecks(name, checked) {
  document.querySelectorAll(`input[name="${name}"]`).forEach((box) => {
    const item = box.closest('.compact-row,.list-check,.sector-option,.sector-picker,.risk-card');
    if (!item || item.style.display !== 'none') box.checked = checked;
  });
}
function toggleAll(checked) { toggleChecks('risk_ids', checked); }

function attachSearch(inputId, selector) {
  const input = document.getElementById(inputId);
  if (!input) return;
  input.addEventListener('input', () => {
    const term = input.value.trim().toLowerCase();
    document.querySelectorAll(selector).forEach((item) => {
      const text = item.dataset.search || item.textContent.toLowerCase();
      item.style.display = text.includes(term) ? '' : 'none';
    });
  });
}

attachSearch('riskSearch', '#riskList .compact-row');
attachSearch('companySearch', '#companyList .compact-row');
attachSearch('sectorGroupSearch', '#sectorGroupList .group-box');
attachSearch('pgrSectorSearch', '#pgrSectorList .sector-option');
attachSearch('riskPickerSearch', '.risk-option');
attachSearch('examPickerSearch', '.exam-option');

function selectedPgrSectorIds() {
  return new Set(Array.from(document.querySelectorAll('input[name="pgr_sector_ids"]:checked')).map((box) => box.value));
}

function syncSectorPickers() {
  const select = document.getElementById('groupFilter');
  const group = select ? select.value : '';
  const selected = selectedPgrSectorIds();
  document.querySelectorAll('.sector-picker').forEach((el) => {
    const matchesGroup = !group || el.dataset.group === group;
    const matchesSector = selected.has(el.dataset.sector);
    const show = matchesGroup && matchesSector;
    el.style.display = show ? '' : 'none';
    if (!show) el.removeAttribute('open');
  });
}

function filterGroupSectors() {
  const select = document.getElementById('groupFilter');
  if (!select) return;
  const group = select.value;
  document.querySelectorAll('.sector-option').forEach((el) => {
    const show = !group || el.dataset.group === group;
    el.style.display = show ? '' : 'none';
  });
  syncSectorPickers();
}
const groupFilter = document.getElementById('groupFilter');
if (groupFilter) groupFilter.addEventListener('change', filterGroupSectors);
document.querySelectorAll('input[name="pgr_sector_ids"]').forEach((box) => {
  box.addEventListener('change', syncSectorPickers);
});
syncSectorPickers();

function cargoRowTemplate() {
  const div = document.createElement('div');
  div.className = 'cargo-row';
  div.innerHTML = `
    <label><span>Cargo</span><input name="cargo_nome[]" required placeholder="Ex.: Auxiliar Administrativo"></label>
    <label><span>CBO</span><input name="cargo_cbo[]" required placeholder="Ex.: 4110-05"></label>
    <label><span>Nº funcionários</span><input name="cargo_nfunc[]" required placeholder="Ex.: 3"></label>
    <label class="span-3"><span>Descrição da atividade</span><textarea name="cargo_descricao[]" required rows="2" placeholder="Descreva as atividades principais do cargo."></textarea></label>
    <div class="cargo-actions"><button class="btn small ghost" type="button" onclick="removeCargoRow(this)">Remover</button></div>`;
  return div;
}
function addCargoRow() {
  const list = document.getElementById('cargoList');
  if (list) list.appendChild(cargoRowTemplate());
}
function removeCargoRow(button) {
  const list = document.getElementById('cargoList');
  const rows = list ? list.querySelectorAll('.cargo-row') : [];
  if (rows.length <= 1) { alert('Mantenha pelo menos um cargo no setor.'); return; }
  button.closest('.cargo-row').remove();
}

function togglePgrRisks(checked) {
  document.querySelectorAll('.pgr-risk-check').forEach((box) => { box.checked = checked; });
}
function togglePgrExams(checked) {
  document.querySelectorAll('.pgr-exam-check').forEach((box) => { box.checked = checked; });
}
function syncRevisionField() {
  const ajuste = document.getElementById('ajustePsicossocial');
  const field = document.getElementById('revisionField');
  const input = document.getElementById('dataDaRevisao');
  if (!ajuste || !field || !input) return;
  const active = ajuste.checked;
  field.style.display = active ? '' : 'none';
  input.required = active;
  if (!active) input.value = '';
}
const ajusteBox = document.getElementById('ajustePsicossocial');
if (ajusteBox) {
  ajusteBox.addEventListener('change', syncRevisionField);
  syncRevisionField();
}

const pgrForm = document.getElementById('pgrSelectionForm');
if (pgrForm) {
  pgrForm.addEventListener('submit', (event) => {
    const submitter = event.submitter;
    const action = submitter ? (submitter.getAttribute('formaction') || '') : '';
    const isComplete = action.includes('gerar-pgr-completo') || action.includes('gerar-pcmso-completo') || action.includes('gerar-ltcat-completo');
    if (!isComplete) return;

    const company = document.getElementById('companySelect');
    const dataCriacao = document.getElementById('dataCriacaoLaudo');
    const dataRevisao = document.getElementById('dataDaRevisao');
    const ajuste = document.getElementById('ajustePsicossocial');
    if (company && !company.value) {
      event.preventDefault();
      company.focus();
      alert('Selecione uma empresa antes de gerar o laudo. Suas marcações foram mantidas.');
      return;
    }
    if (dataCriacao && !dataCriacao.value.trim()) {
      event.preventDefault();
      dataCriacao.focus();
      alert('Preencha a Data de criação do laudo antes de gerar. Suas marcações foram mantidas.');
      return;
    }
    if (ajuste && ajuste.checked && dataRevisao && !dataRevisao.value.trim()) {
      event.preventDefault();
      dataRevisao.focus();
      alert('Preencha a Data da revisão psicossocial antes de gerar. Suas marcações foram mantidas.');
      return;
    }
  });
}

// Seleção compacta de riscos na geração dos laudos
function riskLabelOriginalIndex(label) {
  const raw = label.getAttribute('data-index') || '0';
  const parsed = parseInt(raw, 10);
  return Number.isNaN(parsed) ? 0 : parsed;
}

function updateRiskSelectionLists() {
  let totalSelected = 0;
  document.querySelectorAll('.risk-list-mode').forEach((list) => {
    const labels = Array.from(list.querySelectorAll('.risk-list-row'));
    let selectedCount = 0;

    labels.forEach((label) => {
      const input = label.querySelector('.pgr-risk-check');
      const isChecked = !!(input && input.checked);
      label.classList.toggle('is-selected', isChecked);
      if (isChecked) selectedCount += 1;
    });

    totalSelected += selectedCount;

    labels
      .sort((a, b) => {
        const aChecked = a.classList.contains('is-selected') ? 1 : 0;
        const bChecked = b.classList.contains('is-selected') ? 1 : 0;
        if (aChecked !== bChecked) return bChecked - aChecked;
        return riskLabelOriginalIndex(a) - riskLabelOriginalIndex(b);
      })
      .forEach((label) => list.appendChild(label));

    const picker = list.closest('.sector-picker');
    if (picker) {
      const countText = `${selectedCount} ${selectedCount === 1 ? 'risco selecionado' : 'riscos selecionados'}`;
      picker.querySelectorAll('[data-risk-count], [data-risk-count-badge]').forEach((el) => {
        el.textContent = countText;
      });
    }
  });

  const total = document.getElementById('totalRiskSelected');
  if (total) {
    total.textContent = `${totalSelected} ${totalSelected === 1 ? 'risco selecionado' : 'riscos selecionados'}`;
  }
}

document.querySelectorAll('.pgr-risk-check').forEach((box) => {
  box.addEventListener('change', updateRiskSelectionLists);
});
updateRiskSelectionLists();

// V20 - cadastro rápido de risco dentro da tela Gerar laudos, sem recarregar a página
function openModalElement(modal) {
  if (!modal) return;
  modal.classList.add('is-open');
  modal.setAttribute('aria-hidden', 'false');
  document.body.classList.add('modal-open');
  const firstInput = modal.querySelector('input[name="risco"]');
  if (firstInput) setTimeout(() => firstInput.focus(), 50);
}

function closeModalElement(modal) {
  if (!modal) return;
  modal.classList.remove('is-open');
  modal.setAttribute('aria-hidden', 'true');
  document.body.classList.remove('modal-open');
}

function showQuickRiskErrors(errors) {
  const box = document.getElementById('quickRiskErrors');
  if (!box) return;
  const list = Array.isArray(errors) ? errors : [String(errors || 'Erro ao salvar o risco.')];
  box.innerHTML = list.map((error) => `<div>${error}</div>`).join('');
  box.style.display = list.length ? '' : 'none';
}

function clearQuickRiskErrors() {
  const box = document.getElementById('quickRiskErrors');
  if (!box) return;
  box.innerHTML = '';
  box.style.display = 'none';
}

function createRiskOptionLabel(risk, sectorId, checked) {
  const label = document.createElement('label');
  label.className = `tiny-check risk-option risk-list-row${checked ? ' is-selected' : ''}`;
  label.setAttribute('data-index', '-1');
  label.dataset.search = `${risk.risco || ''} ${risk.tipo_risco || ''} ${risk.grau_nivel_risco || ''}`.toLowerCase();
  label.dataset.dynamicRiskId = risk.id;

  const input = document.createElement('input');
  input.className = 'pgr-risk-check';
  input.type = 'checkbox';
  input.name = `sector_risk_ids_${sectorId}`;
  input.value = risk.id;
  input.checked = !!checked;

  const title = document.createElement('span');
  title.textContent = risk.risco || '';

  const meta = document.createElement('em');
  meta.textContent = `${risk.tipo_risco || ''} · ${risk.grau_nivel_risco || ''}`;

  label.appendChild(input);
  label.appendChild(title);
  label.appendChild(meta);
  return label;
}

function addRiskToGenerationLists(risk, autoSelect) {
  const selectedSectors = selectedPgrSectorIds();
  const searchInput = document.getElementById('riskPickerSearch');
  const term = searchInput ? searchInput.value.trim().toLowerCase() : '';

  document.querySelectorAll('.risk-list-mode').forEach((list) => {
    const sectorId = list.dataset.riskList;
    if (!sectorId) return;
    if (list.querySelector(`input.pgr-risk-check[value="${CSS.escape(risk.id)}"]`)) return;
    const checked = !!autoSelect && selectedSectors.has(sectorId);
    const label = createRiskOptionLabel(risk, sectorId, checked);
    if (term && !label.dataset.search.includes(term)) label.style.display = 'none';
    list.prepend(label);
  });

  updateRiskSelectionLists();
}

const quickRiskModal = document.getElementById('quickRiskModal');
const openQuickRiskModalBtn = document.getElementById('openQuickRiskModal');
const closeQuickRiskModalBtn = document.getElementById('closeQuickRiskModal');
const cancelQuickRiskModalBtn = document.getElementById('cancelQuickRiskModal');
const quickRiskForm = document.getElementById('quickRiskForm');

if (openQuickRiskModalBtn && quickRiskModal) {
  openQuickRiskModalBtn.addEventListener('click', () => {
    clearQuickRiskErrors();
    openModalElement(quickRiskModal);
  });
}
if (closeQuickRiskModalBtn && quickRiskModal) closeQuickRiskModalBtn.addEventListener('click', () => closeModalElement(quickRiskModal));
if (cancelQuickRiskModalBtn && quickRiskModal) cancelQuickRiskModalBtn.addEventListener('click', () => closeModalElement(quickRiskModal));
if (quickRiskModal) {
  quickRiskModal.addEventListener('click', (event) => {
    if (event.target === quickRiskModal) closeModalElement(quickRiskModal);
  });
}
document.addEventListener('keydown', (event) => {
  if (event.key === 'Escape' && quickRiskModal && quickRiskModal.classList.contains('is-open')) {
    closeModalElement(quickRiskModal);
  }
});

document.addEventListener('change', (event) => {
  if (event.target && event.target.classList && event.target.classList.contains('pgr-risk-check')) {
    updateRiskSelectionLists();
  }
});

if (quickRiskForm) {
  quickRiskForm.addEventListener('submit', async (event) => {
    event.preventDefault();
    clearQuickRiskErrors();

    if (!quickRiskForm.reportValidity()) return;

    const button = document.getElementById('saveQuickRiskBtn');
    const originalText = button ? button.textContent : '';
    if (button) {
      button.disabled = true;
      button.textContent = 'Salvando...';
    }

    try {
      const response = await fetch('/api/risco/novo', {
        method: 'POST',
        body: new FormData(quickRiskForm),
        headers: { 'X-Requested-With': 'fetch' },
      });
      const payload = await response.json().catch(() => ({}));
      if (!response.ok || !payload.ok) {
        showQuickRiskErrors(payload.errors || ['Não foi possível salvar o risco. Confira os campos e tente novamente.']);
        return;
      }

      const autoSelect = !!document.getElementById('quickRiskAutoSelect')?.checked;
      addRiskToGenerationLists(payload.risk, autoSelect);
      quickRiskForm.reset();
      const fontes = quickRiskForm.querySelector('[name="fontes_circunstancias"]');
      const jornada = quickRiskForm.querySelector('[name="ltcat_periodicidade_jornada"]');
      const insal = quickRiskForm.querySelector('[name="ltcat_insalubridade"]');
      const grauInsal = quickRiskForm.querySelector('[name="ltcat_grau_insalubridade"]');
      const aposentadoria = quickRiskForm.querySelector('[name="ltcat_aposentadoria_especial"]');
      const autoSelectBox = document.getElementById('quickRiskAutoSelect');
      if (fontes) fontes.value = 'Durante o processo de trabalho.';
      if (jornada) jornada.value = 'Mensal (<= 4 horas < 10% jornada)';
      if (insal) insal.value = 'Não';
      if (grauInsal) grauInsal.value = 'Não aplicável';
      if (aposentadoria) aposentadoria.value = 'Não';
      if (autoSelectBox) autoSelectBox.checked = true;
      closeModalElement(quickRiskModal);
      alert('Risco cadastrado sem recarregar a página. Ele já está disponível na lista de riscos.');
    } catch (error) {
      showQuickRiskErrors(['Erro de conexão ao salvar o risco. Tente novamente.']);
    } finally {
      if (button) {
        button.disabled = false;
        button.textContent = originalText;
      }
    }
  });
}
