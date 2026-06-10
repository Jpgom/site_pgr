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
attachSearch('riskGroupSearch', '#riskGroupList .risk-group-row');
attachSearch('riskGroupFormSearch', '#riskGroupFormList .risk-group-risk-option');
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
    const isComplete = action.includes('gerar-pgr-completo') || action.includes('gerar-pcmso-completo') || action.includes('gerar-ltcat-completo') || action.includes('gerar-pgr-aet-psicossocial') || action.includes('gerar-aet-completa') || action.includes('gerar-pacote-empresa');
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
    if (action.includes('gerar-pgr-aet-psicossocial')) {
      const aetFile = pgrForm.querySelector('input[name="aet_docx"]');
      const psychFile = pgrForm.querySelector('input[name="psicossocial_pdf"]');
      if (aetFile && !aetFile.value) {
        event.preventDefault();
        aetFile.focus();
        alert('Envie a AET em Word para juntar com o PGR. Suas marcações foram mantidas.');
        return;
      }
      if (psychFile && !psychFile.value) {
        event.preventDefault();
        psychFile.focus();
        alert('Envie o Relatório Psicossocial em PDF. Suas marcações foram mantidas.');
        return;
      }
    }
    if (submitter && action.includes('gerar-pgr-aet-psicossocial') && !submitter.dataset.ajaxDownload) {
      submitter.disabled = true;
      submitter.textContent = 'Juntando arquivos... aguarde';
    }
  });
}

// Seleção compacta de riscos na geração dos laudos
function riskLabelOriginalIndex(label) {
  const raw = label.getAttribute('data-index') || '0';
  const parsed = parseInt(raw, 10);
  return Number.isNaN(parsed) ? 0 : parsed;
}

function riskIdsFromGroupBox(groupBox) {
  return (groupBox.dataset.riskIds || '')
    .split(',')
    .map((item) => item.trim())
    .filter(Boolean);
}

function riskInputInSector(sectorId, riskId) {
  const list = document.querySelector(`[data-risk-list="${sectorId}"]`);
  if (!list) return null;
  return Array.from(list.querySelectorAll('.pgr-risk-check')).find((box) => box.value === riskId) || null;
}

function applySectorRiskGroup(groupBox) {
  const sectorId = groupBox.dataset.sector;
  const checked = groupBox.checked;
  riskIdsFromGroupBox(groupBox).forEach((riskId) => {
    const input = riskInputInSector(sectorId, riskId);
    if (input) input.checked = checked;
  });
}

function applyCheckedRiskGroupsOnLoad() {
  document.querySelectorAll('.sector-risk-group-check:checked').forEach((groupBox) => {
    applySectorRiskGroup(groupBox);
  });
}

function syncSectorRiskGroupStates() {
  document.querySelectorAll('.sector-risk-group-check').forEach((groupBox) => {
    const sectorId = groupBox.dataset.sector;
    const ids = riskIdsFromGroupBox(groupBox);
    if (!sectorId || ids.length === 0) {
      groupBox.checked = false;
      groupBox.indeterminate = false;
      return;
    }
    const states = ids.map((riskId) => {
      const input = riskInputInSector(sectorId, riskId);
      return !!(input && input.checked);
    });
    const allSelected = states.length > 0 && states.every(Boolean);
    const someSelected = states.some(Boolean);
    groupBox.checked = allSelected;
    groupBox.indeterminate = someSelected && !allSelected;
  });
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

  syncSectorRiskGroupStates();
}

applyCheckedRiskGroupsOnLoad();

document.querySelectorAll('.pgr-risk-check').forEach((box) => {
  box.addEventListener('change', updateRiskSelectionLists);
});
document.querySelectorAll('.sector-risk-group-check').forEach((box) => {
  box.addEventListener('change', () => {
    applySectorRiskGroup(box);
    updateRiskSelectionLists();
  });
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


// Configurações salvas da tela Gerar laudos
function readReportProfiles() {
  const node = document.getElementById('reportProfilesData');
  if (!node) return [];
  try { return JSON.parse(node.textContent || '[]'); } catch (_) { return []; }
}
const reportProfiles = readReportProfiles();
const reportProfilesById = Object.fromEntries(reportProfiles.map((profile) => [profile.id, profile]));

function filterProfileOptions() {
  const companySelect = document.getElementById('companySelect');
  const profileSelect = document.getElementById('profileSelect');
  if (!companySelect || !profileSelect) return;
  const companyId = companySelect.value;
  let visibleCount = 0;
  Array.from(profileSelect.options).forEach((option) => {
    if (!option.value) return;
    const show = !!companyId && option.dataset.company === companyId;
    option.disabled = !show;
    option.hidden = !show;
    if (show) visibleCount += 1;
  });
  const current = profileSelect.selectedOptions[0];
  if (current && current.value && current.disabled) profileSelect.value = '';
  profileSelect.title = visibleCount ? `${visibleCount} configuração(ões) salvas para esta empresa` : 'Nenhuma configuração salva para esta empresa';
}

function setCheckboxesByName(name, values) {
  const set = new Set(values || []);
  document.querySelectorAll(`input[name="${name}"]`).forEach((box) => { box.checked = set.has(box.value); });
}

function clearGenerationSelections() {
  document.querySelectorAll('input[name="pgr_sector_ids"], .pgr-risk-check, .pgr-exam-check, .sector-risk-group-check').forEach((box) => {
    box.checked = false;
    box.indeterminate = false;
  });
}

function loadReportProfile(profileId) {
  const profile = reportProfilesById[profileId];
  if (!profile) return;
  const state = profile.state || {};
  const companySelect = document.getElementById('companySelect');
  if (companySelect) companySelect.value = profile.company_id || state.company_id || '';
  const dataCriacao = document.getElementById('dataCriacaoLaudo');
  if (dataCriacao) dataCriacao.value = state.data_criacao_laudo || profile.data_criacao_laudo || '';
  const ajuste = document.getElementById('ajustePsicossocial');
  if (ajuste) ajuste.checked = (state.ajuste_psicossocial || profile.ajuste_psicossocial) === '1';
  const dataRevisao = document.getElementById('dataDaRevisao');
  if (dataRevisao) dataRevisao.value = state.data_da_revisao || profile.data_da_revisao || '';
  const profileName = document.getElementById('profileName');
  if (profileName) profileName.value = profile.nome || '';

  clearGenerationSelections();
  setCheckboxesByName('pgr_sector_ids', state.selected_sector_ids || []);
  Object.entries(state.selected_risk_ids_by_sector || {}).forEach(([sectorId, riskIds]) => setCheckboxesByName(`sector_risk_ids_${sectorId}`, riskIds));
  Object.entries(state.selected_risk_group_ids_by_sector || {}).forEach(([sectorId, groupIds]) => setCheckboxesByName(`sector_risk_group_ids_${sectorId}`, groupIds));
  Object.entries(state.selected_exam_ids_by_sector || {}).forEach(([sectorId, examIds]) => setCheckboxesByName(`sector_exam_ids_${sectorId}`, examIds));

  filterProfileOptions();
  syncRevisionField();
  syncSectorPickers();
  applyCheckedRiskGroupsOnLoad();
  updateRiskSelectionLists();
  alert('Configuração carregada. Confira os dados e gere o laudo desejado.');
}

const companySelectForProfiles = document.getElementById('companySelect');
if (companySelectForProfiles) companySelectForProfiles.addEventListener('change', filterProfileOptions);
const loadProfileBtn = document.getElementById('loadProfileBtn');
if (loadProfileBtn) {
  loadProfileBtn.addEventListener('click', () => {
    const profileSelect = document.getElementById('profileSelect');
    if (!profileSelect || !profileSelect.value) {
      alert('Selecione uma configuração salva para carregar.');
      return;
    }
    loadReportProfile(profileSelect.value);
  });
}
filterProfileOptions();

// Preenche o nome da empresa na aba de juntar arquivos avulsos
const joinCompanySelect = document.getElementById('joinCompanySelect');
if (joinCompanySelect) {
  joinCompanySelect.addEventListener('change', () => {
    const input = document.getElementById('joinCompanyName');
    const option = joinCompanySelect.selectedOptions[0];
    if (input && option && option.dataset.name) input.value = option.dataset.name;
  });
}

function getDownloadFilename(disposition, fallback) {
  if (!disposition) return fallback;
  const utf8Match = disposition.match(/filename\*=UTF-8''([^;]+)/i);
  if (utf8Match && utf8Match[1]) return decodeURIComponent(utf8Match[1].replace(/"/g, ''));
  const match = disposition.match(/filename="?([^";]+)"?/i);
  return match && match[1] ? match[1] : fallback;
}

function setAjaxStatus(form, message, type) {
  const status = form.querySelector('[data-ajax-status]') || document.querySelector('[data-ajax-status]');
  if (!status) return;
  status.textContent = message || '';
  status.classList.remove('success', 'error', 'loading');
  if (type) status.classList.add(type);
}

function resetFormButtons(form) {
  form.querySelectorAll('button[type="submit"]').forEach((button) => {
    if (button.dataset.originalText) button.textContent = button.dataset.originalText;
    button.disabled = false;
  });
}

async function handleAjaxDownload(event, form, submitter, action) {
  event.preventDefault();
  const activeButton = submitter || form.querySelector('button[type="submit"]');
  const originalText = activeButton ? activeButton.textContent : '';
  if (activeButton) {
    activeButton.dataset.originalText = originalText;
    activeButton.disabled = true;
    activeButton.textContent = activeButton.dataset.processingText || 'Processando... aguarde';
  }
  setAjaxStatus(form, 'Processando arquivo. Aguarde o download iniciar...', 'loading');

  try {
    const response = await fetch(action || form.action, {
      method: (form.method || 'POST').toUpperCase(),
      body: new FormData(form),
      headers: { 'X-Requested-With': 'XMLHttpRequest' },
    });

    if (!response.ok) {
      let errorMessage = 'Não foi possível gerar o arquivo.';
      const contentType = response.headers.get('content-type') || '';
      if (contentType.includes('application/json')) {
        const data = await response.json();
        errorMessage = data.error || data.message || errorMessage;
      } else {
        const text = await response.text();
        if (text) errorMessage = text.replace(/<[^>]*>/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 600) || errorMessage;
      }
      throw new Error(errorMessage);
    }

    const blob = await response.blob();
    const url = window.URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = getDownloadFilename(response.headers.get('content-disposition'), 'documento_gerado.docx');
    document.body.appendChild(link);
    link.click();
    link.remove();
    window.URL.revokeObjectURL(url);

    form.querySelectorAll('input[type="file"]').forEach((input) => { input.value = ''; });
    setAjaxStatus(form, 'Arquivo gerado. Você já pode selecionar novos arquivos e juntar outro documento sem recarregar a página.', 'success');
  } catch (error) {
    setAjaxStatus(form, error.message || 'Erro ao processar arquivo.', 'error');
    alert(error.message || 'Erro ao processar arquivo.');
  } finally {
    resetFormButtons(form);
  }
}

// Download por AJAX: mantém a página aberta após juntar arquivos e permite gerar outros sem recarregar.
document.querySelectorAll('form').forEach((form) => {
  form.addEventListener('submit', (event) => {
    if (event.defaultPrevented) return;
    const submitter = event.submitter;
    const action = submitter?.getAttribute('formaction') || form.getAttribute('action') || window.location.href;
    const shouldAjaxDownload = form.dataset.ajaxDownload === 'true' || submitter?.dataset.ajaxDownload === 'true' || action.includes('gerar-pgr-aet-psicossocial') || action.includes('gerar-aet-completa') || action.includes('gerar-pacote-empresa');
    if (!shouldAjaxDownload) return;
    if (!form.checkValidity()) return;
    handleAjaxDownload(event, form, submitter, action);
  });
});

// V26 - geração em etapas tipo wizard com barra de progresso
(function initWizardSteps(){
  const form = document.getElementById('pgrSelectionForm');
  if (!form) return;
  const stepNodes = Array.from(form.querySelectorAll('[data-wizard-step]'));
  const navs = Array.from(form.querySelectorAll('[data-step-nav]'));
  if (!stepNodes.length || !navs.length) return;
  let activeStep = Number(sessionStorage.getItem('sstWizardActiveStep') || '1');
  if (!activeStep || activeStep < 1 || activeStep > 6) activeStep = 1;

  function showStep(step){
    activeStep = step;
    sessionStorage.setItem('sstWizardActiveStep', String(step));
    stepNodes.forEach((node) => {
      const nodeStep = Number(node.dataset.wizardStep || '1');
      node.classList.toggle('wizard-hidden', nodeStep !== step);
    });
    navs.forEach((nav) => {
      const navStep = Number(nav.dataset.stepNav || '1');
      nav.classList.toggle('is-active', navStep === step);
      nav.classList.toggle('is-complete', navStep < step);
    });
    window.scrollTo({top:0, behavior:'smooth'});
  }

  function addStepControls(){
    for (let step=1; step<=6; step++) {
      const panels = stepNodes.filter((node) => Number(node.dataset.wizardStep || '0') === step && node.classList.contains('step-panel'));
      const panel = panels[panels.length-1];
      const content = panel ? panel.querySelector('.step-content') : null;
      if (!content || content.querySelector('.wizard-nav-controls')) continue;
      const controls = document.createElement('div');
      controls.className = 'wizard-nav-controls';
      const back = document.createElement('button');
      back.type = 'button';
      back.className = 'btn ghost';
      back.textContent = step === 1 ? 'Início' : 'Voltar';
      back.disabled = step === 1;
      back.addEventListener('click', () => showStep(Math.max(1, step-1)));
      const next = document.createElement('button');
      next.type = 'button';
      next.className = 'btn primary';
      next.textContent = step === 6 ? 'Conferir e gerar' : 'Próximo';
      next.addEventListener('click', () => showStep(Math.min(6, step+1)));
      controls.appendChild(back);
      controls.appendChild(next);
      content.appendChild(controls);
    }
  }

  navs.forEach((nav) => nav.addEventListener('click', () => showStep(Number(nav.dataset.stepNav || '1'))));
  addStepControls();
  showStep(activeStep);
})();

// V26 - motor de regras para sugerir exames com base nos riscos selecionados
const suggestExamsBtn = document.getElementById('suggestExamsBtn');
if (suggestExamsBtn && pgrForm) {
  suggestExamsBtn.addEventListener('click', async () => {
    suggestExamsBtn.disabled = true;
    const originalText = suggestExamsBtn.textContent;
    suggestExamsBtn.textContent = 'Sugerindo...';
    try {
      const response = await fetch('/api/sugerir-exames', {
        method: 'POST',
        body: new FormData(pgrForm),
        headers: { 'X-Requested-With': 'XMLHttpRequest' },
      });
      const payload = await response.json();
      if (!response.ok || !payload.ok) throw new Error(payload.error || 'Não foi possível sugerir exames.');
      let total = 0;
      Object.entries(payload.suggestions || {}).forEach(([sectorId, examIds]) => {
        (examIds || []).forEach((examId) => {
          const input = pgrForm.querySelector(`input[name="sector_exam_ids_${CSS.escape(sectorId)}"][value="${CSS.escape(examId)}"]`);
          if (input && !input.checked) {
            input.checked = true;
            total += 1;
            const label = input.closest('.exam-option');
            if (label) {
              label.classList.add('suggested-exam-highlight');
              setTimeout(() => label.classList.remove('suggested-exam-highlight'), 2500);
            }
          }
        });
      });
      alert(total ? `${total} exame(s) sugerido(s) e marcado(s) pelas regras técnicas.` : 'Nenhum exame novo foi sugerido. Confira se os exames estão cadastrados e se há riscos selecionados.');
    } catch (error) {
      alert(error.message || 'Erro ao sugerir exames.');
    } finally {
      suggestExamsBtn.disabled = false;
      suggestExamsBtn.textContent = originalText;
    }
  });
}
