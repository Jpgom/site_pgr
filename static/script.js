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

function filterGroupSectors() {
  const select = document.getElementById('groupFilter');
  if (!select) return;
  const group = select.value;
  document.querySelectorAll('.sector-option, .sector-picker').forEach((el) => {
    const show = !group || el.dataset.group === group;
    el.style.display = show ? '' : 'none';
  });
}
const groupFilter = document.getElementById('groupFilter');
if (groupFilter) groupFilter.addEventListener('change', filterGroupSectors);

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
