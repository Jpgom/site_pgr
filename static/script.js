function isVisible(card) {
  return !card || card.style.display !== 'none';
}

function toggleChecks(name, checked) {
  document.querySelectorAll(`input[name="${name}"]`).forEach((box) => {
    const card = box.closest('.risk-card');
    if (isVisible(card)) {
      box.checked = checked;
    }
  });
}

// Compatibilidade com versões anteriores.
function toggleAll(checked) {
  toggleChecks('risk_ids', checked);
}

function attachSearch(inputId, cardSelector) {
  const searchInput = document.getElementById(inputId);
  if (!searchInput) return;
  searchInput.addEventListener('input', () => {
    const term = searchInput.value.trim().toLowerCase();
    document.querySelectorAll(cardSelector).forEach((card) => {
      const text = card.dataset.search || '';
      card.style.display = text.includes(term) ? '' : 'none';
    });
  });
}

attachSearch('riskSearch', '#riskList .risk-card');
attachSearch('sectorSearch', '#sectorList .risk-card');

function cargoRowTemplate() {
  const div = document.createElement('div');
  div.className = 'cargo-row';
  div.innerHTML = `
    <label>
      <span>Cargo</span>
      <input name="cargo_nome[]" required placeholder="Ex.: Auxiliar Administrativo">
    </label>
    <label>
      <span>CBO</span>
      <input name="cargo_cbo[]" required placeholder="Ex.: 4110-05">
    </label>
    <label>
      <span>Nº funcionários</span>
      <input name="cargo_nfunc[]" required placeholder="Ex.: 3">
    </label>
    <label class="span-3">
      <span>Descrição da atividade</span>
      <textarea name="cargo_descricao[]" required rows="3" placeholder="Descreva as atividades principais do cargo."></textarea>
    </label>
    <div class="cargo-actions">
      <button class="btn small ghost" type="button" onclick="removeCargoRow(this)">Remover cargo</button>
    </div>
  `;
  return div;
}

function addCargoRow() {
  const list = document.getElementById('cargoList');
  if (list) list.appendChild(cargoRowTemplate());
}

function removeCargoRow(button) {
  const list = document.getElementById('cargoList');
  const rows = list ? list.querySelectorAll('.cargo-row') : [];
  if (rows.length <= 1) {
    alert('Mantenha pelo menos um cargo no setor.');
    return;
  }
  button.closest('.cargo-row').remove();
}

attachSearch('pgrSectorSearch', '#pgrSectorList .risk-card');

function togglePgrRisks(checked) {
  document.querySelectorAll('.pgr-risk-check').forEach((box) => {
    const card = box.closest('.risk-card');
    if (isVisible(card)) {
      box.checked = checked;
    }
  });
}

function togglePgrExams(checked) {
  document.querySelectorAll('.pgr-exam-check').forEach((box) => {
    const card = box.closest('.risk-card');
    if (isVisible(card)) {
      box.checked = checked;
    }
  });
}
