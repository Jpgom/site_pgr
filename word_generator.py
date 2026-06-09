from __future__ import annotations

from copy import deepcopy
from datetime import datetime
from pathlib import Path
from typing import Any, Iterable, Mapping

from docx import Document
from docx.table import Table
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import RGBColor


BASE_DIR = Path(__file__).resolve().parent
TEMPLATE_PLANO_PATH = BASE_DIR / "modelos" / "modelo_plano_acao.docx"
TEMPLATE_PGR_PATH = BASE_DIR / "modelos" / "modelo_risco_pgr.docx"
TEMPLATE_PCMSO_PATH = BASE_DIR / "modelos" / "modelo_pcmso.docx"
TEMPLATE_RELACAO_PATH = BASE_DIR / "modelos" / "modelo_relacao_funcao_atividade.docx"
TEMPLATE_DESCRITIVO_SETOR_PATH = BASE_DIR / "modelos" / "modelo_descritivo_setor.docx"
TEMPLATE_PGR_COMPLETO_PATH = BASE_DIR / "modelos" / "modelo_pgr_completo.docx"
TEMPLATE_PCMSO_COMPLETO_PATH = BASE_DIR / "modelos" / "modelo_pcmso_completo.docx"
TEMPLATE_RISCOS_PCMSO_PATH = BASE_DIR / "modelos" / "modelo_riscos_pcmso.docx"
TEMPLATE_LTCAT_COMPLETO_PATH = BASE_DIR / "modelos" / "modelo_ltcat_completo.docx"

TIPO_RISCO_COLORS = {
    "ACIDENTE(MECÂNICO)": "0066FF",
    "ERGONÔMICO": "FFFF00",
    "ERGONÔMICO PSICOSSOCIAL": "FFFF00",
    "FÍSICO": "009933",
    "BIOLÓGICO": "E36C0A",
    "QUÍMICO": "EE0000",
}

SEVERIDADE_COLORS = {
    "INSIGNIFICANTE": "009900",
    "PEQUENO": "009900",
    "MÉDIO": "FFFF00",
    "GRANDE": "336699",
    "CATASTRÓFICO": "FF0000",
}

POSSIBILIDADE_COLORS = {
    "IMPROVÁVEL": "009900",
    "POSSÍVEL": "336699",
    "PROVÁVEL": "FF0000",
}

NIVEL_RISCO_COLORS = {
    "TRIVIAL": "009900",
    "BAIXO": "FFFF00",
    "MODERADO": "336699",
    "ALTO": "990066",
    "MUITO ALTO": "FF0000",
}


# ---------------------------------------------------------------------------
# Utilitários gerais de OOXML
# ---------------------------------------------------------------------------

def _normalize_option(value: Any) -> str:
    return str(value or "").strip().upper()


def _make_text(value: Any) -> list:
    """Cria nós OOXML de texto, preservando quebras de linha."""
    value = "" if value is None else str(value)
    lines = value.splitlines() or [""]
    nodes = []
    for index, line in enumerate(lines):
        if index > 0:
            nodes.append(OxmlElement("w:br"))
        text_node = OxmlElement("w:t")
        if line != line.strip():
            text_node.set(qn("xml:space"), "preserve")
        text_node.text = line
        nodes.append(text_node)
    return nodes


def _set_font_color(run_properties, color_hex: str | None) -> None:
    if not color_hex:
        return
    existing = run_properties.find(qn("w:color"))
    if existing is None:
        existing = OxmlElement("w:color")
        run_properties.append(existing)
    existing.set(qn("w:val"), color_hex.replace("#", "").upper())


def _set_cell_text(cell_xml, value: Any, font_color: str | None = "000000") -> None:
    """Substitui o texto visível da célula preservando estilo básico."""
    old_paragraphs = cell_xml.findall(qn("w:p"))
    first_paragraph = old_paragraphs[0] if old_paragraphs else None
    paragraph_properties = None
    run_properties = None

    if first_paragraph is not None:
        p_pr = first_paragraph.find(qn("w:pPr"))
        if p_pr is not None:
            paragraph_properties = deepcopy(p_pr)
        first_run = first_paragraph.find(qn("w:r"))
        if first_run is not None:
            r_pr = first_run.find(qn("w:rPr"))
            if r_pr is not None:
                run_properties = deepcopy(r_pr)

    if run_properties is None:
        run_properties = OxmlElement("w:rPr")
    _set_font_color(run_properties, font_color)

    for child in list(cell_xml):
        if child.tag != qn("w:tcPr"):
            cell_xml.remove(child)

    paragraph = OxmlElement("w:p")
    if paragraph_properties is not None:
        paragraph.append(paragraph_properties)
    run = OxmlElement("w:r")
    run.append(run_properties)
    for node in _make_text(value):
        run.append(node)
    paragraph.append(run)
    cell_xml.append(paragraph)


def _set_cell_shading(cell_xml, fill_hex: str | None) -> None:
    if not fill_hex:
        return
    fill_hex = fill_hex.replace("#", "").upper()
    tc_pr = cell_xml.find(qn("w:tcPr"))
    if tc_pr is None:
        tc_pr = OxmlElement("w:tcPr")
        cell_xml.insert(0, tc_pr)
    shd = tc_pr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tc_pr.append(shd)
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill_hex)


def _value_cell(rows, row_index: int, cell_index: int):
    return rows[row_index].findall(qn("w:tc"))[cell_index]


def _row_cell(row_xml, cell_index: int):
    return row_xml.findall(qn("w:tc"))[cell_index]


def _colored_value(cell_xml, value: Any, color_map: Mapping[str, str]) -> None:
    option = _normalize_option(value)
    fill = color_map.get(option)
    _set_cell_shading(cell_xml, fill)
    # Todas as letras devem permanecer pretas, inclusive nas células coloridas.
    _set_cell_text(cell_xml, option, font_color="000000")


def _insert_before_section_or_end(body, element) -> None:
    children = list(body)
    for index, child in enumerate(children):
        if child.tag == qn("w:sectPr"):
            body.insert(index, element)
            return
    body.append(element)


def _blank_paragraph() -> OxmlElement:
    paragraph = OxmlElement("w:p")
    p_pr = OxmlElement("w:pPr")
    spacing = OxmlElement("w:spacing")
    spacing.set(qn("w:before"), "0")
    spacing.set(qn("w:after"), "0")
    p_pr.append(spacing)
    paragraph.append(p_pr)
    return paragraph


def _page_break_paragraph() -> OxmlElement:
    paragraph = OxmlElement("w:p")
    p_pr = OxmlElement("w:pPr")
    spacing = OxmlElement("w:spacing")
    spacing.set(qn("w:before"), "0")
    spacing.set(qn("w:after"), "0")
    p_pr.append(spacing)
    paragraph.append(p_pr)
    run = OxmlElement("w:r")
    br = OxmlElement("w:br")
    br.set(qn("w:type"), "page")
    run.append(br)
    paragraph.append(run)
    return paragraph


# ---------------------------------------------------------------------------
# PLANO DE AÇÃO
# ---------------------------------------------------------------------------

def _fill_plano_row(row_xml, risk: Mapping[str, Any], setor: str = "", data_atual: str = "", data_final: str = "") -> None:
    cells = row_xml.findall(qn("w:tc"))
    if len(cells) >= 7:
        # Novo modelo: GES recebe o setor.
        # Datas do plano de ação:
        # - regra geral: prazo de implantação = Data atual/início da vigência;
        #                prazo de reavaliação = Data final da vigência.
        # - risco ERGONÔMICO PSICOSSOCIAL: implantação em 30 DIAS e reavaliação em 180 DIAS.
        is_psychosocial = _normalize_option(risk.get("tipo_risco")) == "ERGONÔMICO PSICOSSOCIAL"
        prazo_implantacao = "30 DIAS" if is_psychosocial else (data_atual or "")
        prazo_reavaliacao = "180 DIAS" if is_psychosocial else (data_final or "")

        _set_cell_text(cells[0], setor)
        _set_cell_text(cells[1], risk.get("risco", ""))
        _set_cell_text(cells[2], risk.get("acoes", ""))
        # cells[3] mantém o responsável fixo do modelo: ADMINISTRAÇÃO.
        # Alguns modelos do PGR completo têm uma célula extra/mesclada antes dos prazos.
        if len(cells) >= 8:
            _set_cell_text(cells[5], prazo_implantacao)
            _set_cell_text(cells[6], prazo_reavaliacao)
            _set_cell_text(cells[7], risk.get("indicador", ""))
        else:
            _set_cell_text(cells[4], prazo_implantacao)
            _set_cell_text(cells[5], prazo_reavaliacao)
            _set_cell_text(cells[6], risk.get("indicador", ""))
        return

    replacements = {
        "{{SETOR}}": setor,
        "{{risco}}": risk.get("risco", ""),
        "{{AÇÕES PREVENTIVA / CORRETIVA}}": risk.get("acoes", ""),
        "{{INDICADOR DE EFETIVIDADE}}": risk.get("indicador", ""),
    }
    for node in list(row_xml.iter(qn("w:t"))):
        text = node.text or ""
        if "{{mês/ano}}" in text:
            node.text = text.replace("{{mês/ano}}", "")
        else:
            for placeholder, value in replacements.items():
                if placeholder in text:
                    node.text = text.replace(placeholder, str(value or ""))
                    break


def _plano_entries_from_groups_or_risks(items: list[Mapping[str, Any]]) -> list[tuple[str, Mapping[str, Any]]]:
    """Converte a seleção em linhas de Plano de Ação: (setor, risco)."""
    if _is_grouped_payload(items):
        entries: list[tuple[str, Mapping[str, Any]]] = []
        for group in _sanitize_sector_risk_groups(items):
            setor = group["sector"].get("setor", "")
            for risk in group.get("risks", []):
                entries.append((setor, risk))
        return entries
    return [("", risk) for risk in items]


def _find_template_row(table_xml, marker: str = "{{risco}}"):
    for row in table_xml.tr_lst:
        text = "".join(node.text or "" for node in row.iter(qn("w:t")))
        if marker in text:
            return row
    return None


def generate_action_plan_docx(groups_or_risks: Iterable[Mapping[str, Any]], output_path: str | Path, data_atual: str | None = None, data_final: str | None = None) -> Path:
    items = list(groups_or_risks)
    data_atual = (data_atual or "").strip()
    data_final = (data_final or "").strip()
    entries = _plano_entries_from_groups_or_risks(items)
    if not entries:
        raise ValueError("Selecione pelo menos um setor e um risco para gerar o Plano de Ação.")
    if not TEMPLATE_PLANO_PATH.exists():
        raise FileNotFoundError(f"Modelo Word não encontrado: {TEMPLATE_PLANO_PATH}")

    doc = Document(str(TEMPLATE_PLANO_PATH))
    if not doc.tables:
        raise ValueError("O modelo do Plano de Ação precisa ter uma tabela com a linha-modelo.")

    table = doc.tables[0]
    template_row_xml = _find_template_row(table._tbl)
    if template_row_xml is None:
        raise ValueError("O modelo do Plano de Ação precisa ter uma linha-modelo com {{risco}}.")
    template_row_copy = deepcopy(template_row_xml)
    table._tbl.remove(template_row_xml)

    # Mantém a primeira linha (PLANO DE AÇÃO) e a segunda linha (cabeçalho) apenas uma vez.
    # Depois, insere todos os setores/riscos um abaixo do outro, sem espaçamento.
    for setor, risk in entries:
        new_row = deepcopy(template_row_copy)
        _fill_plano_row(new_row, risk, setor=setor, data_atual=data_atual, data_final=data_final)
        table._tbl.append(new_row)

    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(output_path))
    return output_path


# Compatibilidade com a primeira versão do site.
generate_docx_from_risks = generate_action_plan_docx


# ---------------------------------------------------------------------------
# RISCO PGR POR SETOR
# ---------------------------------------------------------------------------

def _sanitize_sector(sector: Mapping[str, Any]) -> dict[str, Any]:
    cargos = sector.get("cargos") or []
    clean_cargos = []
    for cargo in cargos:
        clean_cargos.append(
            {
                "cargo": str(cargo.get("cargo", "")).strip(),
                "cbo": str(cargo.get("cbo", "")).strip(),
                "n_func": str(cargo.get("n_func", "")).strip(),
                "descricao": str(cargo.get("descricao", "")).strip(),
            }
        )
    return {"setor": str(sector.get("setor", "")).strip(), "cargos": clean_cargos}


def _sector_cargos_text(sector: Mapping[str, Any]) -> str:
    cargos = sector.get("cargos") or []
    names = [str(cargo.get("cargo", "")).strip() for cargo in cargos if str(cargo.get("cargo", "")).strip()]
    return ", ".join(names)


def _is_grouped_payload(items: list[Mapping[str, Any]]) -> bool:
    return bool(items) and isinstance(items[0], Mapping) and "risks" in items[0]


def _sanitize_sector_risk_groups(groups: Iterable[Mapping[str, Any]]) -> list[dict[str, Any]]:
    clean_groups: list[dict[str, Any]] = []
    for group in groups:
        sector = _sanitize_sector(group.get("sector") or {})
        risks = [risk for risk in group.get("risks", []) if risk]
        if sector.get("setor") and risks:
            clean_groups.append({"sector": sector, "risks": risks})
    return clean_groups


def _fill_pgr_risk_rows(risk_rows: list, risk: Mapping[str, Any]) -> None:
    if len(risk_rows) < 11:
        raise ValueError("O bloco de risco do PGR precisa manter 11 linhas.")

    tipo = risk.get("tipo_risco", "")
    descricao = risk.get("descricao_agente") or risk.get("risco", "")
    possiveis_lesoes = risk.get("possiveis_lesoes", "")
    fontes = risk.get("fontes_circunstancias") or "Durante o processo de trabalho."
    epis = risk.get("epis", "")
    epcs = risk.get("epcs", "")
    severidade = risk.get("grau_severidade", "")
    possibilidade = risk.get("grau_possibilidade", "")
    nivel = risk.get("grau_nivel_risco", "")

    _colored_value(_row_cell(risk_rows[0], 1), tipo, TIPO_RISCO_COLORS)
    _set_cell_text(_row_cell(risk_rows[1], 1), descricao)
    _set_cell_text(_row_cell(risk_rows[2], 1), possiveis_lesoes)
    _set_cell_text(_row_cell(risk_rows[3], 1), fontes)
    _set_cell_text(_row_cell(risk_rows[5], 1), epis)
    _set_cell_text(_row_cell(risk_rows[6], 1), epcs)

    _colored_value(_row_cell(risk_rows[9], 1), severidade, SEVERIDADE_COLORS)
    _colored_value(_row_cell(risk_rows[9], 3), possibilidade, POSSIBILIDADE_COLORS)
    _colored_value(_row_cell(risk_rows[9], 5), nivel, NIVEL_RISCO_COLORS)


def _fill_pgr_sector_table(table_xml, sector: Mapping[str, Any], risks: list[Mapping[str, Any]]) -> None:
    rows = table_xml.findall(qn("w:tr"))
    if len(rows) < 18:
        raise ValueError("O novo modelo Risco PGR precisa manter a tabela original com 18 linhas.")
    if not risks:
        raise ValueError("Cada setor selecionado precisa ter pelo menos um risco.")

    # 4 primeiras linhas do setor.
    _set_cell_text(_row_cell(rows[0], 0), sector.get("setor", ""))
    _set_cell_text(_row_cell(rows[2], 0), _sector_cargos_text(sector))

    risk_template_rows = [deepcopy(row) for row in rows[4:15]]
    _fill_pgr_risk_rows(rows[4:15], risks[0])

    # A partir do segundo risco, entram apenas as linhas do bloco do risco, sem repetir cabeçalho do setor.
    footer_first_row = rows[15]
    insertion_index = list(table_xml).index(footer_first_row)
    for risk in risks[1:]:
        block_rows = [deepcopy(row) for row in risk_template_rows]
        _fill_pgr_risk_rows(block_rows, risk)
        for block_row in block_rows:
            table_xml.insert(insertion_index, block_row)
            insertion_index += 1

    has_psychosocial = any(_normalize_option(risk.get("tipo_risco")) == "ERGONÔMICO PSICOSSOCIAL" for risk in risks)
    if has_psychosocial:
        # A frase "NENHUM FATOR..." só aparece quando NÃO houver risco psicossocial no setor.
        table_xml.remove(footer_first_row)


def generate_pgr_docx(groups_or_risks: Iterable[Mapping[str, Any]], output_path: str | Path) -> Path:
    items = list(groups_or_risks)
    if not items:
        raise ValueError("Selecione pelo menos um setor e um risco para gerar o PGR.")
    if not TEMPLATE_PGR_PATH.exists():
        raise FileNotFoundError(f"Modelo Risco PGR não encontrado: {TEMPLATE_PGR_PATH}")

    # Compatibilidade: se a chamada antiga passar apenas riscos, cria um setor genérico.
    if _is_grouped_payload(items):
        groups = _sanitize_sector_risk_groups(items)
    else:
        groups = [{"sector": {"setor": "", "cargos": []}, "risks": items}]

    if not groups:
        raise ValueError("Selecione pelo menos um setor e um risco para gerar o PGR.")

    doc = Document(str(TEMPLATE_PGR_PATH))
    if not doc.tables:
        raise ValueError("O modelo Risco PGR precisa ter uma tabela-modelo.")

    body = doc._body._element
    original_table_xml = doc.tables[0]._tbl
    template_table_copy = deepcopy(original_table_xml)
    body.remove(original_table_xml)

    for index, group in enumerate(groups):
        if index > 0:
            _insert_before_section_or_end(body, _page_break_paragraph())
        new_table = deepcopy(template_table_copy)
        _fill_pgr_sector_table(new_table, group["sector"], group["risks"])
        _insert_before_section_or_end(body, new_table)

    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(output_path))
    return output_path


# ---------------------------------------------------------------------------
# PCMSO
# ---------------------------------------------------------------------------

def _fill_pcmso_table(table_xml, risk: Mapping[str, Any]) -> None:
    rows = table_xml.findall(qn("w:tr"))
    if len(rows) < 8:
        raise ValueError("O modelo PCMSO precisa manter a tabela original com 8 linhas.")

    tipo = risk.get("tipo_risco", "")
    descricao = risk.get("descricao_agente") or risk.get("risco", "")
    possiveis_lesoes = risk.get("possiveis_lesoes", "")
    fontes = risk.get("fontes_circunstancias") or "Durante o processo de trabalho."
    epis = risk.get("epis", "")
    epcs = risk.get("epcs", "")

    _colored_value(_value_cell(rows, 0, 1), tipo, TIPO_RISCO_COLORS)
    _set_cell_text(_value_cell(rows, 1, 1), descricao)
    _set_cell_text(_value_cell(rows, 2, 1), possiveis_lesoes)
    _set_cell_text(_value_cell(rows, 3, 1), fontes)
    _set_cell_text(_value_cell(rows, 5, 1), epis)
    _set_cell_text(_value_cell(rows, 6, 1), epcs)


def _flatten_risks_from_groups_or_risks(groups_or_risks: Iterable[Mapping[str, Any]]) -> list[Mapping[str, Any]]:
    items = list(groups_or_risks)
    if _is_grouped_payload(items):
        flattened: list[Mapping[str, Any]] = []
        for group in items:
            flattened.extend(group.get("risks", []))
        return flattened
    return items


def generate_pcmso_docx(groups_or_risks: Iterable[Mapping[str, Any]], output_path: str | Path) -> Path:
    risks = _flatten_risks_from_groups_or_risks(groups_or_risks)
    if not risks:
        raise ValueError("Selecione pelo menos um risco para gerar o PCMSO.")
    if not TEMPLATE_PCMSO_PATH.exists():
        raise FileNotFoundError(f"Modelo PCMSO não encontrado: {TEMPLATE_PCMSO_PATH}")

    doc = Document(str(TEMPLATE_PCMSO_PATH))
    if not doc.tables:
        raise ValueError("O modelo PCMSO precisa ter uma tabela-modelo.")

    body = doc._body._element
    original_table_xml = doc.tables[0]._tbl
    template_table_copy = deepcopy(original_table_xml)
    body.remove(original_table_xml)

    for risk in risks:
        new_table = deepcopy(template_table_copy)
        _fill_pcmso_table(new_table, risk)
        _insert_before_section_or_end(body, new_table)

    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(output_path))
    return output_path


# ---------------------------------------------------------------------------
# RELAÇÃO FUNÇÃO X ATIVIDADE
# ---------------------------------------------------------------------------

def _fill_relacao_cargo_row(row_xml, cargo: Mapping[str, Any]) -> None:
    cells = row_xml.findall(qn("w:tc"))
    if len(cells) < 3:
        raise ValueError("A linha de cargo do modelo Relação Função x Atividade precisa ter 3 células.")
    cargo_text = str(cargo.get("cargo", "")).strip()
    cbo_text = str(cargo.get("cbo", "")).strip()
    func_text = str(cargo.get("n_func", "")).strip()
    desc_text = str(cargo.get("descricao", "")).strip()

    _set_cell_text(cells[0], f"{cargo_text} – CBO: {cbo_text}")
    _set_cell_text(cells[1], func_text)
    _set_cell_text(cells[2], desc_text)


def _fill_relacao_table(table_xml, sector: Mapping[str, Any], data_atual: str, data_final: str) -> None:
    sector = _sanitize_sector(sector)
    rows = table_xml.findall(qn("w:tr"))
    if len(rows) < 4:
        raise ValueError("O modelo Relação Função x Atividade precisa manter as 4 linhas originais.")

    _set_cell_text(_value_cell(rows, 0, 1), sector.get("setor", ""))
    _set_cell_text(_value_cell(rows, 1, 1), f"{data_atual} – {data_final}".strip(" –"))

    template_cargo_row = rows[3]
    template_cargo_copy = deepcopy(template_cargo_row)
    table_xml.remove(template_cargo_row)

    cargos = sector.get("cargos") or []
    if not cargos:
        cargos = [{"cargo": "", "cbo": "", "n_func": "", "descricao": ""}]
    for cargo in cargos:
        new_row = deepcopy(template_cargo_copy)
        _fill_relacao_cargo_row(new_row, cargo)
        table_xml.append(new_row)


def generate_relacao_funcao_atividade_docx(
    sectors: Iterable[Mapping[str, Any]],
    output_path: str | Path,
    data_atual: str | None = None,
    data_final: str | None = None,
    company_extra: Mapping[str, Any] | None = None,
) -> Path:
    sectors = [_sanitize_sector(sector) for sector in sectors]
    if not sectors:
        raise ValueError("Selecione pelo menos um setor para gerar o Word.")
    if not TEMPLATE_RELACAO_PATH.exists():
        raise FileNotFoundError(f"Modelo Relação Função x Atividade não encontrado: {TEMPLATE_RELACAO_PATH}")

    data_atual = (data_atual or datetime.now().strftime("%d/%m/%Y")).strip()
    data_final = (data_final or "").strip()

    doc = Document(str(TEMPLATE_RELACAO_PATH))
    if not doc.tables:
        raise ValueError("O modelo Relação Função x Atividade precisa ter uma tabela-modelo.")

    body = doc._body._element
    original_table_xml = doc.tables[0]._tbl
    template_table_copy = deepcopy(original_table_xml)
    body.remove(original_table_xml)

    for index, sector in enumerate(sectors):
        new_table = deepcopy(template_table_copy)
        _fill_relacao_table(new_table, sector, data_atual, data_final)
        _insert_before_section_or_end(body, new_table)
        if index < len(sectors) - 1:
            _insert_before_section_or_end(body, _blank_paragraph())

    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(output_path))
    return output_path


# ---------------------------------------------------------------------------
# DESCRITIVO SETOR
# ---------------------------------------------------------------------------

def _fill_descritivo_setor_table(table_xml, sector: Mapping[str, Any]) -> None:
    rows = table_xml.findall(qn("w:tr"))
    if len(rows) < 1:
        raise ValueError("O modelo Descritivo Setor precisa manter a tabela original.")
    _set_cell_text(_value_cell(rows, 0, 1), str(sector.get("setor", "")).strip())


def generate_descritivo_setor_docx(sectors: Iterable[Mapping[str, Any]], output_path: str | Path) -> Path:
    sectors = [_sanitize_sector(sector) for sector in sectors]
    if not sectors:
        raise ValueError("Selecione pelo menos um setor para gerar o Word.")
    if not TEMPLATE_DESCRITIVO_SETOR_PATH.exists():
        raise FileNotFoundError(f"Modelo Descritivo Setor não encontrado: {TEMPLATE_DESCRITIVO_SETOR_PATH}")

    doc = Document(str(TEMPLATE_DESCRITIVO_SETOR_PATH))
    if not doc.tables:
        raise ValueError("O modelo Descritivo Setor precisa ter uma tabela-modelo.")

    body = doc._body._element
    original_table_xml = doc.tables[0]._tbl
    template_table_copy = deepcopy(original_table_xml)
    body.remove(original_table_xml)

    for sector in sectors:
        new_table = deepcopy(template_table_copy)
        _fill_descritivo_setor_table(new_table, sector)
        _insert_before_section_or_end(body, new_table)

    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(output_path))
    return output_path


# ---------------------------------------------------------------------------
# PGR COMPLETO
# ---------------------------------------------------------------------------


def _xml_text(element) -> str:
    return "".join(node.text or "" for node in element.iter(qn("w:t")))


def _table_contains_all(table_xml, markers: Iterable[str]) -> bool:
    text = _xml_text(table_xml)
    return all(marker in text for marker in markers)


def _find_table_xml(doc: Document, markers: Iterable[str]):
    markers = list(markers)
    for table in doc.tables:
        if _table_contains_all(table._tbl, markers):
            return table._tbl
    raise ValueError(f"Não encontrei a tabela-modelo com os marcadores: {', '.join(markers)}")


def _replace_xml_element_with(parent_element, original_element, new_elements: Iterable) -> None:
    index = list(parent_element).index(original_element)
    parent_element.remove(original_element)
    for offset, element in enumerate(new_elements):
        parent_element.insert(index + offset, element)


def _replace_text_in_paragraph(paragraph, replacements: Mapping[str, str]) -> None:
    full_text = paragraph.text or ""
    if not any(marker in full_text for marker in replacements):
        return

    # Primeiro tenta substituir dentro dos próprios runs para preservar ao máximo a formatação.
    for run in paragraph.runs:
        text = run.text
        for marker, value in replacements.items():
            text = text.replace(marker, value)
        run.text = text

    # Se o Word tiver quebrado o marcador em vários runs, reconstrói o parágrafo usando o estilo do primeiro run.
    full_text = paragraph.text or ""
    if any(marker in full_text for marker in replacements):
        for marker, value in replacements.items():
            full_text = full_text.replace(marker, value)
        first_rpr = None
        if paragraph.runs and paragraph.runs[0]._r.rPr is not None:
            first_rpr = deepcopy(paragraph.runs[0]._r.rPr)
        paragraph.clear()
        run = paragraph.add_run(full_text)
        if first_rpr is not None:
            run._r.insert(0, first_rpr)


def _replace_doc_placeholders(doc: Document, replacements: Mapping[str, Any]) -> None:
    clean = {key: "" if value is None else str(value) for key, value in replacements.items()}

    def replace_in_table(table):
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    _replace_text_in_paragraph(paragraph, clean)

    for paragraph in doc.paragraphs:
        _replace_text_in_paragraph(paragraph, clean)
    for table in doc.tables:
        replace_in_table(table)

    for section in doc.sections:
        parts = [section.header, section.footer]
        # Em alguns modelos o Word cria variações de primeira página/par/ímpar.
        for attr in ("first_page_header", "first_page_footer", "even_page_header", "even_page_footer"):
            part = getattr(section, attr, None)
            if part is not None:
                parts.append(part)
        for part in parts:
            for paragraph in part.paragraphs:
                _replace_text_in_paragraph(paragraph, clean)
            for table in part.tables:
                replace_in_table(table)



def _force_page_breaks_before_headings(doc: Document, headings: Iterable[str]) -> None:
    targets = tuple(headings)
    for paragraph in doc.paragraphs:
        text = (paragraph.text or "").strip()
        if any(text.startswith(target) for target in targets):
            paragraph.paragraph_format.page_break_before = True


def _compact_first_page_date_block(doc: Document) -> None:
    """Mantém o bloco Data/Revisão compacto para caber na primeira página da capa."""
    date_index = None
    for index, paragraph in enumerate(doc.paragraphs):
        text = (paragraph.text or "").strip()
        if text.startswith("Data do documento:"):
            date_index = index
            break

    if date_index is None:
        return

    # Remove a quebra forçada antes da data, caso tenha sido aplicada em versões anteriores.
    for paragraph in doc.paragraphs[date_index: min(len(doc.paragraphs), date_index + 5)]:
        paragraph.paragraph_format.page_break_before = False
        paragraph.paragraph_format.space_before = 0
        paragraph.paragraph_format.space_after = 0

    # Compacta parágrafos vazios imediatamente anteriores ao bloco da data.
    for paragraph in reversed(doc.paragraphs[:date_index]):
        if (paragraph.text or "").strip():
            break
        paragraph.text = ""
        paragraph.paragraph_format.space_before = 0
        paragraph.paragraph_format.space_after = 0
        paragraph.paragraph_format.line_spacing = 1


def _clone_paragraph_style(new_p, source_p) -> None:
    p_pr = source_p.find(qn("w:pPr"))
    if p_pr is not None:
        new_p.insert(0, deepcopy(p_pr))


def _make_revision_paragraph(text: str, source_paragraph) -> OxmlElement:
    p = OxmlElement("w:p")
    _clone_paragraph_style(p, source_paragraph._p)
    r = OxmlElement("w:r")
    r_pr = None
    first_run = source_paragraph._p.find(qn("w:r"))
    if first_run is not None:
        maybe_rpr = first_run.find(qn("w:rPr"))
        if maybe_rpr is not None:
            r_pr = deepcopy(maybe_rpr)
    if r_pr is not None:
        r.append(r_pr)
    for node in _make_text(text):
        r.append(node)
    p.append(r)
    return p


def _insert_psychosocial_revision_line(doc: Document, company: Mapping[str, str]) -> None:
    """Insere a linha de revisão de inclusão de risco psicossocial entre Data do documento e Revisão periódica."""
    if company.get("ajuste_psicossocial") != "1":
        return
    data_revisao = (company.get("data_da_revisao") or "").strip()
    if not data_revisao:
        return
    text = f"Revisão periódica de inclusão de Risco Psicossocial: {data_revisao}"
    for paragraph in doc.paragraphs:
        p_text = (paragraph.text or "").strip()
        if p_text.startswith("Data do documento:"):
            # Evita duplicar a linha caso o documento seja processado novamente.
            next_el = paragraph._p.getnext()
            while next_el is not None and next_el.tag == qn("w:p"):
                next_text = "".join(t.text or "" for t in next_el.iter(qn("w:t"))).strip()
                if next_text:
                    if next_text.startswith("Revisão periódica de inclusão de Risco Psicossocial"):
                        return
                    break
                next_el = next_el.getnext()
            new_p = _make_revision_paragraph(text, paragraph)
            paragraph._p.addnext(new_p)
            return


def _compact_known_static_spacers(doc: Document) -> None:
    """Suaviza espaçamentos exagerados de páginas estáticas do modelo, sem alterar tabelas."""
    compact_starts = (
        "FICHA DE ENTREGA GRATUITA DE EPI",
        "TERMO DE RESPONSABILIDADE",
    )
    page_break_and_compact = (
        "IDENTIFICAÇÃO DA EMPRESA",
    )
    for paragraph in doc.paragraphs:
        text = (paragraph.text or "").strip()
        if any(text.startswith(item) for item in compact_starts):
            paragraph.paragraph_format.space_before = 0
            paragraph.paragraph_format.space_after = 0
        if any(text.startswith(item) for item in page_break_and_compact):
            paragraph.paragraph_format.page_break_before = True
            paragraph.paragraph_format.space_before = 0
            paragraph.paragraph_format.space_after = 0


def _build_relacao_elements(template_table_xml, sectors: list[Mapping[str, Any]], data_atual: str, data_final: str) -> list:
    elements: list = []
    for index, sector in enumerate(sectors):
        new_table = deepcopy(template_table_xml)
        _fill_relacao_table(new_table, sector, data_atual, data_final)
        elements.append(new_table)
        if index < len(sectors) - 1:
            elements.append(_blank_paragraph())
    return elements


def _build_descritivo_elements(template_table_xml, sectors: list[Mapping[str, Any]]) -> list:
    elements: list = []
    for sector in sectors:
        new_table = deepcopy(template_table_xml)
        _fill_descritivo_setor_table(new_table, sector)
        elements.append(new_table)
    return elements


def _build_risco_pgr_elements(template_table_xml, groups: list[Mapping[str, Any]], break_before_first: bool = False) -> list:
    elements: list = []
    for index, group in enumerate(groups):
        if index > 0 or (index == 0 and break_before_first):
            elements.append(_page_break_paragraph())
        new_table = deepcopy(template_table_xml)
        _fill_pgr_sector_table(new_table, group["sector"], group["risks"])
        elements.append(new_table)
    return elements


def _fill_action_plan_table_xml(table_xml, groups: list[Mapping[str, Any]], data_atual: str = "", data_final: str = "") -> None:
    entries = _plano_entries_from_groups_or_risks(groups)
    if not entries:
        raise ValueError("Selecione pelo menos um setor e um risco para preencher o Plano de Ação no PGR.")

    template_row_xml = _find_template_row(table_xml)
    if template_row_xml is None:
        raise ValueError("A tabela de Plano de Ação do PGR precisa ter uma linha-modelo com {{risco}}.")
    template_row_copy = deepcopy(template_row_xml)
    table_xml.remove(template_row_xml)

    for setor, risk in entries:
        new_row = deepcopy(template_row_copy)
        _fill_plano_row(new_row, risk, setor=setor, data_atual=data_atual, data_final=data_final)
        table_xml.append(new_row)


def _build_action_plan_elements(template_table_xml, groups: list[Mapping[str, Any]], data_atual: str = "", data_final: str = "") -> list:
    new_table = deepcopy(template_table_xml)
    _fill_action_plan_table_xml(new_table, groups, data_atual=data_atual, data_final=data_final)
    return [new_table]


def _unique_sectors_from_groups(groups: list[Mapping[str, Any]]) -> list[dict[str, Any]]:
    sectors: list[dict[str, Any]] = []
    seen: set[str] = set()
    for group in groups:
        sector = _sanitize_sector(group.get("sector") or {})
        key = sector.get("setor", "").strip().lower()
        if key and key not in seen:
            sectors.append(sector)
            seen.add(key)
    return sectors




def _company_dict_for_docs(empresa: str, cnpj: str, data_atual: str, data_final: str, extra: Mapping[str, Any] | None = None) -> dict[str, str]:
    extra = extra or {}
    def pick(*keys: str) -> str:
        for key in keys:
            value = extra.get(key)
            if value not in (None, ""):
                return str(value).strip()
        return ""
    data_criacao = pick("data_criacao_laudo", "datacriacaolaudo", "data_avaliacao")
    return {
        "empresa": str(empresa or pick("empresa", "nome")).strip(),
        "cnpj": str(cnpj or pick("cnpj")).strip(),
        "data_atual": str(data_atual or pick("data_atual")).strip(),
        "data_final": str(data_final or pick("data_final")).strip(),
        "endereco": pick("endereco"),
        "bairro_cidade": pick("bairro_cidade", "bairro/cidade"),
        "cep": pick("cep"),
        "cnae": pick("cnae", "cnae1"),
        "descricao_atividade": pick("descricao_atividade", "descricao1"),
        "grau_risco": pick("grau_risco", "grau1"),
        "cnae_secundario": pick("cnae_secundario", "cnae2"),
        "descricao_atividade_secundaria": pick("descricao_atividade_secundaria", "descricao2"),
        "grau_risco_secundario": pick("grau_risco_secundario", "grau2"),
        "funcionarios": pick("funcionarios"),
        "email": pick("email"),
        "fone": pick("fone"),
        "data_avaliacao": data_criacao or str(data_atual or "").strip(),
        "data_criacao_laudo": data_criacao,
        "ajuste_psicossocial": pick("ajuste_psicossocial"),
        "data_da_revisao": pick("data_da_revisao", "data_revisao_ajuste"),
    }


def _total_funcionarios_from_sectors(sectors: list[Mapping[str, Any]]) -> str:
    total = 0
    for sector in sectors:
        for cargo in sector.get("cargos", []) or []:
            try:
                total += int(str(cargo.get("n_func", "")).strip() or "0")
            except ValueError:
                pass
    return str(total) if total else ""


def _company_replacements(company: Mapping[str, str]) -> dict[str, str]:
    funcionarios = company.get("funcionarios", "")
    return {
        "{{EMPRESA}}": company.get("empresa", ""),
        "{{empresa}}": company.get("empresa", ""),
        "{{CNPJ}}": company.get("cnpj", ""),
        "{{DATAATUAL}}": company.get("data_atual", ""),
        "{{DATAFINAL}}": company.get("data_final", ""),
        "{{DATA DO LAUDO}}": company.get("data_atual", ""),
        "{{ENDERECO}}": company.get("endereco", ""),
        "{{BAIRRO/CIDADE}}": company.get("bairro_cidade", ""),
        "{{CEP}}": company.get("cep", ""),
        "{{CNAE1}}": company.get("cnae", ""),
        "{{DESCRICAO1}}": company.get("descricao_atividade", ""),
        "{{GRAU1}}": company.get("grau_risco", ""),
        "{{CNAE2}}": company.get("cnae_secundario", ""),
        "{{DESCRICAO2}}": company.get("descricao_atividade_secundaria", ""),
        "{{GRAU2}}": company.get("grau_risco_secundario", ""),
        "{{N°FUNCIONARIOS}}": funcionarios,
        "{{EMAIL}}": company.get("email", ""),
        "{{FONE}}": company.get("fone", ""),
        "{{DATACRIACAOLAUDO}}": company.get("data_criacao_laudo") or company.get("data_avaliacao") or company.get("data_atual", ""),
        "{{datacriacaolaudo}}": company.get("data_criacao_laudo") or company.get("data_avaliacao") or company.get("data_atual", ""),
        "{{DATADAREVISAO}}": company.get("data_da_revisao", ""),
        "{{datadarevisao}}": company.get("data_da_revisao", ""),
    }


def _fill_company_identification_tables(doc: Document, company: Mapping[str, str], sectors: list[Mapping[str, Any]]) -> None:
    funcionarios = company.get("funcionarios") or _total_funcionarios_from_sectors(sectors)
    values_by_label = {
        "EMPRESA": company.get("empresa", ""),
        "ENDEREÇO": company.get("endereco", ""),
        "BAIRRO / CIDADE": company.get("bairro_cidade", ""),
        "CEP": company.get("cep", ""),
        "CNPJ": company.get("cnpj", ""),
        "CNAE": company.get("cnae", ""),
        "GRAU DE RISCO": company.get("grau_risco", ""),
        "CNAE (SECUNDÁRIO)": company.get("cnae_secundario", ""),
        "GRAU DE RISCO (SECUNDÁRIO)": company.get("grau_risco_secundario", ""),
        "FUNCIONÁRIOS": funcionarios,
        "VIGÊNCIA": f"{company.get('data_atual','')} – {company.get('data_final','')}",
        "EMAIL": company.get("email", ""),
        "FONE": company.get("fone", ""),
    }
    for table in doc.tables:
        labels = [row.cells[0].text.strip().upper() for row in table.rows if row.cells]
        if "EMPRESA" not in labels or not any(label in labels for label in ("CNPJ", "VIGÊNCIA", "CNAE", "FUNCIONÁRIOS")):
            continue
        descricao_seen = 0
        for row in table.rows:
            if len(row.cells) < 2:
                continue
            label = row.cells[0].text.strip().upper()
            if label == "DESCRIÇÃO DA ATIVIDADE":
                descricao_seen += 1
                value = company.get("descricao_atividade_secundaria", "") if descricao_seen == 2 else company.get("descricao_atividade", "")
                _ltcat_set_cell_text(row.cells[1], value)
            elif label in values_by_label:
                _ltcat_set_cell_text(row.cells[1], values_by_label[label])

def generate_complete_pgr_docx(
    groups_or_risks: Iterable[Mapping[str, Any]],
    output_path: str | Path,
    empresa: str,
    cnpj: str,
    data_atual: str | None = None,
    data_final: str | None = None,
    company_extra: Mapping[str, Any] | None = None,
) -> Path:
    """Gera o PGR completo preenchendo o modelo principal com todos os blocos já criados."""
    items = list(groups_or_risks)
    if not items:
        raise ValueError("Selecione pelo menos um setor e um risco para gerar o PGR completo.")
    if not empresa or not str(empresa).strip():
        raise ValueError("Preencha o nome da empresa para gerar o PGR completo.")
    if not cnpj or not str(cnpj).strip():
        raise ValueError("Preencha o CNPJ da empresa para gerar o PGR completo.")
    if not TEMPLATE_PGR_COMPLETO_PATH.exists():
        raise FileNotFoundError(f"Modelo PGR completo não encontrado: {TEMPLATE_PGR_COMPLETO_PATH}")

    if _is_grouped_payload(items):
        groups = _sanitize_sector_risk_groups(items)
    else:
        groups = [{"sector": {"setor": "", "cargos": []}, "risks": items}]
    if not groups:
        raise ValueError("Selecione pelo menos um setor e um risco para gerar o PGR completo.")

    empresa = str(empresa).strip()
    cnpj = str(cnpj).strip()
    data_atual = (data_atual or datetime.now().strftime("%d/%m/%Y")).strip()
    data_final = (data_final or "").strip()
    company = _company_dict_for_docs(empresa, cnpj, data_atual, data_final, company_extra)

    doc = Document(str(TEMPLATE_PGR_COMPLETO_PATH))

    relation_table = _find_table_xml(doc, ["{{CARGO}}", "{{CBOCARGO}}", "{{N°FUNC}}", "{{DESCRIÇÃO ATIVIDADE}}"])
    descritivo_table = _find_table_xml(doc, ["PAREDE:", "{{SETOR}}", "VENTILAÇÃO:"])
    risco_pgr_table = _find_table_xml(doc, ["{{CARGOS}}", "{{TIPO DE RISCO}}", "{{GRAU DE NIVEL DE RISCO}}"])
    plano_table = _find_table_xml(doc, ["PLANO", "{{risco}}", "{{AÇÕES PREVENTIVA / CORRETIVA}}", "{{INDICADOR DE EFETIVIDADE}}"])

    sectors = _unique_sectors_from_groups(groups)
    if not sectors:
        raise ValueError("Cadastre e selecione pelo menos um setor para gerar o PGR completo.")

    # Substitui os blocos-modelo por blocos gerados, mantendo a posição original dentro do PGR.
    _replace_xml_element_with(relation_table.getparent(), relation_table, _build_relacao_elements(relation_table, sectors, data_atual, data_final))
    _replace_xml_element_with(descritivo_table.getparent(), descritivo_table, _build_descritivo_elements(descritivo_table, sectors))
    _replace_xml_element_with(risco_pgr_table.getparent(), risco_pgr_table, _build_risco_pgr_elements(risco_pgr_table, groups, break_before_first=False))
    _replace_xml_element_with(plano_table.getparent(), plano_table, _build_action_plan_elements(plano_table, groups, data_atual=data_atual, data_final=data_final))

    _fill_company_identification_tables(doc, company, sectors)
    _insert_psychosocial_revision_line(doc, company)
    _replace_doc_placeholders(doc, _company_replacements(company))

    # Ajustes de layout solicitados:
    # - manter o bloco Data do documento/Revisão compacto e ainda na primeira página;
    # - evitar páginas em branco entre o descritivo e o inventário de riscos.
    _compact_first_page_date_block(doc)
    _compact_known_static_spacers(doc)

    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(output_path))
    return output_path


# ---------------------------------------------------------------------------
# PCMSO COMPLETO E BLOCO DE RISCOS/EXAMES
# ---------------------------------------------------------------------------


def _fill_pcmso_risk_rows(risk_rows: list, risk: Mapping[str, Any]) -> None:
    if len(risk_rows) != 4:
        raise ValueError("O bloco de risco do PCMSO precisa manter 4 linhas.")
    tipo = risk.get("tipo_risco", "")
    descricao = risk.get("descricao_agente") or risk.get("risco", "")
    possiveis_lesoes = risk.get("possiveis_lesoes", "")
    fontes = risk.get("fontes_circunstancias") or "Durante o processo de trabalho."

    _colored_value(_row_cell(risk_rows[0], 1), tipo, TIPO_RISCO_COLORS)
    _set_cell_text(_row_cell(risk_rows[1], 1), descricao)
    _set_cell_text(_row_cell(risk_rows[2], 1), possiveis_lesoes)
    _set_cell_text(_row_cell(risk_rows[3], 1), fontes)


def _fill_pcmso_exam_row(row_xml, exam: Mapping[str, Any]) -> None:
    cells = row_xml.findall(qn("w:tc"))
    if len(cells) < 7:
        raise ValueError("A linha de exame do modelo PCMSO precisa manter 7 células.")
    values = [
        exam.get("exame", ""),
        exam.get("periodicidade", ""),
        exam.get("admissional", ""),
        exam.get("periodico", ""),
        exam.get("retorno", ""),
        exam.get("mudanca", ""),
        exam.get("demissional", ""),
    ]
    for cell, value in zip(cells[:7], values):
        _set_cell_text(cell, value)


def _sanitize_exam(exam: Mapping[str, Any]) -> dict[str, Any]:
    return {
        "id": str(exam.get("id", "")).strip(),
        "exame": str(exam.get("exame", "")).strip(),
        "periodicidade": str(exam.get("periodicidade", "")).strip(),
        "admissional": str(exam.get("admissional", "")).strip(),
        "periodico": str(exam.get("periodico", "")).strip(),
        "retorno": str(exam.get("retorno", "")).strip(),
        "mudanca": str(exam.get("mudanca", "")).strip(),
        "demissional": str(exam.get("demissional", "")).strip(),
    }


def _sanitize_pcmso_groups(groups: Iterable[Mapping[str, Any]]) -> list[dict[str, Any]]:
    clean_groups: list[dict[str, Any]] = []
    for group in groups:
        sector = _sanitize_sector(group.get("sector") or {})
        risks = [risk for risk in (group.get("risks") or []) if risk]
        exams = [_sanitize_exam(exam) for exam in (group.get("exams") or []) if exam]
        if sector.get("setor") and risks:
            clean_groups.append({"sector": sector, "risks": risks, "exams": exams})
    return clean_groups


def _fill_pcmso_riscos_sector_table(table_xml, sector: Mapping[str, Any], risks: list[Mapping[str, Any]], exams: list[Mapping[str, Any]]) -> None:
    rows = table_xml.findall(qn("w:tr"))
    if len(rows) < 12:
        raise ValueError("O modelo riscos PCMSO precisa manter 12 linhas.")
    if not risks:
        raise ValueError("Cada setor selecionado precisa ter pelo menos um risco.")

    _set_cell_text(_row_cell(rows[0], 0), sector.get("setor", ""))
    _set_cell_text(_row_cell(rows[2], 0), _sector_cargos_text(sector))

    risk_template_rows = [deepcopy(row) for row in rows[4:8]]
    _fill_pcmso_risk_rows(rows[4:8], risks[0])

    phrase_row = rows[8]
    insertion_index = list(table_xml).index(phrase_row)
    for risk in risks[1:]:
        block_rows = [deepcopy(row) for row in risk_template_rows]
        _fill_pcmso_risk_rows(block_rows, risk)
        for block_row in block_rows:
            table_xml.insert(insertion_index, block_row)
            insertion_index += 1

    has_psychosocial = any(_normalize_option(risk.get("tipo_risco")) == "ERGONÔMICO PSICOSSOCIAL" for risk in risks)
    # A frase fica somente quando NÃO existe risco ERGONÔMICO PSICOSSOCIAL no setor.
    if has_psychosocial:
        table_xml.remove(phrase_row)

    rows = table_xml.findall(qn("w:tr"))
    exam_template_row = None
    for row in rows:
        if "{{exame}}" in _xml_text(row):
            exam_template_row = row
            break
    if exam_template_row is None:
        raise ValueError("O modelo riscos PCMSO precisa ter uma linha de exame com {{exame}}.")
    exam_template_copy = deepcopy(exam_template_row)
    insert_index = list(table_xml).index(exam_template_row)
    table_xml.remove(exam_template_row)

    if not exams:
        exams = [{"exame": "", "periodicidade": "", "admissional": "", "periodico": "", "retorno": "", "mudanca": "", "demissional": ""}]
    for exam in exams:
        row = deepcopy(exam_template_copy)
        _fill_pcmso_exam_row(row, exam)
        table_xml.insert(insert_index, row)
        insert_index += 1


def _build_pcmso_riscos_elements(template_table_xml, groups: list[Mapping[str, Any]], page_break_between_sectors: bool = True) -> list:
    elements: list = []
    for index, group in enumerate(groups):
        if index > 0 and page_break_between_sectors:
            elements.append(_page_break_paragraph())
        table = deepcopy(template_table_xml)
        _fill_pcmso_riscos_sector_table(table, group["sector"], group["risks"], group.get("exams", []))
        elements.append(table)
    return elements


def generate_riscos_pcmso_docx(groups: Iterable[Mapping[str, Any]], output_path: str | Path) -> Path:
    groups = _sanitize_pcmso_groups(groups)
    if not groups:
        raise ValueError("Selecione pelo menos um setor e um risco para gerar o bloco de riscos do PCMSO.")
    if not TEMPLATE_RISCOS_PCMSO_PATH.exists():
        raise FileNotFoundError(f"Modelo riscos PCMSO não encontrado: {TEMPLATE_RISCOS_PCMSO_PATH}")
    doc = Document(str(TEMPLATE_RISCOS_PCMSO_PATH))
    if not doc.tables:
        raise ValueError("O modelo riscos PCMSO precisa ter uma tabela-modelo.")
    body = doc._body._element
    original_table_xml = doc.tables[0]._tbl
    template_table_copy = deepcopy(original_table_xml)
    body.remove(original_table_xml)
    for element in _build_pcmso_riscos_elements(template_table_copy, groups):
        _insert_before_section_or_end(body, element)
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(output_path))
    return output_path


def _remove_element(element) -> None:
    parent = element.getparent()
    if parent is not None:
        parent.remove(element)


def _replace_pcmso_riscos_area(doc: Document, groups: list[Mapping[str, Any]]) -> None:
    risco_table = _find_table_xml(doc, ["{{CARGOS}}", "{{TIPO DE RISCO}}", "{{risco}}"])
    phrase_table = _find_table_xml(doc, ["NENHUM FATOR DE RISCO PSICOSSOCIAL"])
    exams_table = _find_table_xml(doc, ["EXAMES RECOMENDADOS", "{{tempo de peridiocidade}}"])

    template_doc = Document(str(TEMPLATE_RISCOS_PCMSO_PATH))
    template_table = deepcopy(template_doc.tables[0]._tbl)
    generated = [_page_break_paragraph()] + _build_pcmso_riscos_elements(template_table, groups)

    parent = risco_table.getparent()
    index = list(parent).index(risco_table)
    _remove_element(risco_table)
    _remove_element(phrase_table)
    _remove_element(exams_table)
    for offset, element in enumerate(generated):
        parent.insert(index + offset, element)


def generate_complete_pcmso_docx(
    groups: Iterable[Mapping[str, Any]],
    output_path: str | Path,
    empresa: str,
    cnpj: str,
    data_atual: str | None = None,
    data_final: str | None = None,
    company_extra: Mapping[str, Any] | None = None,
) -> Path:
    groups = _sanitize_pcmso_groups(groups)
    if not groups:
        raise ValueError("Selecione pelo menos um setor e um risco para gerar o PCMSO completo.")
    empresa = str(empresa or "").strip()
    cnpj = str(cnpj or "").strip()
    if not empresa:
        raise ValueError("Preencha o nome da empresa.")
    if not cnpj:
        raise ValueError("Preencha o CNPJ.")
    if not TEMPLATE_PCMSO_COMPLETO_PATH.exists():
        raise FileNotFoundError(f"Modelo PCMSO completo não encontrado: {TEMPLATE_PCMSO_COMPLETO_PATH}")
    data_atual = (data_atual or datetime.now().strftime("%m/%Y")).strip()
    data_final = (data_final or "").strip()

    company = _company_dict_for_docs(empresa, cnpj, data_atual, data_final, company_extra)
    doc = Document(str(TEMPLATE_PCMSO_COMPLETO_PATH))
    relation_table = _find_table_xml(doc, ["{{CARGO}}", "{{CBOCARGO}}", "{{N°FUNC}}", "{{DESCRIÇÃO ATIVIDADE}}"])
    sectors = _unique_sectors_from_groups(groups)
    _replace_xml_element_with(relation_table.getparent(), relation_table, _build_relacao_elements(relation_table, sectors, data_atual, data_final))
    _replace_pcmso_riscos_area(doc, groups)

    _fill_company_identification_tables(doc, company, sectors)
    _insert_psychosocial_revision_line(doc, company)
    _replace_doc_placeholders(doc, _company_replacements(company))
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(output_path))
    return output_path


# ---------------------------------------------------------------------------
# LTCAT COMPLETO
# ---------------------------------------------------------------------------

TIPOS_RISCO_LTCAT = {"FÍSICO", "QUÍMICO", "BIOLÓGICO"}
DEFAULT_LTCAT_ENQUADRAMENTO = (
    "Considerando a atividade que pode causar é recomendado adotar medidas preventivas "
    "até que as avaliações quantitativas sejam realizadas."
)
DEFAULT_LTCAT_PARECER = (
    "Considerando a atividade que pode causar é recomendado adotar medidas preventivas "
    "até que as avaliações quantitativas sejam realizadas."
)


def _ltcat_table_has_text(table, text: str) -> bool:
    return text in "\n".join(cell.text for row in table.rows for cell in row.cells)


def _ltcat_replace_in_table(table, replacements: Mapping[str, str]) -> None:
    clean = {key: "" if value is None else str(value) for key, value in replacements.items()}
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                _replace_text_in_paragraph(paragraph, clean)


def _ltcat_set_cell_text(cell, text_value: Any) -> None:
    text_value = "" if text_value is None else str(text_value)
    if not cell.paragraphs:
        paragraph = cell.add_paragraph()
        paragraph.add_run(text_value)
        return
    paragraph = cell.paragraphs[0]
    if paragraph.runs:
        paragraph.runs[0].text = text_value
        for run in paragraph.runs[1:]:
            run.text = ""
    else:
        paragraph.add_run(text_value)
    for extra in cell.paragraphs[1:]:
        for run in extra.runs:
            run.text = ""


def _ltcat_shade_cell(cell, fill: str | None) -> None:
    if not fill:
        return
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tc_pr.append(shd)
    shd.set(qn("w:fill"), fill)


def _ltcat_set_cell_font_color(cell, rgb: str = "000000") -> None:
    for paragraph in cell.paragraphs:
        for run in paragraph.runs:
            run.font.color.rgb = RGBColor.from_string(rgb)


def _ltcat_insert_tbl_before(anchor_table, tbl_element) -> None:
    anchor_table._tbl.addprevious(tbl_element)


def _ltcat_remove_table(table) -> None:
    table._tbl.getparent().remove(table._tbl)


def _ltcat_insert_blank_before(anchor_table) -> None:
    p = OxmlElement("w:p")
    p_pr = OxmlElement("w:pPr")
    spacing = OxmlElement("w:spacing")
    spacing.set(qn("w:before"), "0")
    spacing.set(qn("w:after"), "0")
    p_pr.append(spacing)
    p.append(p_pr)
    anchor_table._tbl.addprevious(p)


def _ltcat_insert_page_break_before(anchor_table) -> None:
    p = OxmlElement("w:p")
    p_pr = OxmlElement("w:pPr")
    spacing = OxmlElement("w:spacing")
    spacing.set(qn("w:before"), "0")
    spacing.set(qn("w:after"), "0")
    p_pr.append(spacing)
    p.append(p_pr)
    r = OxmlElement("w:r")
    br = OxmlElement("w:br")
    br.set(qn("w:type"), "page")
    r.append(br)
    p.append(r)
    anchor_table._tbl.addprevious(p)


def _ltcat_sector_cargos_text(sector: Mapping[str, Any]) -> str:
    return _sector_cargos_text(sector).upper()


def _ltcat_company_dict(empresa: str, cnpj: str, data_atual: str, data_final: str, extra: Mapping[str, Any] | None = None) -> dict[str, str]:
    extra = extra or {}
    data_criacao = str(extra.get("data_criacao_laudo") or extra.get("datacriacaolaudo") or extra.get("data_avaliacao") or "").strip()
    return {
        "empresa": str(empresa or "").strip(),
        "cnpj": str(cnpj or "").strip(),
        "data_atual": str(data_atual or "").strip(),
        "data_final": str(data_final or "").strip(),
        "endereco": str(extra.get("endereco", "") or "").strip(),
        "bairro_cidade": str(extra.get("bairro_cidade", "") or "").strip(),
        "cep": str(extra.get("cep", "") or "").strip(),
        "cnae": str(extra.get("cnae", extra.get("cnae1", "")) or "").strip(),
        "descricao_atividade": str(extra.get("descricao_atividade", extra.get("descricao1", "")) or "").strip(),
        "grau_risco": str(extra.get("grau_risco", extra.get("grau1", "")) or "").strip(),
        "cnae_secundario": str(extra.get("cnae_secundario", extra.get("cnae2", "")) or "").strip(),
        "descricao_atividade_secundaria": str(extra.get("descricao_atividade_secundaria", extra.get("descricao2", "")) or "").strip(),
        "grau_risco_secundario": str(extra.get("grau_risco_secundario", extra.get("grau2", "")) or "").strip(),
        "funcionarios": str(extra.get("funcionarios", "") or "").strip(),
        "email": str(extra.get("email", "") or "").strip(),
        "fone": str(extra.get("fone", "") or "").strip(),
        "data_avaliacao": data_criacao or str(data_atual or "").strip(),
        "data_criacao_laudo": data_criacao,
        "ajuste_psicossocial": str(extra.get("ajuste_psicossocial", "") or "").strip(),
        "data_da_revisao": str(extra.get("data_da_revisao", extra.get("data_revisao_ajuste", "")) or "").strip(),
    }


def _ltcat_total_funcionarios(sectors: list[Mapping[str, Any]]) -> str:
    total = 0
    for sector in sectors:
        for cargo in sector.get("cargos", []) or []:
            try:
                total += int(str(cargo.get("n_func", "")).strip() or "0")
            except ValueError:
                pass
    return str(total) if total else ""


def _ltcat_fill_company_table(doc: Document, company: Mapping[str, str], sectors: list[Mapping[str, Any]]) -> None:
    for table in doc.tables:
        if len(table.rows) >= 14 and table.cell(0, 0).text.strip().upper() == "EMPRESA":
            funcionarios = company.get("funcionarios") or _ltcat_total_funcionarios(sectors)
            values_by_label = {
                "EMPRESA": company.get("empresa", ""),
                "ENDEREÇO": company.get("endereco", ""),
                "BAIRRO / CIDADE": company.get("bairro_cidade", ""),
                "CEP": company.get("cep", ""),
                "CNPJ": company.get("cnpj", ""),
                "CNAE": company.get("cnae", ""),
                "GRAU DE RISCO": company.get("grau_risco", ""),
                "CNAE (SECUNDÁRIO)": company.get("cnae_secundario", ""),
                "GRAU DE RISCO (SECUNDÁRIO)": company.get("grau_risco_secundario", ""),
                "FUNCIONÁRIOS": funcionarios,
                "VIGÊNCIA": f"{company.get('data_atual','')} – {company.get('data_final','')}",
                "EMAIL": company.get("email", ""),
                "FONE": company.get("fone", ""),
            }
            descricao_seen = 0
            for row in table.rows:
                label = row.cells[0].text.strip().upper()
                if label == "DESCRIÇÃO DA ATIVIDADE":
                    descricao_seen += 1
                    value = company.get("descricao_atividade_secundaria", "") if descricao_seen == 2 else company.get("descricao_atividade", "")
                    _ltcat_set_cell_text(row.cells[1], value)
                elif label in values_by_label:
                    _ltcat_set_cell_text(row.cells[1], values_by_label[label])
            break


def _ltcat_fill_relacao_funcoes(doc: Document, sectors: list[Mapping[str, Any]], data_atual: str, data_final: str) -> None:
    relation_tables = [t for t in doc.tables if _ltcat_table_has_text(t, "{{CARGO}}") and _ltcat_table_has_text(t, "{{CBOCARGO}}")]
    if not relation_tables:
        return
    anchor = relation_tables[0]
    template_tbl = deepcopy(anchor._tbl)
    for index, sector in enumerate(sectors):
        new_tbl = deepcopy(template_tbl)
        _ltcat_insert_tbl_before(anchor, new_tbl)
        table = Table(new_tbl, doc)
        _fill_relacao_table(table._tbl, sector, data_atual, data_final)
        if index < len(sectors) - 1:
            _ltcat_insert_blank_before(anchor)
    _ltcat_remove_table(anchor)


def _ltcat_fill_descritivo_setores(doc: Document, sectors: list[Mapping[str, Any]]) -> None:
    desc_tables = [t for t in doc.tables if _ltcat_table_has_text(t, "PAREDE:") and _ltcat_table_has_text(t, "{{SETOR}}")]
    if not desc_tables:
        return
    anchor = desc_tables[0]
    template_tbl = deepcopy(anchor._tbl)
    for index, sector in enumerate(sectors):
        new_tbl = deepcopy(template_tbl)
        _ltcat_insert_tbl_before(anchor, new_tbl)
        table = Table(new_tbl, doc)
        _ltcat_replace_in_table(table, {"{{SETOR}}": str(sector.get("setor", ""))})
        if index < len(sectors) - 1:
            _ltcat_insert_blank_before(anchor)
    _ltcat_remove_table(anchor)


def _ltcat_risk_from_common(risk: Mapping[str, Any]) -> dict[str, Any]:
    return {
        "id": str(risk.get("id", "")),
        "risco": str(risk.get("risco", "") or ""),
        "tipo_risco": _normalize_option(risk.get("tipo_risco")),
        "meio_propagacao": str(risk.get("ltcat_meio_propagacao") or risk.get("fontes_circunstancias") or ""),
        "epis": str(risk.get("epis", "") or ""),
        "epcs": str(risk.get("epcs", "") or ""),
        "insalubridade": str(risk.get("ltcat_insalubridade") or "Não"),
        "grau_insalubridade": str(risk.get("ltcat_grau_insalubridade") or "Não aplicável"),
        "aposentadoria_especial": str(risk.get("ltcat_aposentadoria_especial") or "Não"),
        "enquadramento_tecnico": str(risk.get("ltcat_enquadramento_tecnico") or DEFAULT_LTCAT_ENQUADRAMENTO),
        "parecer_previdenciario": str(risk.get("ltcat_parecer_previdenciario") or DEFAULT_LTCAT_PARECER),
        "periodicidade_jornada": str(risk.get("ltcat_periodicidade_jornada") or "Mensal (<= 4 horas < 10% jornada)"),
    }


def _ltcat_set_repeated_cells(row, col_start: int, col_end: int, text_value: Any) -> None:
    for col in range(col_start, min(col_end + 1, len(row.cells))):
        _ltcat_set_cell_text(row.cells[col], text_value)


def _ltcat_fill_single_risk_block(table, start: int, risk: Mapping[str, Any], data_avaliacao: str) -> None:
    tipo = _normalize_option(risk.get("tipo_risco"))
    fill = TIPO_RISCO_COLORS.get(tipo)

    def row(offset: int):
        return table.rows[start + offset]

    _ltcat_set_repeated_cells(row(0), 3, 5, tipo)
    for col in range(3, min(6, len(row(0).cells))):
        _ltcat_shade_cell(row(0).cells[col], fill)
        _ltcat_set_cell_font_color(row(0).cells[col], "000000")

    _ltcat_set_repeated_cells(row(1), 3, 5, risk.get("risco", ""))
    _ltcat_set_repeated_cells(row(2), 3, 5, risk.get("meio_propagacao", ""))
    _ltcat_set_repeated_cells(row(8), 3, 5, risk.get("epis", ""))
    _ltcat_set_repeated_cells(row(9), 3, 5, risk.get("epcs", ""))

    if risk.get("periodicidade_jornada") and len(row(12).cells) > 1:
        _ltcat_set_cell_text(row(12).cells[1], risk.get("periodicidade_jornada", ""))
    _ltcat_set_repeated_cells(row(12), 5, 5, data_avaliacao)

    _ltcat_set_repeated_cells(row(14), 2, 5, risk.get("insalubridade", ""))
    _ltcat_set_repeated_cells(row(15), 2, 5, risk.get("grau_insalubridade", ""))
    _ltcat_set_repeated_cells(row(16), 2, 5, risk.get("enquadramento_tecnico", ""))
    _ltcat_set_repeated_cells(row(18), 2, 5, risk.get("aposentadoria_especial", ""))
    _ltcat_set_repeated_cells(row(19), 2, 5, risk.get("parecer_previdenciario", ""))

    replacements = {
        "{{TIPO DE RISCO}}": tipo,
        "{{risco}}": risk.get("risco", ""),
        "{{meiodepropagacao}}": risk.get("meio_propagacao", ""),
        "{{EPIs}}": risk.get("epis", ""),
        "{{EPCs}}": risk.get("epcs", ""),
        "{{datacriacaolaudo}}": data_avaliacao,
        "{{simounao}}": risk.get("insalubridade", ""),
        "{{graudeinsalubridade}}": risk.get("grau_insalubridade", ""),
    }
    for off in range(0, 25):
        if start + off >= len(table.rows):
            continue
        for cell in table.rows[start + off].cells:
            for paragraph in cell.paragraphs:
                _replace_text_in_paragraph(paragraph, replacements)


def _ltcat_fill_risk_table(table, sector: Mapping[str, Any], risks: list[Mapping[str, Any]], data_avaliacao: str) -> None:
    _ltcat_replace_in_table(table, {
        "{{SETOR}}": str(sector.get("setor", "")),
        "{{CARGOS}}": _ltcat_sector_cargos_text(sector),
    })
    original_block_rows = [deepcopy(table.rows[i]._tr) for i in range(4, len(table.rows))]
    for _ in range(max(0, len(risks) - 1)):
        for tr in original_block_rows:
            table._tbl.append(deepcopy(tr))
    block_size = len(original_block_rows)
    for index, risk in enumerate(risks):
        _ltcat_fill_single_risk_block(table, 4 + index * block_size, risk, data_avaliacao)


def _ltcat_fill_riscos_area(doc: Document, groups: list[Mapping[str, Any]], data_avaliacao: str) -> None:
    absence_tables = [t for t in doc.tables if _ltcat_table_has_text(t, "AUSÊNCIA DE RISCOS") and _ltcat_table_has_text(t, "{{SETOR}}")]
    risk_tables = [t for t in doc.tables if _ltcat_table_has_text(t, "{{meiodepropagacao}}") and _ltcat_table_has_text(t, "{{graudeinsalubridade}}")]
    if not absence_tables or not risk_tables:
        return
    absence_anchor = absence_tables[0]
    risk_anchor = risk_tables[0]
    absence_template = deepcopy(absence_anchor._tbl)
    risk_template = deepcopy(risk_anchor._tbl)

    for index, group in enumerate(groups):
        sector = group["sector"]
        if index > 0:
            _ltcat_insert_page_break_before(absence_anchor)
        raw_risks = group.get("risks") or []
        risks = [_ltcat_risk_from_common(risk) for risk in raw_risks]
        risks = [risk for risk in risks if _normalize_option(risk.get("tipo_risco")) in TIPOS_RISCO_LTCAT]
        if risks:
            new_tbl = deepcopy(risk_template)
            _ltcat_insert_tbl_before(absence_anchor, new_tbl)
            table = Table(new_tbl, doc)
            _ltcat_fill_risk_table(table, sector, risks, data_avaliacao)
        else:
            new_tbl = deepcopy(absence_template)
            _ltcat_insert_tbl_before(absence_anchor, new_tbl)
            table = Table(new_tbl, doc)
            _ltcat_replace_in_table(table, {
                "{{SETOR}}": str(sector.get("setor", "")),
                "{{CARGOS}}": _ltcat_sector_cargos_text(sector),
            })
    _ltcat_remove_table(absence_anchor)
    _ltcat_remove_table(risk_anchor)


def _ltcat_fill_signature(doc: Document, company: Mapping[str, str]) -> None:
    for table in doc.tables:
        if _ltcat_table_has_text(table, "RESPONSÁVEL LEGAL DA EMPRESA") and _ltcat_table_has_text(table, "{{empresa}}"):
            _ltcat_replace_in_table(table, {
                "{{empresa}}": company.get("empresa", ""),
                "{{DATACRIACAOLAUDO}}": company.get("data_criacao_laudo") or company.get("data_avaliacao") or company.get("data_atual", ""),
                "{{datacriacaolaudo}}": company.get("data_criacao_laudo") or company.get("data_avaliacao") or company.get("data_atual", ""),
            })


def _sanitize_ltcat_groups(groups: Iterable[Mapping[str, Any]]) -> list[dict[str, Any]]:
    clean_groups: list[dict[str, Any]] = []
    for group in groups:
        sector = _sanitize_sector(group.get("sector") or {})
        risks = [risk for risk in (group.get("risks") or []) if risk]
        if sector.get("setor"):
            clean_groups.append({"sector": sector, "risks": risks})
    return clean_groups


def generate_complete_ltcat_docx(
    groups: Iterable[Mapping[str, Any]],
    output_path: str | Path,
    empresa: str,
    cnpj: str,
    data_atual: str | None = None,
    data_final: str | None = None,
    company_extra: Mapping[str, Any] | None = None,
) -> Path:
    groups = _sanitize_ltcat_groups(groups)
    if not groups:
        raise ValueError("Selecione pelo menos um setor para gerar o LTCAT completo.")
    empresa = str(empresa or "").strip()
    cnpj = str(cnpj or "").strip()
    if not empresa:
        raise ValueError("Preencha o nome da empresa.")
    if not cnpj:
        raise ValueError("Preencha o CNPJ.")
    if not TEMPLATE_LTCAT_COMPLETO_PATH.exists():
        raise FileNotFoundError(f"Modelo LTCAT completo não encontrado: {TEMPLATE_LTCAT_COMPLETO_PATH}")

    data_atual = (data_atual or datetime.now().strftime("%m/%Y")).strip()
    data_final = (data_final or "").strip()
    company = _ltcat_company_dict(empresa, cnpj, data_atual, data_final, company_extra)

    doc = Document(str(TEMPLATE_LTCAT_COMPLETO_PATH))
    sectors = [group["sector"] for group in groups]
    data_avaliacao = company.get("data_avaliacao") or data_atual

    _insert_psychosocial_revision_line(doc, company)
    _replace_doc_placeholders(doc, _company_replacements(company))
    _ltcat_fill_company_table(doc, company, sectors)
    _ltcat_fill_relacao_funcoes(doc, sectors, data_atual, data_final)
    _ltcat_fill_descritivo_setores(doc, sectors)
    _ltcat_fill_riscos_area(doc, groups, data_avaliacao)
    _ltcat_fill_signature(doc, company)

    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(output_path))
    return output_path
