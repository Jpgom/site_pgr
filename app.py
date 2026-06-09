from __future__ import annotations

import json
import os
import uuid
from datetime import datetime
from pathlib import Path
from typing import Any

from flask import Flask, flash, redirect, render_template, request, send_file, url_for
from flask_sqlalchemy import SQLAlchemy
from sqlalchemy import Text, inspect, text

from word_generator import (
    NIVEL_RISCO_COLORS,
    POSSIBILIDADE_COLORS,
    SEVERIDADE_COLORS,
    TIPO_RISCO_COLORS,
    generate_action_plan_docx,
    generate_complete_pgr_docx,
    generate_complete_pcmso_docx,
    generate_complete_ltcat_docx,
    generate_descritivo_setor_docx,
    generate_pcmso_docx,
    generate_riscos_pcmso_docx,
    generate_pgr_docx,
    generate_relacao_funcao_atividade_docx,
)

BASE_DIR = Path(__file__).resolve().parent
DATA_FILE = BASE_DIR / "data" / "riscos.json"
SETORES_FILE = BASE_DIR / "data" / "setores.json"
EXAMES_FILE = BASE_DIR / "data" / "exames.json"
OUTPUT_DIR = BASE_DIR / "outputs"
INSTANCE_DIR = BASE_DIR / "instance"


def _database_uri() -> str:
    """Usa PostgreSQL no Render via DATABASE_URL e SQLite local como fallback."""
    url = os.environ.get("DATABASE_URL", "").strip()
    if url.startswith("postgres://"):
        url = url.replace("postgres://", "postgresql+psycopg://", 1)
    elif url.startswith("postgresql://"):
        url = url.replace("postgresql://", "postgresql+psycopg://", 1)
    if url:
        return url
    INSTANCE_DIR.mkdir(parents=True, exist_ok=True)
    return f"sqlite:///{INSTANCE_DIR / 'sst_riscos.db'}"


app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "troque-esta-chave-em-producao")
app.config["SQLALCHEMY_DATABASE_URI"] = _database_uri()
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False
app.config["SQLALCHEMY_ENGINE_OPTIONS"] = {"pool_pre_ping": True}

db = SQLAlchemy(app)


class Risk(db.Model):
    __tablename__ = "risks"

    id = db.Column(db.String(32), primary_key=True, default=lambda: uuid.uuid4().hex)
    risco = db.Column(db.String(255), nullable=False)
    acoes = db.Column(Text, nullable=False)
    indicador = db.Column(Text, nullable=False)
    tipo_risco = db.Column(db.String(80), nullable=False)
    descricao_agente = db.Column(Text, default="")
    possiveis_lesoes = db.Column(Text, nullable=False)
    fontes_circunstancias = db.Column(Text, nullable=False, default="Durante o processo de trabalho.")
    epis = db.Column(Text, nullable=False)
    epcs = db.Column(Text, nullable=False)
    ltcat_meio_propagacao = db.Column(Text, default="")
    ltcat_insalubridade = db.Column(db.String(80), default="Não")
    ltcat_grau_insalubridade = db.Column(db.String(120), default="Não aplicável")
    ltcat_aposentadoria_especial = db.Column(db.String(80), default="Não")
    ltcat_enquadramento_tecnico = db.Column(Text, default="")
    ltcat_parecer_previdenciario = db.Column(Text, default="")
    ltcat_periodicidade_jornada = db.Column(Text, default="Mensal (<= 4 horas < 10% jornada)")
    grau_severidade = db.Column(db.String(80), nullable=False)
    grau_possibilidade = db.Column(db.String(80), nullable=False)
    grau_nivel_risco = db.Column(db.String(80), nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow, nullable=False)

    def to_dict(self) -> dict[str, Any]:
        return {
            "id": self.id,
            "risco": self.risco,
            "acoes": self.acoes,
            "prazo_implantacao": "",
            "prazo_reavaliacao": "",
            "indicador": self.indicador,
            "tipo_risco": self.tipo_risco,
            "descricao_agente": self.descricao_agente or "",
            "possiveis_lesoes": self.possiveis_lesoes,
            "fontes_circunstancias": self.fontes_circunstancias or "Durante o processo de trabalho.",
            "epis": self.epis,
            "epcs": self.epcs,
            "ltcat_meio_propagacao": self.ltcat_meio_propagacao or "",
            "ltcat_insalubridade": self.ltcat_insalubridade or "Não",
            "ltcat_grau_insalubridade": self.ltcat_grau_insalubridade or "Não aplicável",
            "ltcat_aposentadoria_especial": self.ltcat_aposentadoria_especial or "Não",
            "ltcat_enquadramento_tecnico": self.ltcat_enquadramento_tecnico or "",
            "ltcat_parecer_previdenciario": self.ltcat_parecer_previdenciario or "",
            "ltcat_periodicidade_jornada": self.ltcat_periodicidade_jornada or "Mensal (<= 4 horas < 10% jornada)",
            "grau_severidade": self.grau_severidade,
            "grau_possibilidade": self.grau_possibilidade,
            "grau_nivel_risco": self.grau_nivel_risco,
            "created_at": self.created_at.strftime("%Y-%m-%d %H:%M:%S") if self.created_at else "",
            "updated_at": self.updated_at.strftime("%Y-%m-%d %H:%M:%S") if self.updated_at else "",
        }


class SectorGroup(db.Model):
    __tablename__ = "sector_groups"

    id = db.Column(db.String(32), primary_key=True, default=lambda: uuid.uuid4().hex)
    nome = db.Column(db.String(255), nullable=False, unique=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow, nullable=False)

    def to_dict(self) -> dict[str, Any]:
        return {
            "id": self.id,
            "nome": self.nome,
            "created_at": self.created_at.strftime("%Y-%m-%d %H:%M:%S") if self.created_at else "",
            "updated_at": self.updated_at.strftime("%Y-%m-%d %H:%M:%S") if self.updated_at else "",
        }


class Sector(db.Model):
    __tablename__ = "sectors"

    id = db.Column(db.String(32), primary_key=True, default=lambda: uuid.uuid4().hex)
    setor = db.Column(db.String(255), nullable=False)
    group_id = db.Column(db.String(32), nullable=True)
    cargos = db.Column(db.JSON, nullable=False, default=list)
    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow, nullable=False)

    def to_dict(self) -> dict[str, Any]:
        cargos = self.cargos or []
        group = db.session.get(SectorGroup, self.group_id) if self.group_id else None
        return {
            "id": self.id,
            "setor": self.setor,
            "grupo_id": self.group_id or "",
            "grupo_nome": group.nome if group else "Sem grupo",
            "cargos": cargos if isinstance(cargos, list) else [],
            "created_at": self.created_at.strftime("%Y-%m-%d %H:%M:%S") if self.created_at else "",
            "updated_at": self.updated_at.strftime("%Y-%m-%d %H:%M:%S") if self.updated_at else "",
        }


class Exam(db.Model):
    __tablename__ = "exams"

    id = db.Column(db.String(32), primary_key=True, default=lambda: uuid.uuid4().hex)
    exame = db.Column(db.String(255), nullable=False)
    periodicidade = db.Column(db.String(120), default="")
    admissional = db.Column(db.String(120), default="")
    periodico = db.Column(db.String(120), default="")
    retorno = db.Column(db.String(120), default="")
    mudanca = db.Column(db.String(120), default="")
    demissional = db.Column(db.String(120), default="")
    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow, nullable=False)

    def to_dict(self) -> dict[str, Any]:
        return {
            "id": self.id,
            "exame": self.exame,
            "periodicidade": self.periodicidade or "",
            "admissional": self.admissional or "",
            "periodico": self.periodico or "",
            "retorno": self.retorno or "",
            "mudanca": self.mudanca or "",
            "demissional": self.demissional or "",
            "created_at": self.created_at.strftime("%Y-%m-%d %H:%M:%S") if self.created_at else "",
            "updated_at": self.updated_at.strftime("%Y-%m-%d %H:%M:%S") if self.updated_at else "",
        }


class Company(db.Model):
    __tablename__ = "companies"

    id = db.Column(db.String(32), primary_key=True, default=lambda: uuid.uuid4().hex)
    nome = db.Column(db.String(255), nullable=False)
    cnpj = db.Column(db.String(40), nullable=False)
    endereco = db.Column(db.String(255), default="")
    bairro_cidade = db.Column(db.String(255), default="")
    cep = db.Column(db.String(40), default="")
    cnae1 = db.Column(db.String(80), default="")
    descricao1 = db.Column(Text, default="")
    grau1 = db.Column(db.String(40), default="")
    cnae2 = db.Column(db.String(80), default="")
    descricao2 = db.Column(Text, default="")
    grau2 = db.Column(db.String(40), default="")
    funcionarios = db.Column(db.String(40), default="")
    data_atual = db.Column(db.String(40), default="")
    data_final = db.Column(db.String(40), default="")
    email = db.Column(db.String(255), default="")
    fone = db.Column(db.String(80), default="")
    data_avaliacao = db.Column(db.String(40), default="")
    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow, nullable=False)

    def to_dict(self) -> dict[str, Any]:
        return {
            "id": self.id,
            "nome": self.nome,
            "empresa": self.nome,
            "cnpj": self.cnpj,
            "endereco": self.endereco or "",
            "bairro_cidade": self.bairro_cidade or "",
            "cep": self.cep or "",
            "cnae1": self.cnae1 or "",
            "descricao1": self.descricao1 or "",
            "grau1": self.grau1 or "",
            "cnae2": self.cnae2 or "",
            "descricao2": self.descricao2 or "",
            "grau2": self.grau2 or "",
            "funcionarios": self.funcionarios or "",
            "data_atual": self.data_atual or "",
            "data_final": self.data_final or "",
            "email": self.email or "",
            "fone": self.fone or "",
            "data_avaliacao": self.data_avaliacao or "",
            # aliases usados pelos geradores antigos
            "cnae": self.cnae1 or "",
            "descricao_atividade": self.descricao1 or "",
            "grau_risco": self.grau1 or "",
            "cnae_secundario": self.cnae2 or "",
            "descricao_atividade_secundaria": self.descricao2 or "",
            "grau_risco_secundario": self.grau2 or "",
            "created_at": self.created_at.strftime("%Y-%m-%d %H:%M:%S") if self.created_at else "",
            "updated_at": self.updated_at.strftime("%Y-%m-%d %H:%M:%S") if self.updated_at else "",
        }


FORM_OPTIONS = {
    "tipos_risco": list(TIPO_RISCO_COLORS.keys()),
    "severidades": list(SEVERIDADE_COLORS.keys()),
    "possibilidades": list(POSSIBILIDADE_COLORS.keys()),
    "niveis_risco": list(NIVEL_RISCO_COLORS.keys()),
}


def _read_json_list(path: Path) -> list[dict[str, Any]]:
    if not path.exists():
        return []
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        return data if isinstance(data, list) else []
    except json.JSONDecodeError:
        return []


def _migrate_json_files_if_needed() -> None:
    """Importa cadastros das versões antigas em JSON somente se o banco estiver vazio."""
    if Risk.query.count() == 0:
        for item in _read_json_list(DATA_FILE):
            risk = Risk(
                id=item.get("id") or uuid.uuid4().hex,
                risco=str(item.get("risco", "")).strip(),
                acoes=str(item.get("acoes", "")).strip(),
                indicador=str(item.get("indicador", "")).strip(),
                tipo_risco=str(item.get("tipo_risco", "")).strip(),
                descricao_agente=str(item.get("descricao_agente", "")).strip(),
                possiveis_lesoes=str(item.get("possiveis_lesoes", "")).strip(),
                fontes_circunstancias=str(item.get("fontes_circunstancias", "Durante o processo de trabalho.")).strip(),
                epis=str(item.get("epis", "")).strip(),
                epcs=str(item.get("epcs", "")).strip(),
                grau_severidade=str(item.get("grau_severidade", "")).strip(),
                grau_possibilidade=str(item.get("grau_possibilidade", "")).strip(),
                grau_nivel_risco=str(item.get("grau_nivel_risco", "")).strip(),
            )
            if risk.risco:
                db.session.merge(risk)

    if Sector.query.count() == 0:
        for item in _read_json_list(SETORES_FILE):
            sector = Sector(
                id=item.get("id") or uuid.uuid4().hex,
                setor=str(item.get("setor", "")).strip(),
                cargos=item.get("cargos") if isinstance(item.get("cargos"), list) else [],
            )
            if sector.setor:
                db.session.merge(sector)

    if Exam.query.count() == 0:
        for item in _read_json_list(EXAMES_FILE):
            exam = Exam(
                id=item.get("id") or uuid.uuid4().hex,
                exame=str(item.get("exame", "")).strip(),
                periodicidade=str(item.get("periodicidade", "")).strip(),
                admissional=str(item.get("admissional", "")).strip(),
                periodico=str(item.get("periodico", "")).strip(),
                retorno=str(item.get("retorno", "")).strip(),
                mudanca=str(item.get("mudanca", "")).strip(),
                demissional=str(item.get("demissional", "")).strip(),
            )
            if exam.exame:
                db.session.merge(exam)

    db.session.commit()


def _ensure_schema_columns() -> None:
    """Adiciona colunas novas em bancos já existentes sem apagar dados."""
    inspector = inspect(db.engine)
    tables = inspector.get_table_names()

    def has_column(table_name: str, column_name: str) -> bool:
        if table_name not in tables:
            return False
        return column_name in {col["name"] for col in inspector.get_columns(table_name)}

    dialect = db.engine.dialect.name

    def add_column(table_name: str, column_sqlite: str, column_pg: str | None = None) -> None:
        column_name = column_sqlite.split()[0]
        if has_column(table_name, column_name):
            return
        if dialect == "postgresql":
            stmt = f"ALTER TABLE {table_name} ADD COLUMN IF NOT EXISTS {column_pg or column_sqlite}"
        else:
            stmt = f"ALTER TABLE {table_name} ADD COLUMN {column_sqlite}"
        db.session.execute(text(stmt))

    add_column("sectors", "group_id VARCHAR(32)")
    add_column("risks", "ltcat_meio_propagacao TEXT")
    add_column("risks", "ltcat_insalubridade VARCHAR(80)")
    add_column("risks", "ltcat_grau_insalubridade VARCHAR(120)")
    add_column("risks", "ltcat_aposentadoria_especial VARCHAR(80)")
    add_column("risks", "ltcat_enquadramento_tecnico TEXT")
    add_column("risks", "ltcat_parecer_previdenciario TEXT")
    add_column("risks", "ltcat_periodicidade_jornada TEXT")
    # Evolução segura do banco para o cadastro completo de empresas.
    for col in [
        "nome VARCHAR(255)", "cnpj VARCHAR(40)", "endereco VARCHAR(255)", "bairro_cidade VARCHAR(255)",
        "cep VARCHAR(40)", "cnae1 VARCHAR(80)", "descricao1 TEXT", "grau1 VARCHAR(40)",
        "cnae2 VARCHAR(80)", "descricao2 TEXT", "grau2 VARCHAR(40)", "funcionarios VARCHAR(40)",
        "data_atual VARCHAR(40)", "data_final VARCHAR(40)", "email VARCHAR(255)", "fone VARCHAR(80)",
        "data_avaliacao VARCHAR(40)",
    ]:
        add_column("companies", col)
    db.session.commit()


def init_database() -> None:
    db.create_all()
    _ensure_schema_columns()
    _migrate_json_files_if_needed()


def _field(name: str) -> str:
    return request.form.get(name, "").strip()


def _form_to_risk(existing_id: str | None = None) -> dict[str, Any]:
    return {
        "id": existing_id or uuid.uuid4().hex,
        "risco": _field("risco"),
        "acoes": _field("acoes"),
        "prazo_implantacao": "",
        "prazo_reavaliacao": "",
        "indicador": _field("indicador"),
        "tipo_risco": _field("tipo_risco"),
        "descricao_agente": _field("descricao_agente"),
        "possiveis_lesoes": _field("possiveis_lesoes"),
        "fontes_circunstancias": _field("fontes_circunstancias") or "Durante o processo de trabalho.",
        "epis": _field("epis"),
        "epcs": _field("epcs"),
        "ltcat_meio_propagacao": _field("ltcat_meio_propagacao"),
        "ltcat_insalubridade": _field("ltcat_insalubridade") or "Não",
        "ltcat_grau_insalubridade": _field("ltcat_grau_insalubridade") or "Não aplicável",
        "ltcat_aposentadoria_especial": _field("ltcat_aposentadoria_especial") or "Não",
        "ltcat_enquadramento_tecnico": _field("ltcat_enquadramento_tecnico"),
        "ltcat_parecer_previdenciario": _field("ltcat_parecer_previdenciario"),
        "ltcat_periodicidade_jornada": _field("ltcat_periodicidade_jornada") or "Mensal (<= 4 horas < 10% jornada)",
        "grau_severidade": _field("grau_severidade"),
        "grau_possibilidade": _field("grau_possibilidade"),
        "grau_nivel_risco": _field("grau_nivel_risco"),
    }


def _form_to_sector(existing_id: str | None = None) -> dict[str, Any]:
    cargos = []
    cargo_names = request.form.getlist("cargo_nome[]")
    cargo_cbos = request.form.getlist("cargo_cbo[]")
    cargo_nfuncs = request.form.getlist("cargo_nfunc[]")
    cargo_descricoes = request.form.getlist("cargo_descricao[]")

    total = max(len(cargo_names), len(cargo_cbos), len(cargo_nfuncs), len(cargo_descricoes), 1)
    for index in range(total):
        cargo = {
            "id": uuid.uuid4().hex,
            "cargo": cargo_names[index].strip() if index < len(cargo_names) else "",
            "cbo": cargo_cbos[index].strip() if index < len(cargo_cbos) else "",
            "n_func": cargo_nfuncs[index].strip() if index < len(cargo_nfuncs) else "",
            "descricao": cargo_descricoes[index].strip() if index < len(cargo_descricoes) else "",
        }
        if any(cargo[key] for key in ["cargo", "cbo", "n_func", "descricao"]):
            cargos.append(cargo)

    return {
        "id": existing_id or uuid.uuid4().hex,
        "setor": _field("setor"),
        "grupo_id": _field("grupo_id"),
        "cargos": cargos,
    }


def _form_to_exam(existing_id: str | None = None) -> dict[str, Any]:
    return {
        "id": existing_id or uuid.uuid4().hex,
        "exame": _field("exame"),
        "periodicidade": _field("periodicidade"),
        "admissional": _field("admissional"),
        "periodico": _field("periodico"),
        "retorno": _field("retorno"),
        "mudanca": _field("mudanca"),
        "demissional": _field("demissional"),
    }


def _risk_from_dict(data: dict[str, Any], risk: Risk | None = None) -> Risk:
    risk = risk or Risk(id=data["id"])
    risk.risco = data["risco"]
    risk.acoes = data["acoes"]
    risk.indicador = data["indicador"]
    risk.tipo_risco = data["tipo_risco"]
    risk.descricao_agente = data.get("descricao_agente", "")
    risk.possiveis_lesoes = data["possiveis_lesoes"]
    risk.fontes_circunstancias = data.get("fontes_circunstancias") or "Durante o processo de trabalho."
    risk.epis = data["epis"]
    risk.epcs = data["epcs"]
    risk.ltcat_meio_propagacao = data.get("ltcat_meio_propagacao", "")
    risk.ltcat_insalubridade = data.get("ltcat_insalubridade") or "Não"
    risk.ltcat_grau_insalubridade = data.get("ltcat_grau_insalubridade") or "Não aplicável"
    risk.ltcat_aposentadoria_especial = data.get("ltcat_aposentadoria_especial") or "Não"
    risk.ltcat_enquadramento_tecnico = data.get("ltcat_enquadramento_tecnico", "")
    risk.ltcat_parecer_previdenciario = data.get("ltcat_parecer_previdenciario", "")
    risk.ltcat_periodicidade_jornada = data.get("ltcat_periodicidade_jornada") or "Mensal (<= 4 horas < 10% jornada)"
    risk.grau_severidade = data["grau_severidade"]
    risk.grau_possibilidade = data["grau_possibilidade"]
    risk.grau_nivel_risco = data["grau_nivel_risco"]
    risk.updated_at = datetime.utcnow()
    return risk


def _sector_from_dict(data: dict[str, Any], sector: Sector | None = None) -> Sector:
    sector = sector or Sector(id=data["id"])
    sector.setor = data["setor"]
    sector.group_id = data.get("grupo_id") or None
    sector.cargos = data.get("cargos", [])
    sector.updated_at = datetime.utcnow()
    return sector




def _form_to_company(existing_id: str | None = None) -> dict[str, Any]:
    return {
        "id": existing_id or uuid.uuid4().hex,
        "nome": _field("nome") or _field("empresa"),
        "empresa": _field("nome") or _field("empresa"),
        "cnpj": _field("cnpj"),
        "endereco": _field("endereco"),
        "bairro_cidade": _field("bairro_cidade"),
        "cep": _field("cep"),
        "cnae1": _field("cnae1") or _field("cnae"),
        "descricao1": _field("descricao1") or _field("descricao_atividade"),
        "grau1": _field("grau1") or _field("grau_risco"),
        "cnae2": _field("cnae2") or _field("cnae_secundario"),
        "descricao2": _field("descricao2") or _field("descricao_atividade_secundaria"),
        "grau2": _field("grau2") or _field("grau_risco_secundario"),
        "funcionarios": _field("funcionarios"),
        "data_atual": _field("data_atual"),
        "data_final": _field("data_final"),
        "email": _field("email"),
        "fone": _field("fone"),
        "data_avaliacao": _field("data_avaliacao"),
    }


def _company_from_dict(data: dict[str, Any], company: Company | None = None) -> Company:
    company = company or Company(id=data["id"])
    company.nome = data.get("nome") or data.get("empresa", "")
    company.cnpj = data.get("cnpj", "")
    company.endereco = data.get("endereco", "")
    company.bairro_cidade = data.get("bairro_cidade", "")
    company.cep = data.get("cep", "")
    company.cnae1 = data.get("cnae1", "")
    company.descricao1 = data.get("descricao1", "")
    company.grau1 = data.get("grau1", "")
    company.cnae2 = data.get("cnae2", "")
    company.descricao2 = data.get("descricao2", "")
    company.grau2 = data.get("grau2", "")
    company.funcionarios = data.get("funcionarios", "")
    company.data_atual = data.get("data_atual", "")
    company.data_final = data.get("data_final", "")
    company.email = data.get("email", "")
    company.fone = data.get("fone", "")
    company.data_avaliacao = data.get("data_avaliacao", "")
    company.updated_at = datetime.utcnow()
    return company


def _validate_company(company: dict[str, Any]) -> list[str]:
    errors: list[str] = []
    if not company.get("nome"):
        errors.append("Preencha o nome da empresa.")
    if not company.get("cnpj"):
        errors.append("Preencha o CNPJ da empresa.")
    return errors

def _exam_from_dict(data: dict[str, Any], exam: Exam | None = None) -> Exam:
    exam = exam or Exam(id=data["id"])
    exam.exame = data["exame"]
    exam.periodicidade = data.get("periodicidade", "")
    exam.admissional = data.get("admissional", "")
    exam.periodico = data.get("periodico", "")
    exam.retorno = data.get("retorno", "")
    exam.mudanca = data.get("mudanca", "")
    exam.demissional = data.get("demissional", "")
    exam.updated_at = datetime.utcnow()
    return exam


def _validate_exam(exam: dict[str, Any]) -> list[str]:
    errors: list[str] = []
    if not exam.get("exame"):
        errors.append("Preencha o nome do exame.")
    return errors


def _validate_risk(risk: dict[str, Any]) -> list[str]:
    errors: list[str] = []
    required = {
        "risco": "Perigo / Risco",
        "acoes": "Ações Preventiva / Corretiva",
        "indicador": "Indicador de Efetividade",
        "tipo_risco": "Tipo de Risco",
        "possiveis_lesoes": "Possíveis lesões ou agravos à saúde",
        "fontes_circunstancias": "Fontes ou circunstâncias",
        "epis": "EPI",
        "epcs": "EPC",
        "grau_severidade": "Grau de Severidade",
        "grau_possibilidade": "Grau de Possibilidade",
        "grau_nivel_risco": "Grau de Nível de Risco",
    }
    for key, label in required.items():
        if not risk.get(key):
            errors.append(f"Preencha o campo: {label}.")

    option_checks = {
        "tipo_risco": (FORM_OPTIONS["tipos_risco"], "Tipo de Risco"),
        "grau_severidade": (FORM_OPTIONS["severidades"], "Grau de Severidade"),
        "grau_possibilidade": (FORM_OPTIONS["possibilidades"], "Grau de Possibilidade"),
        "grau_nivel_risco": (FORM_OPTIONS["niveis_risco"], "Grau de Nível de Risco"),
    }
    for key, (valid_options, label) in option_checks.items():
        if risk.get(key) and risk[key] not in valid_options:
            errors.append(f"Selecione uma opção válida em: {label}.")
    return errors


def _validate_sector(sector: dict[str, Any]) -> list[str]:
    errors: list[str] = []
    if not sector.get("setor"):
        errors.append("Preencha o nome do setor.")
    if not sector.get("cargos"):
        errors.append("Adicione pelo menos um cargo ao setor.")
        return errors

    for index, cargo in enumerate(sector.get("cargos", []), start=1):
        if not cargo.get("cargo"):
            errors.append(f"Preencha o nome do cargo na linha {index}.")
        if not cargo.get("cbo"):
            errors.append(f"Preencha o CBO do cargo na linha {index}.")
        if not cargo.get("n_func"):
            errors.append(f"Preencha o número de funcionários na linha {index}.")
        if not cargo.get("descricao"):
            errors.append(f"Preencha a descrição da atividade na linha {index}.")
    return errors


def _sorted_risks() -> list[dict[str, Any]]:
    return [risk.to_dict() for risk in Risk.query.order_by(Risk.risco.asc()).all()]


def _sorted_sectors() -> list[dict[str, Any]]:
    return [sector.to_dict() for sector in Sector.query.order_by(Sector.setor.asc()).all()]


def _sorted_exams() -> list[dict[str, Any]]:
    return [exam.to_dict() for exam in Exam.query.order_by(Exam.exame.asc()).all()]


def _sorted_groups() -> list[dict[str, Any]]:
    return [group.to_dict() for group in SectorGroup.query.order_by(SectorGroup.nome.asc()).all()]




def _sorted_companies() -> list[dict[str, Any]]:
    return [company.to_dict() for company in Company.query.order_by(Company.nome.asc()).all()]


def _grouped_sectors() -> list[dict[str, Any]]:
    sectors = _sorted_sectors()
    grouped: dict[str, dict[str, Any]] = {}
    for sector in sectors:
        gid = sector.get("grupo_id") or "__sem_grupo__"
        if gid not in grouped:
            grouped[gid] = {
                "id": gid,
                "nome": sector.get("grupo_nome") or "Sem grupo",
                "sectors": [],
            }
        grouped[gid]["sectors"].append(sector)
    return sorted(grouped.values(), key=lambda item: (item["nome"] == "Sem grupo", item["nome"].lower()))


def _generation_form_extras() -> dict[str, str]:
    """Dados específicos da finalização do laudo, preenchidos na tela Gerar laudos."""
    ajuste = request.form.get("ajuste_psicossocial") == "1"
    return {
        "data_criacao_laudo": _field("data_criacao_laudo"),
        "datacriacaolaudo": _field("data_criacao_laudo"),
        "ajuste_psicossocial": "1" if ajuste else "",
        "data_da_revisao": _field("data_da_revisao") if ajuste else "",
        "data_revisao_ajuste": _field("data_da_revisao") if ajuste else "",
    }


def _company_payload_from_form() -> dict[str, str]:
    company_id = _field("company_id")
    if company_id:
        company = db.session.get(Company, company_id)
        if company:
            data = company.to_dict()
            # Permite sobrescrever a vigência se um campo de formulário existir no futuro.
            data["data_atual"] = _field("data_atual") or data.get("data_atual", "")
            data["data_final"] = _field("data_final") or data.get("data_final", "")
            data.update(_generation_form_extras())
            return data
    data = _form_to_company()
    data["empresa"] = data.get("nome", "")
    data["cnae"] = data.get("cnae1", "")
    data["descricao_atividade"] = data.get("descricao1", "")
    data["grau_risco"] = data.get("grau1", "")
    data["cnae_secundario"] = data.get("cnae2", "")
    data["descricao_atividade_secundaria"] = data.get("descricao2", "")
    data["grau_risco_secundario"] = data.get("grau2", "")
    data.update(_generation_form_extras())
    return data

def _selected_risks() -> list[dict[str, Any]]:
    selected_ids = request.form.getlist("risk_ids")
    if not selected_ids:
        return []
    risks = Risk.query.filter(Risk.id.in_(selected_ids)).all()
    order = {risk_id: index for index, risk_id in enumerate(selected_ids)}
    risks.sort(key=lambda item: order.get(item.id, 999999))
    return [risk.to_dict() for risk in risks]


def _selected_sectors() -> list[dict[str, Any]]:
    selected_ids = request.form.getlist("sector_ids") or request.form.getlist("pgr_sector_ids")
    if not selected_ids:
        return []
    sectors = Sector.query.filter(Sector.id.in_(selected_ids)).all()
    order = {sector_id: index for index, sector_id in enumerate(selected_ids)}
    sectors.sort(key=lambda item: order.get(item.id, 999999))
    return [sector.to_dict() for sector in sectors]


def _selected_sector_risk_groups() -> tuple[list[dict[str, Any]], list[str]]:
    selected_sector_ids = request.form.getlist("pgr_sector_ids")
    risks = {risk.id: risk.to_dict() for risk in Risk.query.all()}
    sectors = {sector.id: sector.to_dict() for sector in Sector.query.all()}
    exams = {exam.id: exam.to_dict() for exam in Exam.query.all()}

    groups: list[dict[str, Any]] = []
    errors: list[str] = []

    if not selected_sector_ids:
        errors.append("Selecione pelo menos um setor para gerar o PGR/PCMSO.")
        return groups, errors

    for sector_id in selected_sector_ids:
        sector = sectors.get(sector_id)
        if not sector:
            continue
        risk_ids = request.form.getlist(f"sector_risk_ids_{sector_id}")
        exam_ids = request.form.getlist(f"sector_exam_ids_{sector_id}")
        selected_risks = [risks[risk_id] for risk_id in risk_ids if risk_id in risks]
        selected_exams = [exams[exam_id] for exam_id in exam_ids if exam_id in exams]
        if not selected_risks:
            errors.append(f"Selecione pelo menos um risco para o setor: {sector.get('setor', '')}.")
        else:
            groups.append({"sector": sector, "risks": selected_risks, "exams": selected_exams})

    return groups, errors


@app.route("/")
def index():
    return redirect(url_for("cadastro"))


@app.route("/cadastro")
def cadastro():
    return render_template("cadastro.html", risks=_sorted_risks(), options=FORM_OPTIONS)


@app.route("/setores")
def setores():
    return render_template("setores.html", sectors=_sorted_sectors(), groups=_sorted_groups(), grouped_sectors=_grouped_sectors())


@app.post("/grupo/novo")
def create_group():
    nome = _field("nome_grupo")
    if not nome:
        flash("Informe o nome do grupo.", "error")
        return redirect(url_for("setores"))
    existing = SectorGroup.query.filter(db.func.lower(SectorGroup.nome) == nome.lower()).first()
    if existing:
        flash("Esse grupo já existe.", "error")
        return redirect(url_for("setores"))
    db.session.add(SectorGroup(nome=nome))
    db.session.commit()
    flash("Grupo cadastrado com sucesso.", "success")
    return redirect(url_for("setores"))


@app.post("/grupo/<group_id>/excluir")
def delete_group(group_id: str):
    group = db.session.get(SectorGroup, group_id)
    if group:
        Sector.query.filter_by(group_id=group_id).update({"group_id": None})
        db.session.delete(group)
        db.session.commit()
        flash("Grupo excluído. Os setores foram mantidos sem grupo.", "success")
    return redirect(url_for("setores"))




@app.route("/empresas")
def empresas():
    return render_template("empresas.html", companies=_sorted_companies())


@app.route("/empresa/nova", methods=["GET", "POST"])
def create_company():
    if request.method == "POST":
        company = _form_to_company()
        errors = _validate_company(company)
        if errors:
            for error in errors:
                flash(error, "error")
            return render_template("empresa_form.html", company=company, title="Nova empresa")
        db.session.add(_company_from_dict(company))
        db.session.commit()
        flash("Empresa cadastrada com sucesso.", "success")
        return redirect(url_for("empresas"))
    today = datetime.now().strftime("%m/%Y")
    next_year = datetime.now().replace(year=datetime.now().year + 1).strftime("%m/%Y")
    return render_template("empresa_form.html", company={"data_atual": today, "data_final": next_year}, title="Nova empresa")


@app.route("/empresa/<company_id>/editar", methods=["GET", "POST"])
def edit_company(company_id: str):
    company_model = db.session.get(Company, company_id)
    if not company_model:
        flash("Empresa não encontrada.", "error")
        return redirect(url_for("empresas"))
    if request.method == "POST":
        updated = _form_to_company(existing_id=company_id)
        errors = _validate_company(updated)
        if errors:
            for error in errors:
                flash(error, "error")
            return render_template("empresa_form.html", company={**company_model.to_dict(), **updated}, title="Editar empresa")
        _company_from_dict(updated, company_model)
        db.session.commit()
        flash("Empresa atualizada com sucesso.", "success")
        return redirect(url_for("empresas"))
    return render_template("empresa_form.html", company=company_model.to_dict(), title="Editar empresa")


@app.post("/empresa/<company_id>/excluir")
def delete_company(company_id: str):
    company_model = db.session.get(Company, company_id)
    if company_model:
        db.session.delete(company_model)
        db.session.commit()
        flash("Empresa excluída.", "success")
    return redirect(url_for("empresas"))

@app.route("/exames")
def exames():
    return render_template("exames.html", exams=_sorted_exams())


@app.route("/exame/novo", methods=["GET", "POST"])
def create_exam():
    if request.method == "POST":
        new_exam = _form_to_exam()
        errors = _validate_exam(new_exam)
        if errors:
            for error in errors:
                flash(error, "error")
            return render_template("exame_form.html", exam=new_exam, title="Novo exame")
        db.session.add(_exam_from_dict(new_exam))
        db.session.commit()
        flash("Exame cadastrado com sucesso.", "success")
        return redirect(url_for("exames"))
    return render_template("exame_form.html", exam={}, title="Novo exame")


@app.route("/exame/<exam_id>/editar", methods=["GET", "POST"])
def edit_exam(exam_id: str):
    exam_model = db.session.get(Exam, exam_id)
    if not exam_model:
        flash("Exame não encontrado.", "error")
        return redirect(url_for("exames"))
    if request.method == "POST":
        updated = _form_to_exam(existing_id=exam_id)
        errors = _validate_exam(updated)
        if errors:
            for error in errors:
                flash(error, "error")
            return render_template("exame_form.html", exam={**exam_model.to_dict(), **updated}, title="Editar exame")
        _exam_from_dict(updated, exam_model)
        db.session.commit()
        flash("Exame atualizado com sucesso.", "success")
        return redirect(url_for("exames"))
    return render_template("exame_form.html", exam=exam_model.to_dict(), title="Editar exame")


@app.post("/exame/<exam_id>/excluir")
def delete_exam(exam_id: str):
    exam_model = db.session.get(Exam, exam_id)
    if exam_model:
        db.session.delete(exam_model)
        db.session.commit()
        flash("Exame excluído.", "success")
    return redirect(url_for("exames"))


def _gerar_form_state_from_request() -> dict[str, Any]:
    sector_ids = request.form.getlist("pgr_sector_ids")
    risks_by_sector = {sector_id: request.form.getlist(f"sector_risk_ids_{sector_id}") for sector_id in sector_ids}
    exams_by_sector = {sector_id: request.form.getlist(f"sector_exam_ids_{sector_id}") for sector_id in sector_ids}
    return {
        "company_id": _field("company_id"),
        "data_criacao_laudo": _field("data_criacao_laudo"),
        "ajuste_psicossocial": "1" if request.form.get("ajuste_psicossocial") == "1" else "",
        "data_da_revisao": _field("data_da_revisao"),
        "selected_sector_ids": sector_ids,
        "selected_risk_ids_by_sector": risks_by_sector,
        "selected_exam_ids_by_sector": exams_by_sector,
    }


def _gerar_context(form_state: dict[str, Any] | None = None) -> dict[str, Any]:
    today = datetime.now().strftime("%m/%Y")
    next_year = datetime.now().replace(year=datetime.now().year + 1).strftime("%m/%Y")
    return {
        "risks": _sorted_risks(),
        "sectors": _sorted_sectors(),
        "exams": _sorted_exams(),
        "groups": _sorted_groups(),
        "grouped_sectors": _grouped_sectors(),
        "companies": _sorted_companies(),
        "options": FORM_OPTIONS,
        "today": today,
        "next_year": next_year,
        "form_state": form_state or {},
    }


def _render_gerar_with_current_form():
    return render_template("gerar.html", **_gerar_context(_gerar_form_state_from_request()))


def _validate_complete_report_fields(company: dict[str, str], label: str) -> list[str]:
    errors: list[str] = []
    empresa = company.get("empresa") or company.get("nome", "")
    if not empresa:
        errors.append(f"Selecione ou cadastre uma empresa para gerar o {label} completo.")
    if not company.get("cnpj"):
        errors.append(f"Cadastre o CNPJ da empresa para gerar o {label} completo.")
    if not company.get("data_criacao_laudo"):
        errors.append("Preencha a Data de criação do laudo para finalizar o documento.")
    if company.get("ajuste_psicossocial") == "1" and not company.get("data_da_revisao"):
        errors.append("Preencha a Data da revisão psicossocial quando marcar que o laudo é apenas um ajuste.")
    return errors


@app.route("/gerar")
def gerar():
    return render_template("gerar.html", **_gerar_context())


@app.post("/risco/novo")
def create_risk():
    new_risk = _form_to_risk()
    errors = _validate_risk(new_risk)
    if errors:
        for error in errors:
            flash(error, "error")
        return redirect(url_for("cadastro"))
    db.session.add(_risk_from_dict(new_risk))
    db.session.commit()
    flash("Risco cadastrado com sucesso.", "success")
    return redirect(url_for("cadastro"))


@app.route("/risco/<risk_id>/editar", methods=["GET", "POST"])
def edit_risk(risk_id: str):
    risk_model = db.session.get(Risk, risk_id)
    if not risk_model:
        flash("Risco não encontrado.", "error")
        return redirect(url_for("cadastro"))

    if request.method == "POST":
        updated = _form_to_risk(existing_id=risk_id)
        errors = _validate_risk(updated)
        if errors:
            for error in errors:
                flash(error, "error")
            return render_template("edit.html", risk={**risk_model.to_dict(), **updated}, options=FORM_OPTIONS)
        _risk_from_dict(updated, risk_model)
        db.session.commit()
        flash("Risco atualizado com sucesso.", "success")
        return redirect(url_for("cadastro"))

    return render_template("edit.html", risk=risk_model.to_dict(), options=FORM_OPTIONS)


@app.post("/risco/<risk_id>/excluir")
def delete_risk(risk_id: str):
    risk_model = db.session.get(Risk, risk_id)
    if risk_model:
        db.session.delete(risk_model)
        db.session.commit()
        flash("Risco excluído.", "success")
    return redirect(url_for("cadastro"))


@app.post("/setor/novo")
def create_sector():
    new_sector = _form_to_sector()
    errors = _validate_sector(new_sector)
    if errors:
        for error in errors:
            flash(error, "error")
        return redirect(url_for("setores"))
    db.session.add(_sector_from_dict(new_sector))
    db.session.commit()
    flash("Setor e cargos cadastrados com sucesso.", "success")
    return redirect(url_for("setores"))


@app.route("/setor/<sector_id>/editar", methods=["GET", "POST"])
def edit_sector(sector_id: str):
    sector_model = db.session.get(Sector, sector_id)
    if not sector_model:
        flash("Setor não encontrado.", "error")
        return redirect(url_for("setores"))

    if request.method == "POST":
        updated = _form_to_sector(existing_id=sector_id)
        errors = _validate_sector(updated)
        if errors:
            for error in errors:
                flash(error, "error")
            return render_template("edit_setor.html", sector={**sector_model.to_dict(), **updated}, groups=_sorted_groups())
        _sector_from_dict(updated, sector_model)
        db.session.commit()
        flash("Setor atualizado com sucesso.", "success")
        return redirect(url_for("setores"))

    return render_template("edit_setor.html", sector=sector_model.to_dict(), groups=_sorted_groups())


@app.post("/setor/<sector_id>/excluir")
def delete_sector(sector_id: str):
    sector_model = db.session.get(Sector, sector_id)
    if sector_model:
        db.session.delete(sector_model)
        db.session.commit()
        flash("Setor excluído.", "success")
    return redirect(url_for("setores"))


@app.post("/setores/apagar-todos")
def delete_all_sectors():
    count = Sector.query.count()
    Sector.query.delete()
    db.session.commit()
    flash(f"{count} setor(es) e seus cargos foram apagados. Os riscos cadastrados foram mantidos.", "success")
    return redirect(url_for("setores"))


def _selected_sector_ltcat_groups() -> tuple[list[dict[str, Any]], list[str]]:
    selected_sector_ids = request.form.getlist("pgr_sector_ids")
    risks = {risk.id: risk.to_dict() for risk in Risk.query.all()}
    sectors = {sector.id: sector.to_dict() for sector in Sector.query.all()}

    groups: list[dict[str, Any]] = []
    errors: list[str] = []

    if not selected_sector_ids:
        errors.append("Selecione pelo menos um setor para gerar o LTCAT.")
        return groups, errors

    for sector_id in selected_sector_ids:
        sector = sectors.get(sector_id)
        if not sector:
            continue
        risk_ids = request.form.getlist(f"sector_risk_ids_{sector_id}")
        selected_risks = [risks[risk_id] for risk_id in risk_ids if risk_id in risks]
        groups.append({"sector": sector, "risks": selected_risks, "exams": []})

    return groups, errors


def _company_extra_from_form() -> dict[str, str]:
    return _company_payload_from_form()


def _send_generated_docx(generator, selected: list[dict[str, Any]], stem: str, download_name: str, *args):
    if not selected:
        flash("Selecione pelo menos um item para gerar o arquivo Word.", "error")
        return _render_gerar_with_current_form()

    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    output_path = OUTPUT_DIR / f"{stem}_{stamp}.docx"
    generator(selected, output_path, *args)
    return send_file(
        output_path,
        as_attachment=True,
        download_name=download_name,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )


@app.post("/gerar-pgr-completo")
def generate_complete_pgr():
    groups, errors = _selected_sector_risk_groups()
    company = _company_payload_from_form()
    errors.extend(_validate_complete_report_fields(company, "PGR"))
    if errors:
        for error in errors:
            flash(error, "error")
        return _render_gerar_with_current_form()
    try:
        empresa = company.get("empresa") or company.get("nome", "")
        cnpj = company.get("cnpj", "")
        data_atual = company.get("data_atual", "")
        data_final = company.get("data_final", "")
        return _send_generated_docx(
            generate_complete_pgr_docx,
            groups,
            "pgr_completo",
            "PGR_COMPLETO.docx",
            empresa,
            cnpj,
            data_atual,
            data_final,
            company,
        )
    except Exception as exc:
        flash(f"Erro ao gerar PGR completo: {exc}", "error")
        return _render_gerar_with_current_form()


@app.post("/gerar-plano-acao")
def generate_action_plan():
    try:
        if request.form.getlist("pgr_sector_ids"):
            groups, errors = _selected_sector_risk_groups()
            if errors:
                for error in errors:
                    flash(error, "error")
                return redirect(url_for("gerar"))
            selected = groups
        else:
            selected = _selected_risks()
        return _send_generated_docx(
            generate_action_plan_docx,
            selected,
            "plano_de_acao_riscos",
            "PLANO_DE_ACAO_RISCOS.docx",
            _field("data_atual"),
            _field("data_final"),
        )
    except Exception as exc:
        flash(f"Erro ao gerar Plano de Ação: {exc}", "error")
        return redirect(url_for("gerar"))


@app.post("/gerar-risco-pgr")
def generate_risco_pgr():
    groups, errors = _selected_sector_risk_groups()
    if errors:
        for error in errors:
            flash(error, "error")
        return redirect(url_for("gerar"))
    try:
        return _send_generated_docx(generate_pgr_docx, groups, "risco_pgr", "RISCO_PGR.docx")
    except Exception as exc:
        flash(f"Erro ao gerar Risco PGR: {exc}", "error")
        return redirect(url_for("gerar"))


@app.post("/gerar-riscos-pcmso")
def generate_riscos_pcmso():
    groups, errors = _selected_sector_risk_groups()
    if errors:
        for error in errors:
            flash(error, "error")
        return redirect(url_for("gerar"))
    try:
        return _send_generated_docx(generate_riscos_pcmso_docx, groups, "riscos_pcmso", "RISCOS_PCMSO.docx")
    except Exception as exc:
        flash(f"Erro ao gerar riscos/exames do PCMSO: {exc}", "error")
        return redirect(url_for("gerar"))


@app.post("/gerar-ltcat-completo")
def generate_complete_ltcat():
    groups, errors = _selected_sector_ltcat_groups()
    company = _company_payload_from_form()
    errors.extend(_validate_complete_report_fields(company, "LTCAT"))
    if errors:
        for error in errors:
            flash(error, "error")
        return _render_gerar_with_current_form()
    try:
        empresa = company.get("empresa") or company.get("nome", "")
        cnpj = company.get("cnpj", "")
        data_atual = company.get("data_atual", "")
        data_final = company.get("data_final", "")
        return _send_generated_docx(
            generate_complete_ltcat_docx,
            groups,
            "ltcat_completo",
            "LTCAT_COMPLETO.docx",
            empresa,
            cnpj,
            data_atual,
            data_final,
            company,
        )
    except Exception as exc:
        flash(f"Erro ao gerar LTCAT completo: {exc}", "error")
        return _render_gerar_with_current_form()


@app.post("/gerar-pcmso-completo")
def generate_complete_pcmso():
    groups, errors = _selected_sector_risk_groups()
    company = _company_payload_from_form()
    errors.extend(_validate_complete_report_fields(company, "PCMSO"))
    if errors:
        for error in errors:
            flash(error, "error")
        return _render_gerar_with_current_form()
    try:
        empresa = company.get("empresa") or company.get("nome", "")
        cnpj = company.get("cnpj", "")
        data_atual = company.get("data_atual", "")
        data_final = company.get("data_final", "")
        return _send_generated_docx(
            generate_complete_pcmso_docx,
            groups,
            "pcmso_completo",
            "PCMSO_COMPLETO.docx",
            empresa,
            cnpj,
            data_atual,
            data_final,
            company,
        )
    except Exception as exc:
        flash(f"Erro ao gerar PCMSO completo: {exc}", "error")
        return _render_gerar_with_current_form()


@app.post("/gerar-pcmso")
def generate_pcmso():
    return generate_riscos_pcmso()


@app.post("/gerar-relacao-funcao-atividade")
def generate_relacao_funcao_atividade():
    try:
        return _send_generated_docx(
            generate_relacao_funcao_atividade_docx,
            _selected_sectors(),
            "relacao_funcao_atividade",
            "RELACAO_FUNCAO_X_ATIVIDADE.docx",
            _field("data_atual"),
            _field("data_final"),
        )
    except Exception as exc:
        flash(f"Erro ao gerar Relação Função x Atividade: {exc}", "error")
        return redirect(url_for("gerar"))


@app.post("/gerar-descritivo-setor")
def generate_descritivo_setor():
    try:
        return _send_generated_docx(generate_descritivo_setor_docx, _selected_sectors(), "descritivo_setor", "DESCRITIVO_SETOR.docx")
    except Exception as exc:
        flash(f"Erro ao gerar Descritivo Setor: {exc}", "error")
        return redirect(url_for("gerar"))


@app.post("/gerar-word")
def generate_word_legacy():
    return generate_action_plan()


with app.app_context():
    init_database()


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)), debug=True)
