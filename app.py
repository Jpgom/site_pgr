from __future__ import annotations

import json
import os
import re
import shutil
import tempfile
import uuid
import zipfile
from datetime import datetime
from pathlib import Path
from typing import Any

from flask import Flask, flash, jsonify, redirect, render_template, request, send_file, url_for
from werkzeug.utils import secure_filename
from werkzeug.exceptions import RequestEntityTooLarge
from flask_sqlalchemy import SQLAlchemy
from sqlalchemy import Text, inspect, text

try:
    from openpyxl import load_workbook
except Exception:  # pragma: no cover - usado apenas se a dependência não estiver instalada
    load_workbook = None

from word_generator import (
    NIVEL_RISCO_COLORS,
    POSSIBILIDADE_COLORS,
    SEVERIDADE_COLORS,
    TIPO_RISCO_COLORS,
    generate_action_plan_docx,
    generate_complete_pgr_docx,
    generate_complete_pcmso_docx,
    generate_complete_ltcat_docx,
    generate_aet_docx,
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
RISK_IMPORT_TEMPLATE = BASE_DIR / "modelos" / "modelo_importacao_riscos.xlsx"
SECTOR_IMPORT_TEMPLATE = BASE_DIR / "modelos" / "modelo_importacao_setores.xlsx"
LINK_PGR_AET_TEMPLATE = BASE_DIR / "modelos" / "link_pgr_para_aet.docx"
LINK_AET_PSICOSSOCIAL_TEMPLATE = BASE_DIR / "modelos" / "link_aet_para_psicossocial.docx"
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
# Evita uploads gigantes travarem o serviço no Render. Ajuste por variável de ambiente se precisar.
app.config["MAX_CONTENT_LENGTH"] = int(os.environ.get("MAX_UPLOAD_MB", "120")) * 1024 * 1024
# A tela Gerar Laudos envia muitos campos quando há muitos setores, riscos, exames e dados de AET.
# Sem estes limites maiores, o Werkzeug pode retornar: Request Entity Too Large.
app.config["MAX_FORM_MEMORY_SIZE"] = int(os.environ.get("MAX_FORM_MEMORY_MB", "80")) * 1024 * 1024
app.config["MAX_FORM_PARTS"] = int(os.environ.get("MAX_FORM_PARTS", "50000"))

# Configuração da conversão visual do PDF psicossocial para Word.
# Usamos JPEG em DPI moderado para preservar o visual sem travar por arquivos enormes.
PSICOSSOCIAL_RENDER_DPI = max(90, min(180, int(os.environ.get("PSICOSSOCIAL_RENDER_DPI", "135"))))
PSICOSSOCIAL_JPEG_QUALITY = max(60, min(92, int(os.environ.get("PSICOSSOCIAL_JPEG_QUALITY", "82"))))
PSICOSSOCIAL_MAX_PAGES = int(os.environ.get("PSICOSSOCIAL_MAX_PAGES", "80"))

db = SQLAlchemy(app)

@app.errorhandler(RequestEntityTooLarge)
def handle_request_entity_too_large(error):
    flash(
        "O formulário ou arquivo enviado ficou maior que o limite configurado. "
        "Tente gerar novamente. Se estiver juntando arquivos, reduza o PDF ou aumente MAX_UPLOAD_MB/MAX_FORM_MEMORY_MB no Render.",
        "error",
    )
    return redirect(request.referrer or url_for("generate"))


risk_group_items = db.Table(
    "risk_group_items",
    db.Column("risk_group_id", db.String(32), db.ForeignKey("risk_groups.id"), primary_key=True),
    db.Column("risk_id", db.String(32), db.ForeignKey("risks.id"), primary_key=True),
)


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
    ltcat_insalubridade = db.Column(Text, default="Não")
    ltcat_grau_insalubridade = db.Column(Text, default="Não aplicável")
    ltcat_aposentadoria_especial = db.Column(Text, default="Não")
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


class RiskGroup(db.Model):
    __tablename__ = "risk_groups"

    id = db.Column(db.String(32), primary_key=True, default=lambda: uuid.uuid4().hex)
    nome = db.Column(db.String(255), nullable=False, unique=True)
    descricao = db.Column(Text, default="")
    risks = db.relationship("Risk", secondary=risk_group_items, lazy="select")
    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow, nullable=False)

    def to_dict(self) -> dict[str, Any]:
        ordered_risks = sorted(self.risks or [], key=lambda item: (item.risco or "").lower())
        return {
            "id": self.id,
            "nome": self.nome,
            "descricao": self.descricao or "",
            "risk_ids": [risk.id for risk in ordered_risks],
            "risk_names": [risk.risco for risk in ordered_risks],
            "risk_count": len(ordered_risks),
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
    nome = db.Column(Text, nullable=False)
    cnpj = db.Column(Text, nullable=False)
    endereco = db.Column(Text, default="")
    bairro_cidade = db.Column(Text, default="")
    cep = db.Column(Text, default="")
    cnae1 = db.Column(Text, default="")
    descricao1 = db.Column(Text, default="")
    grau1 = db.Column(Text, default="")
    cnae2 = db.Column(Text, default="")
    descricao2 = db.Column(Text, default="")
    grau2 = db.Column(Text, default="")
    funcionarios = db.Column(Text, default="")
    data_atual = db.Column(Text, default="")
    data_final = db.Column(Text, default="")
    email = db.Column(Text, default="")
    fone = db.Column(Text, default="")
    data_avaliacao = db.Column(Text, default="")
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


class ReportProfile(db.Model):
    """Configuração salva da tela Gerar laudos para uma empresa.

    Guarda setores, riscos, grupos de riscos, exames e datas usadas na finalização,
    permitindo regenerar os laudos depois sem refazer toda a seleção.
    """

    __tablename__ = "report_profiles"

    id = db.Column(db.String(32), primary_key=True, default=lambda: uuid.uuid4().hex)
    company_id = db.Column(db.String(32), nullable=False, index=True)
    nome = db.Column(db.String(255), nullable=False)
    data_criacao_laudo = db.Column(Text, default="")
    ajuste_psicossocial = db.Column(db.String(1), default="")
    data_da_revisao = db.Column(Text, default="")
    state = db.Column(db.JSON, nullable=False, default=dict)
    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow, nullable=False)

    def to_dict(self) -> dict[str, Any]:
        company = db.session.get(Company, self.company_id) if self.company_id else None
        return {
            "id": self.id,
            "company_id": self.company_id,
            "company_nome": company.nome if company else "",
            "nome": self.nome,
            "data_criacao_laudo": self.data_criacao_laudo or "",
            "ajuste_psicossocial": self.ajuste_psicossocial or "",
            "data_da_revisao": self.data_da_revisao or "",
            "state": self.state or {},
            "created_at": self.created_at.strftime("%Y-%m-%d %H:%M:%S") if self.created_at else "",
            "updated_at": self.updated_at.strftime("%Y-%m-%d %H:%M:%S") if self.updated_at else "",
        }


FORM_OPTIONS = {
    "tipos_risco": list(TIPO_RISCO_COLORS.keys()),
    "severidades": list(SEVERIDADE_COLORS.keys()),
    "possibilidades": list(POSSIBILIDADE_COLORS.keys()),
    "niveis_risco": list(NIVEL_RISCO_COLORS.keys()),
}

TIPOS_RISCO_LTCAT = {"FÍSICO", "QUÍMICO", "BIOLÓGICO"}

# Presets técnicos de AET por CNAE/atividade.
# Usados apenas para pré-marcar o formulário; o usuário pode alterar tudo antes de gerar.
AET_CNAE_PRESETS = [
    {
        "keywords": ["47.81", "vestuário", "vestuario", "comércio varejista", "loja", "vendas"],
        "label": "Comércio varejista / vestuário",
        "general": {
            "tipo_documento": "AET completa com formulário ergonômico",
            "motivo_analise": "Atendimento à NR-17 e integração com o PGR",
            "condicao_ergonomica_geral": "Adequada com recomendações",
        },
        "sector": {
            "postura": ["Em pé", "Deslocamento frequente", "Alternado"],
            "tipo_atividade": "Atendimento, organização de mercadorias e apoio operacional",
            "exigencia_fisica": "Moderada",
            "exigencia_cognitiva": "Moderada",
            "ritmo_trabalho": "Por demanda e fluxo de clientes",
            "pausas": "Intervalos legais e pausas breves conforme organização da loja",
            "mobiliario": "Parcialmente adequado",
            "ambiente": "A avaliar conforto térmico, iluminação e circulação",
            "organizacao": "Atendimento ao público, organização de prioridades e comunicação com liderança",
            "equipamentos": "Balcão, prateleiras, araras, computador/sistema e materiais de apoio",
            "fatores": ["Atendimento ao público", "Deslocamentos no setor", "Postura em pé", "Organização de demandas"],
            "medidas": ["Alternância postural", "Pausas breves", "Organização do posto", "Orientação NR-17"],
            "prioridade": "Média",
            "prazo": "30 dias",
            "responsavel": "Empresa / Administração",
        },
    },
    {
        "keywords": ["81.21", "81.22", "81.29", "limpeza", "serviços combinados", "servicos combinados", "serviços gerais", "servicos gerais"],
        "label": "Limpeza / serviços gerais",
        "general": {
            "tipo_documento": "AET completa com formulário ergonômico",
            "motivo_analise": "Atendimento à NR-17 e avaliação das exigências físicas das atividades",
            "condicao_ergonomica_geral": "Adequada com recomendações",
        },
        "sector": {
            "postura": ["Em pé", "Deslocamento frequente", "Inclinação/flexão de tronco", "Agachamento eventual"],
            "tipo_atividade": "Limpeza, conservação, deslocamento e manuseio de materiais",
            "exigencia_fisica": "Moderada",
            "exigencia_cognitiva": "Baixa",
            "ritmo_trabalho": "Por rotina e demandas do ambiente",
            "pausas": "Pausas e alternância conforme intensidade da atividade",
            "mobiliario": "Não aplicável diretamente; avaliar equipamentos e ferramentas manuais",
            "ambiente": "A avaliar piso, circulação, ventilação e disponibilidade de local para descanso",
            "organizacao": "Rotina operacional com definição clara de prioridades",
            "equipamentos": "Vassouras, rodos, baldes, panos, produtos saneantes e carrinhos quando disponíveis",
            "fatores": ["Esforço físico", "Posturas incômodas", "Deslocamentos frequentes", "Manuseio de materiais"],
            "medidas": ["Alternância de tarefas", "Pausas", "Ferramentas adequadas", "Orientação NR-17"],
            "prioridade": "Média",
            "prazo": "30 dias",
            "responsavel": "Empresa / Administração",
        },
    },
    {
        "keywords": ["administrativo", "escritório", "escritorio", "82.", "69.", "70."],
        "label": "Administrativo / escritório",
        "general": {
            "tipo_documento": "AET documental com análise por setor",
            "motivo_analise": "Atendimento à NR-17 e análise do posto administrativo",
            "condicao_ergonomica_geral": "Adequada com recomendações",
        },
        "sector": {
            "postura": ["Sentado", "Alternado"],
            "tipo_atividade": "Atividades administrativas, digitação, atendimento e organização documental",
            "exigencia_fisica": "Baixa",
            "exigencia_cognitiva": "Moderada",
            "ritmo_trabalho": "Por demanda administrativa",
            "pausas": "Pausas breves e alternância postural recomendadas",
            "mobiliario": "A avaliar cadeira, mesa, monitor e acessórios",
            "ambiente": "A avaliar iluminação, reflexos, ventilação e conforto acústico",
            "organizacao": "Rotina administrativa com atenção a prazos, prioridades e comunicação",
            "equipamentos": "Computador, monitor, teclado, mouse, telefone e documentos",
            "fatores": ["Trabalho sentado", "Digitação", "Atenção contínua", "Organização de prioridades"],
            "medidas": ["Ajuste de mobiliário", "Alternância postural", "Pausas visuais", "Orientação NR-17"],
            "prioridade": "Baixa",
            "prazo": "60 dias",
            "responsavel": "Empresa / Administração",
        },
    },
    {
        "keywords": ["portaria", "porteiro", "vigilância", "vigilancia", "condomínio", "condominio", "81.12"],
        "label": "Portaria / condomínio",
        "general": {
            "tipo_documento": "AET completa com formulário ergonômico",
            "motivo_analise": "Atendimento à NR-17 e análise de atividade com atenção contínua",
            "condicao_ergonomica_geral": "Adequada com recomendações",
        },
        "sector": {
            "postura": ["Sentado", "Em pé", "Alternado"],
            "tipo_atividade": "Controle de acesso, atendimento, monitoramento e comunicação",
            "exigencia_fisica": "Baixa",
            "exigencia_cognitiva": "Elevada",
            "ritmo_trabalho": "Contínuo, com atenção permanente e demandas variáveis",
            "pausas": "Pausas conforme escala, mantendo cobertura operacional",
            "mobiliario": "A avaliar cadeira, bancada, campo visual e acesso aos controles",
            "ambiente": "A avaliar ventilação, iluminação, conforto térmico e ruído",
            "organizacao": "Atenção constante, comunicação e controle de prioridades",
            "equipamentos": "Portão, interfone, rádio/telefone, computador, câmeras e controles",
            "fatores": ["Atenção contínua", "Atendimento ao público", "Comunicação", "Postura alternada"],
            "medidas": ["Ajuste do posto", "Pausas", "Definição de procedimentos", "Orientação NR-17"],
            "prioridade": "Média",
            "prazo": "30 dias",
            "responsavel": "Empresa / Administração",
        },
    },
]

# Motor simples de regras técnicas para sugestão de exames por risco.
# O sistema não trava a geração; ele apenas marca/sugere exames já cadastrados
# quando o nome do exame combina com palavras-chave.
EXAM_RULES = [
    {"keywords": ["ruído", "ruido", "audição", "audiometria"], "exams": ["audiometria"]},
    {"keywords": ["poeira", "poeiras", "sílica", "silica", "fumos", "névoa", "nevoa", "respirável", "respiravel", "gases", "vapores"], "exams": ["espirometria", "raio x", "radiografia", "exame clínico", "exame clinico"]},
    {"keywords": ["químico", "quimico", "produto químico", "solvente", "hidrocarboneto", "gasolina", "diesel", "óleo", "oleo", "graxa"], "exams": ["exame clínico", "exame clinico", "hemograma", "tgo", "tgp"]},
    {"keywords": ["biológico", "biologico", "sangue", "vírus", "virus", "bactéria", "bacteria", "fungo", "parasita", "resíduo", "residuo"], "exams": ["exame clínico", "exame clinico", "hemograma", "vacinação", "vacina"]},
    {"keywords": ["altura", "queda", "nível", "nivel", "telhado", "escada"], "exams": ["exame clínico", "exame clinico", "acuidade visual", "eletrocardiograma", "ecg"]},
    {"keywords": ["eletricidade", "elétrico", "eletrico", "choque"], "exams": ["exame clínico", "exame clinico", "eletrocardiograma", "ecg"]},
    {"keywords": ["calor", "temperatura", "frio", "câmara frigorífica", "camara frigorifica"], "exams": ["exame clínico", "exame clinico"]},
    {"keywords": ["ergonômico", "ergonomico", "postura", "repetitivo", "carga", "esforço", "esforco"], "exams": ["exame clínico", "exame clinico", "anamnese ocupacional"]},
    {"keywords": ["psicossocial", "assédio", "assedio", "estresse", "sobrecarga", "conflito", "comunicação hostil", "comunicacao hostil"], "exams": ["anamnese psicossocial", "srq", "srq-20", "exame clínico", "exame clinico"]},
    {"keywords": ["trânsito", "transito", "motorista", "direção", "direcao"], "exams": ["exame clínico", "exame clinico", "acuidade visual"]},
]


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

    def alter_company_columns_to_text() -> None:
        # Bancos já criados no Render podem ter campos curtos (VARCHAR).
        # Como alguns campos recebem descrições completas de CNAE/atividade,
        # ampliamos para TEXT sem apagar nenhum dado.
        if dialect != "postgresql":
            return
        text_columns = [
            "nome", "cnpj", "endereco", "bairro_cidade", "cep",
            "cnae1", "descricao1", "grau1", "cnae2", "descricao2", "grau2",
            "funcionarios", "data_atual", "data_final", "email", "fone", "data_avaliacao",
        ]
        for column_name in text_columns:
            if has_column("companies", column_name):
                db.session.execute(text(f"ALTER TABLE companies ALTER COLUMN {column_name} TYPE TEXT"))

    def alter_risk_columns_to_text() -> None:
        # A importação em lote usa textos técnicos longos em vários campos.
        # Em versões antigas alguns campos foram criados como VARCHAR(80/120),
        # o que causava erro no PostgreSQL ao importar planilhas completas.
        if dialect != "postgresql":
            return
        text_columns = [
            "acoes", "indicador", "descricao_agente", "possiveis_lesoes",
            "fontes_circunstancias", "epis", "epcs", "ltcat_meio_propagacao",
            "ltcat_insalubridade", "ltcat_grau_insalubridade",
            "ltcat_aposentadoria_especial", "ltcat_enquadramento_tecnico",
            "ltcat_parecer_previdenciario", "ltcat_periodicidade_jornada",
        ]
        for column_name in text_columns:
            if has_column("risks", column_name):
                db.session.execute(text(f"ALTER TABLE risks ALTER COLUMN {column_name} TYPE TEXT"))

    add_column("sectors", "group_id VARCHAR(32)")
    add_column("risks", "ltcat_meio_propagacao TEXT")
    add_column("risks", "ltcat_insalubridade TEXT")
    add_column("risks", "ltcat_grau_insalubridade TEXT")
    add_column("risks", "ltcat_aposentadoria_especial TEXT")
    add_column("risks", "ltcat_enquadramento_tecnico TEXT")
    add_column("risks", "ltcat_parecer_previdenciario TEXT")
    add_column("risks", "ltcat_periodicidade_jornada TEXT")
    # Evolução segura do banco para o cadastro completo de empresas.
    for col in [
        "nome TEXT", "cnpj TEXT", "endereco TEXT", "bairro_cidade TEXT",
        "cep TEXT", "cnae1 TEXT", "descricao1 TEXT", "grau1 TEXT",
        "cnae2 TEXT", "descricao2 TEXT", "grau2 TEXT", "funcionarios TEXT",
        "data_atual TEXT", "data_final TEXT", "email TEXT", "fone TEXT",
        "data_avaliacao TEXT",
    ]:
        add_column("companies", col)
    alter_company_columns_to_text()
    alter_risk_columns_to_text()
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


def _is_numeric_text(value: str) -> bool:
    value = str(value or "").strip()
    return value == "" or value.isdigit()


def _validate_company(company: dict[str, Any]) -> list[str]:
    errors: list[str] = []
    if not company.get("nome"):
        errors.append("Preencha o nome da empresa.")
    if not company.get("cnpj"):
        errors.append("Preencha o CNPJ da empresa.")

    # O grau de risco é sempre número. Essa validação evita o erro técnico do banco
    # quando a descrição da atividade é colada no campo de grau por engano.
    if not _is_numeric_text(company.get("grau1", "")):
        errors.append("O Grau de risco principal deve conter somente número, como 1, 2, 3 ou 4. Verifique se a descrição da atividade foi colocada no campo correto.")
    if not _is_numeric_text(company.get("grau2", "")):
        errors.append("O Grau de risco secundário deve conter somente número, como 1, 2, 3 ou 4. Verifique se a descrição da atividade secundária foi colocada no campo correto.")
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


def _sorted_risk_groups() -> list[dict[str, Any]]:
    return [group.to_dict() for group in RiskGroup.query.order_by(RiskGroup.nome.asc()).all()]


def _sorted_sectors() -> list[dict[str, Any]]:
    return [sector.to_dict() for sector in Sector.query.order_by(Sector.setor.asc()).all()]


def _sorted_exams() -> list[dict[str, Any]]:
    return [exam.to_dict() for exam in Exam.query.order_by(Exam.exame.asc()).all()]


def _sorted_groups() -> list[dict[str, Any]]:
    return [group.to_dict() for group in SectorGroup.query.order_by(SectorGroup.nome.asc()).all()]




def _sorted_companies() -> list[dict[str, Any]]:
    return [company.to_dict() for company in Company.query.order_by(Company.nome.asc()).all()]


def _simple_norm(value: str) -> str:
    import unicodedata
    value = unicodedata.normalize("NFKD", str(value or "")).lower()
    return "".join(ch for ch in value if not unicodedata.combining(ch))


def _extract_text_from_upload(file_storage) -> str:
    """Extrai texto básico de DOCX/PDF para importação inteligente de laudos antigos."""
    if not file_storage or not getattr(file_storage, "filename", ""):
        raise ValueError("Envie pelo menos um arquivo antigo em DOCX ou PDF.")
    filename = secure_filename(file_storage.filename)
    ext = Path(filename).suffix.lower()
    with tempfile.TemporaryDirectory() as tmp:
        tmpdir = Path(tmp)
        path = tmpdir / f"{uuid.uuid4().hex}_{filename}"
        file_storage.save(path)
        if ext == ".docx":
            from docx import Document
            doc = Document(str(path))
            parts: list[str] = []
            parts.extend([p.text for p in doc.paragraphs if (p.text or "").strip()])
            for table in doc.tables:
                for row in table.rows:
                    cells = [cell.text.strip() for cell in row.cells if cell.text.strip()]
                    if cells:
                        parts.append(" | ".join(cells))
            return "\n".join(parts)
        if ext == ".pdf":
            import fitz
            doc = fitz.open(str(path))
            try:
                return "\n".join(page.get_text("text") for page in doc)
            finally:
                doc.close()
        raise ValueError("Envie apenas arquivos .docx ou .pdf para a importação inteligente.")


def _unique_clean_lines(values: list[str]) -> list[str]:
    result: list[str] = []
    seen: set[str] = set()
    for value in values:
        cleaned = re.sub(r"\s+", " ", str(value or "")).strip(" :-|\t")
        if not cleaned or len(cleaned) < 2:
            continue
        key = cleaned.lower()
        if key not in seen:
            seen.add(key)
            result.append(cleaned)
    return result


def _smart_extract_laudo_data(text_value: str) -> dict[str, Any]:
    """Extrai empresa/CNPJ/setores/riscos/exames com heurísticas simples e editáveis."""
    text_value = text_value or ""
    lines = [re.sub(r"\s+", " ", line).strip() for line in text_value.splitlines() if line.strip()]
    cnpj_match = re.search(r"\b\d{2}\.?\d{3}\.?\d{3}/?\d{4}-?\d{2}\b", text_value)
    cnpj = cnpj_match.group(0) if cnpj_match else ""

    empresa = ""
    for pattern in [r"Raz[aã]o Social[:\s]+(.+)", r"EMPRESA[:\s|]+(.+)", r"Empresa[:\s]+(.+)"]:
        match = re.search(pattern, text_value, flags=re.I)
        if match:
            empresa = match.group(1).strip().split("|")[0].strip()
            break
    if not empresa:
        for line in lines[:25]:
            up = line.upper()
            if cnpj and cnpj in line:
                continue
            if any(skip in up for skip in ["PROGRAMA", "PCMSO", "PGR", "LTCAT", "RELATÓRIO", "LAUDO", "DATA", "FONE", "EMAIL", "ELABORAÇÃO"]):
                continue
            if len(line) >= 5 and ("LTDA" in up or "ME" in up or "EPP" in up):
                empresa = line
                break

    setores: list[str] = []
    for match in re.finditer(r"(?:SETOR(?:ES)?|DEPARTAMENTO(?:S)?|GES)\s*[:\-]?\s*([^\n;]+)", text_value, flags=re.I):
        value = match.group(1).strip()
        value = re.split(r"(?:\s{2,}|\||Total|Categoria|Risco|Cargo)", value)[0].strip()
        for item in re.split(r",|/", value):
            if 2 <= len(item.strip()) <= 60:
                setores.append(item.strip())
    # Captura linhas curtas em caixa alta comuns como títulos de setor.
    for line in lines:
        up = line.upper()
        if line == up and 3 <= len(line) <= 45 and not any(word in up for word in ["EMPRESA", "CNPJ", "RISCO", "PGR", "PCMSO", "LTCAT", "IDENTIFICAÇÃO", "CONCLUSÃO", "PLANO", "AÇÃO", "DATA", "EMAIL", "FONE"]):
            if re.search(r"[A-ZÁÉÍÓÚÂÊÔÃÕÇ]", line):
                setores.append(line)

    riscos: list[str] = []
    known_risks = [risk.risco for risk in Risk.query.order_by(Risk.risco.asc()).all()]
    norm_text = _simple_norm(text_value)
    for risk_name in known_risks:
        if risk_name and _simple_norm(risk_name) in norm_text:
            riscos.append(risk_name)
    for match in re.finditer(r"(?:Risco|Perigo|Fator de risco)\s*[:\-]\s*([^\n|;]+)", text_value, flags=re.I):
        val = match.group(1).strip()
        if 4 <= len(val) <= 120:
            riscos.append(val)

    exames: list[str] = []
    known_exams = [exam.exame for exam in Exam.query.order_by(Exam.exame.asc()).all()]
    for exam_name in known_exams:
        if exam_name and _simple_norm(exam_name) in norm_text:
            exames.append(exam_name)

    return {
        "empresa": empresa,
        "cnpj": cnpj,
        "setores": _unique_clean_lines(setores)[:80],
        "riscos": _unique_clean_lines(riscos)[:160],
        "exames": _unique_clean_lines(exames)[:80],
        "texto_preview": text_value[:6000],
    }


def _exam_rule_suggestions_for_groups(groups: list[dict[str, Any]]) -> dict[str, list[str]]:
    exams = _sorted_exams()
    suggestions: dict[str, list[str]] = {}
    for group in groups:
        sector = group.get("sector") or {}
        sector_id = sector.get("id", "")
        combined = " ".join([
            str(risk.get("risco", "")) + " " + str(risk.get("tipo_risco", "")) + " " + str(risk.get("descricao_agente", ""))
            for risk in group.get("risks", [])
        ])
        norm_combined = _simple_norm(combined)
        wanted_terms: list[str] = []
        for rule in EXAM_RULES:
            if any(_simple_norm(keyword) in norm_combined for keyword in rule.get("keywords", [])):
                wanted_terms.extend(rule.get("exams", []))
        wanted_terms = _dedupe_preserve_order(wanted_terms)
        matched_ids: list[str] = []
        for exam in exams:
            exam_norm = _simple_norm(exam.get("exame", ""))
            if any(_simple_norm(term) in exam_norm or exam_norm in _simple_norm(term) for term in wanted_terms if term):
                matched_ids.append(exam["id"])
        suggestions[sector_id] = _dedupe_preserve_order(matched_ids)
    return suggestions


def _build_generation_preview(groups: list[dict[str, Any]], company: dict[str, str]) -> dict[str, Any]:
    total_risks = sum(len(group.get("risks", [])) for group in groups)
    total_exams = sum(len(group.get("exams", [])) for group in groups)
    psychosocial = sum(1 for group in groups for risk in group.get("risks", []) if str(risk.get("tipo_risco", "")).strip().upper() == "ERGONÔMICO PSICOSSOCIAL")
    environmental = sum(1 for group in groups for risk in group.get("risks", []) if str(risk.get("tipo_risco", "")).strip().upper() in TIPOS_RISCO_LTCAT)
    warnings = []
    if not groups:
        warnings.append("Nenhum setor selecionado.")
    for group in groups:
        setor = (group.get("sector") or {}).get("setor", "")
        if not group.get("risks"):
            warnings.append(f"Setor {setor}: sem riscos selecionados.")
        if not group.get("exams"):
            warnings.append(f"Setor {setor}: sem exames selecionados para PCMSO.")
    if environmental == 0:
        warnings.append("Nenhum risco ambiental selecionado para LTCAT; os setores sairão como ausência de riscos ambientais.")
    if not company.get("data_criacao_laudo"):
        warnings.append("Data de criação do laudo não preenchida.")
    return {"setores": len(groups), "riscos": total_risks, "exames": total_exams, "psicossociais": psychosocial, "ambientais": environmental, "avisos": warnings}


def _normalize_filename(value: str) -> str:
    value = re.sub(r"[^A-Za-z0-9._-]+", "_", str(value or "").strip())
    return value.strip("_") or "empresa"


def _send_zip_with_cleanup(path: Path, download_name: str):
    response = send_file(path, as_attachment=True, download_name=download_name, mimetype="application/zip")
    @response.call_on_close
    def _cleanup() -> None:
        try:
            path.unlink(missing_ok=True)
        except Exception:
            pass
    return response


def _normalize_import_header(value: Any) -> str:
    import re
    import unicodedata

    text_value = str(value or "").strip().lower()
    text_value = unicodedata.normalize("NFKD", text_value)
    text_value = "".join(ch for ch in text_value if not unicodedata.combining(ch))
    text_value = re.sub(r"[^a-z0-9]+", "_", text_value).strip("_")
    return text_value


def _xlsx_rows_from_upload(file_storage) -> tuple[list[dict[str, str]], list[str]]:
    """Lê a primeira aba da planilha enviada e retorna linhas por cabeçalho."""
    errors: list[str] = []
    if load_workbook is None:
        return [], ["A biblioteca openpyxl não está instalada. Rode: pip install -r requirements.txt"]
    if not file_storage or not getattr(file_storage, "filename", ""):
        return [], ["Selecione uma planilha Excel para importar."]
    if not file_storage.filename.lower().endswith(".xlsx"):
        return [], ["Envie uma planilha no formato .xlsx."]

    try:
        wb = load_workbook(file_storage.stream, data_only=True)
    except Exception as exc:
        return [], [f"Não foi possível ler a planilha. Verifique se o arquivo é .xlsx válido. Detalhe: {exc}"]

    ws = wb.active
    header_values = [cell.value for cell in ws[1]]
    headers = [_normalize_import_header(value) for value in header_values]
    if not any(headers):
        return [], ["A primeira linha da planilha precisa conter os cabeçalhos."]

    rows: list[dict[str, str]] = []
    for row_index, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        if not any(value not in (None, "") for value in row):
            continue
        item = {headers[i]: str(row[i] if row[i] is not None else "").strip() for i in range(min(len(headers), len(row))) if headers[i]}
        item["__linha"] = str(row_index)
        rows.append(item)
    return rows, errors


def _row_get(row: dict[str, str], *aliases: str) -> str:
    for alias in aliases:
        value = row.get(_normalize_import_header(alias), "")
        if value:
            return value.strip()
    return ""


def _truthy_import_value(value: str) -> str:
    """Normaliza marcações simples de planilha para campos textuais."""
    value = str(value or "").strip()
    if value.lower() in {"sim", "s", "x", "1", "true", "verdadeiro"}:
        return "X"
    return value


def _dedupe_preserve_order(items: list[str]) -> list[str]:
    seen = set()
    out: list[str] = []
    for item in items:
        key = item.strip().lower()
        if item and key not in seen:
            seen.add(key)
            out.append(item)
    return out


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
            data["aet"] = _aet_form_data_from_request()
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
    data["aet"] = _aet_form_data_from_request()
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


def _risk_ids_from_group_ids(group_ids: list[str]) -> list[str]:
    if not group_ids:
        return []
    groups = RiskGroup.query.filter(RiskGroup.id.in_(group_ids)).all()
    risk_ids: list[str] = []
    for group in groups:
        risk_ids.extend([risk.id for risk in group.risks or []])
    return _dedupe_preserve_order(risk_ids)


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
        group_ids = request.form.getlist(f"sector_risk_group_ids_{sector_id}")
        risk_ids = _dedupe_preserve_order(risk_ids + _risk_ids_from_group_ids(group_ids))
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
    return render_template("cadastro.html", risks=_sorted_risks(), risk_groups=_sorted_risk_groups(), options=FORM_OPTIONS)


@app.get("/modelo-importacao-riscos")
def download_risk_import_template():
    return send_file(
        RISK_IMPORT_TEMPLATE,
        as_attachment=True,
        download_name="modelo_importacao_riscos.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.post("/importar-riscos")
def import_risks():
    rows, read_errors = _xlsx_rows_from_upload(request.files.get("arquivo"))
    if read_errors:
        for error in read_errors:
            flash(error, "error")
        return redirect(url_for("cadastro"))

    existing_names = {risk.risco.strip().lower() for risk in Risk.query.all()}
    imported = 0
    skipped = 0
    errors: list[str] = []

    for row in rows:
        line = row.get("__linha", "?")
        risk_data = {
            "id": uuid.uuid4().hex,
            "risco": _row_get(row, "risco", "perigo/fator de risco", "perigo", "nome do risco"),
            "acoes": _row_get(row, "ações preventiva / corretiva", "acoes preventiva corretiva", "ações", "acoes"),
            "indicador": _row_get(row, "indicador de efetividade", "indicador"),
            "tipo_risco": _row_get(row, "tipo de risco", "tipo"),
            "descricao_agente": _row_get(row, "descrição do agente", "descricao do agente"),
            "possiveis_lesoes": _row_get(row, "possíveis lesões ou agravos à saúde", "possiveis lesoes ou agravos a saude", "lesões", "lesoes"),
            "fontes_circunstancias": _row_get(row, "fontes ou circunstâncias", "fontes ou circunstancias") or "Durante o processo de trabalho.",
            "epis": _row_get(row, "epi", "epis"),
            "epcs": _row_get(row, "epc", "epcs"),
            "grau_severidade": _row_get(row, "grau de severidade", "severidade"),
            "grau_possibilidade": _row_get(row, "grau de possibilidade", "possibilidade", "probabilidade"),
            "grau_nivel_risco": _row_get(row, "grau de nível de risco", "grau de nivel de risco", "nível de risco", "nivel de risco"),
            "ltcat_meio_propagacao": _row_get(row, "ltcat meio de propagação / via de exposição", "meio de propagação", "via de exposição", "ltcat meio propagacao"),
            "ltcat_periodicidade_jornada": _row_get(row, "ltcat periodicidade / jornada de exposição", "periodicidade / jornada de exposição", "jornada de exposição"),
            "ltcat_insalubridade": _row_get(row, "ltcat insalubridade", "insalubridade") or "Não",
            "ltcat_grau_insalubridade": _row_get(row, "ltcat grau de insalubridade", "grau de insalubridade") or "Não aplicável",
            "ltcat_aposentadoria_especial": _row_get(row, "ltcat aposentadoria especial", "aposentadoria especial") or "Não",
            "ltcat_enquadramento_tecnico": _row_get(row, "ltcat enquadramento técnico", "enquadramento técnico", "enquadramento tecnico"),
            "ltcat_parecer_previdenciario": _row_get(row, "ltcat parecer previdenciário", "parecer previdenciário", "parecer previdenciario"),
        }

        if risk_data["risco"].strip().lower() in existing_names:
            skipped += 1
            continue

        validation_errors = _validate_risk(risk_data)
        if validation_errors:
            skipped += 1
            errors.extend([f"Linha {line}: {error}" for error in validation_errors])
            continue

        db.session.add(_risk_from_dict(risk_data))
        existing_names.add(risk_data["risco"].strip().lower())
        imported += 1

    db.session.commit()
    if imported:
        flash(f"{imported} risco(s) importado(s) com sucesso.", "success")
    if skipped:
        flash(f"{skipped} linha(s) não foram importadas por erro ou duplicidade.", "error")
    for error in errors[:10]:
        flash(error, "error")
    if len(errors) > 10:
        flash(f"Há mais {len(errors) - 10} erro(s) não exibidos. Corrija a planilha e importe novamente.", "error")
    return redirect(url_for("cadastro"))


@app.get("/modelo-importacao-setores")
def download_sector_import_template():
    return send_file(
        SECTOR_IMPORT_TEMPLATE,
        as_attachment=True,
        download_name="modelo_importacao_setores.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.post("/importar-setores")
def import_sectors():
    group_id = _field("grupo_id_importacao")
    group = db.session.get(SectorGroup, group_id) if group_id else None
    if not group:
        flash("Selecione o grupo onde os setores serão cadastrados antes de importar a planilha.", "error")
        return redirect(url_for("setores"))

    rows, read_errors = _xlsx_rows_from_upload(request.files.get("arquivo"))
    if read_errors:
        for error in read_errors:
            flash(error, "error")
        return redirect(url_for("setores"))

    grouped: dict[str, list[dict[str, str]]] = {}
    errors: list[str] = []
    for row in rows:
        line = row.get("__linha", "?")
        setor = _row_get(row, "setor", "nome do setor")
        cargo = _row_get(row, "cargo", "função", "funcao")
        cbo = _row_get(row, "cbo")
        n_func = _row_get(row, "número de funcionários", "numero de funcionarios", "nº funcionários", "n_func", "funcionários")
        descricao = _row_get(row, "descrição da atividade", "descricao da atividade", "atividade")
        if not setor:
            errors.append(f"Linha {line}: preencha o Setor.")
            continue
        if not cargo:
            errors.append(f"Linha {line}: preencha o Cargo.")
            continue
        if not cbo:
            errors.append(f"Linha {line}: preencha o CBO.")
            continue
        if not n_func:
            errors.append(f"Linha {line}: preencha o número de funcionários.")
            continue
        if not descricao:
            errors.append(f"Linha {line}: preencha a descrição da atividade.")
            continue
        grouped.setdefault(setor.strip(), []).append({
            "id": uuid.uuid4().hex,
            "cargo": cargo.strip(),
            "cbo": cbo.strip(),
            "n_func": n_func.strip(),
            "descricao": descricao.strip(),
        })

    imported_sectors = 0
    imported_cargos = 0
    for setor_nome, cargos in grouped.items():
        sector = Sector.query.filter(
            db.func.lower(Sector.setor) == setor_nome.lower(),
            Sector.group_id == group.id,
        ).first()
        if not sector:
            sector = Sector(id=uuid.uuid4().hex, setor=setor_nome, group_id=group.id, cargos=[])
            db.session.add(sector)
            imported_sectors += 1

        existing_cargos = list(sector.cargos) if isinstance(sector.cargos, list) else []
        existing_keys = {(str(c.get("cargo", "")).strip().lower(), str(c.get("cbo", "")).strip().lower()) for c in existing_cargos}
        for cargo in cargos:
            key = (cargo["cargo"].strip().lower(), cargo["cbo"].strip().lower())
            if key not in existing_keys:
                existing_cargos.append(cargo)
                existing_keys.add(key)
                imported_cargos += 1
        sector.cargos = existing_cargos
        sector.updated_at = datetime.utcnow()

    db.session.commit()
    if imported_sectors or imported_cargos:
        flash(f"Importação concluída: {imported_sectors} setor(es) novo(s) e {imported_cargos} cargo(s) cadastrado(s) no grupo {group.nome}.", "success")
    if errors:
        flash(f"{len(errors)} linha(s) não foram importadas por erro.", "error")
        for error in errors[:10]:
            flash(error, "error")
        if len(errors) > 10:
            flash(f"Há mais {len(errors) - 10} erro(s) não exibidos.", "error")
    return redirect(url_for("setores"))


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


def _aet_form_data_from_request() -> dict[str, Any]:
    """Dados detalhados da AET informados no wizard de geração.

    Fica salvo junto da configuração da empresa para permitir regeneração sem
    perder análise ergonômica, observações e conclusões por setor.
    """
    sector_ids = request.form.getlist("pgr_sector_ids")
    general = {
        "tipo_aet": _field("aet_tipo") or "AET completa com formulário ergonômico",
        "tipo_documento": _field("aet_tipo_documento") or _field("aet_tipo") or "AET - Análise Ergonômica do Trabalho",
        "motivo_analise": _field("aet_motivo_analise") or "Atendimento à NR-17 e integração com o PGR",
        "responsavel_tecnico": _field("aet_responsavel_tecnico"),
        "metodologia": request.form.getlist("aet_metodologia"),
        "origem_dados": request.form.getlist("aet_origem_dados"),
        "objetivo_complementar": _field("aet_objetivo_complementar"),
        "criterios_analise": _field("aet_criterios_analise"),
        "limitacoes_analise": _field("aet_limitacoes_analise"),
        "condicao_ergonomica_geral": _field("aet_condicao_ergonomica_geral"),
        "conclusao_geral_manual": _field("aet_conclusao_geral"),
    }
    by_sector: dict[str, Any] = {}
    for sector_id in sector_ids:
        by_sector[sector_id] = {
            "postura_predominante": request.form.getlist(f"aet_postura_{sector_id}"),
            "tipo_atividade": _field(f"aet_tipo_atividade_{sector_id}"),
            "exigencia_fisica": _field(f"aet_exigencia_fisica_{sector_id}"),
            "exigencia_cognitiva": _field(f"aet_exigencia_cognitiva_{sector_id}"),
            "levantamento_cargas": _field(f"aet_levantamento_cargas_{sector_id}"),
            "movimentos_repetitivos": _field(f"aet_movimentos_repetitivos_{sector_id}"),
            "atencao_concentracao": _field(f"aet_atencao_concentracao_{sector_id}"),
            "atendimento_publico": _field(f"aet_atendimento_publico_{sector_id}"),
            "autonomia": _field(f"aet_autonomia_{sector_id}"),
            "metas_prioridades": _field(f"aet_metas_prioridades_{sector_id}"),
            "comunicacao": _field(f"aet_comunicacao_{sector_id}"),
            "ritmo_trabalho": _field(f"aet_ritmo_trabalho_{sector_id}"),
            "pausas": _field(f"aet_pausas_{sector_id}"),
            "mobiliario": _field(f"aet_mobiliario_{sector_id}"),
            "ambiente": _field(f"aet_ambiente_{sector_id}"),
            "organizacao": _field(f"aet_organizacao_{sector_id}"),
            "equipamentos": _field(f"aet_equipamentos_{sector_id}"),
            "fatores_organizacionais": request.form.getlist(f"aet_fatores_organizacionais_{sector_id}"),
            "medidas_recomendadas": request.form.getlist(f"aet_medidas_recomendadas_{sector_id}"),
            "queixas": _field(f"aet_queixas_{sector_id}"),
            "observacoes": _field(f"aet_observacoes_{sector_id}"),
            "recomendacoes": _field(f"aet_recomendacoes_{sector_id}"),
            "prioridade": _field(f"aet_prioridade_{sector_id}"),
            "prazo": _field(f"aet_prazo_{sector_id}"),
            "responsavel": _field(f"aet_responsavel_{sector_id}"),
            "conclusao_setor": _field(f"aet_conclusao_setor_{sector_id}"),
        }
    return {"general": general, "by_sector": by_sector}


def _gerar_form_state_from_request() -> dict[str, Any]:
    sector_ids = request.form.getlist("pgr_sector_ids")
    risks_by_sector = {sector_id: request.form.getlist(f"sector_risk_ids_{sector_id}") for sector_id in sector_ids}
    risk_groups_by_sector = {sector_id: request.form.getlist(f"sector_risk_group_ids_{sector_id}") for sector_id in sector_ids}
    exams_by_sector = {sector_id: request.form.getlist(f"sector_exam_ids_{sector_id}") for sector_id in sector_ids}
    return {
        "company_id": _field("company_id"),
        "data_criacao_laudo": _field("data_criacao_laudo"),
        "ajuste_psicossocial": "1" if request.form.get("ajuste_psicossocial") == "1" else "",
        "data_da_revisao": _field("data_da_revisao"),
        "profile_id": _field("profile_id"),
        "profile_name": _field("profile_name"),
        "selected_sector_ids": sector_ids,
        "selected_risk_ids_by_sector": risks_by_sector,
        "selected_risk_group_ids_by_sector": risk_groups_by_sector,
        "selected_exam_ids_by_sector": exams_by_sector,
        "aet": _aet_form_data_from_request(),
    }




def _report_profile_state_from_request() -> dict[str, Any]:
    """Estado recarregável da tela Gerar laudos."""
    return _gerar_form_state_from_request()


def _sorted_report_profiles() -> list[dict[str, Any]]:
    return [profile.to_dict() for profile in ReportProfile.query.order_by(ReportProfile.updated_at.desc()).all()]


def _save_report_profile_from_form(auto: bool = False) -> ReportProfile | None:
    company_id = _field("company_id")
    if not company_id:
        return None
    company = db.session.get(Company, company_id)
    if not company:
        return None

    state = _report_profile_state_from_request()
    data_criacao = state.get("data_criacao_laudo", "")
    timestamp = datetime.now().strftime("%d/%m/%Y %H:%M")
    default_name = f"Geração {timestamp}"
    nome = _field("profile_name") or default_name

    # Modo automático: mantém uma configuração padrão por empresa, atualizada
    # sempre que o usuário finaliza um laudo completo.
    profile_id = _field("profile_id")
    profile = db.session.get(ReportProfile, profile_id) if profile_id else None
    if profile is None and auto:
        profile = ReportProfile.query.filter_by(company_id=company_id, nome="Última geração salva").first()
        nome = "Última geração salva"

    if profile is None:
        profile = ReportProfile(id=uuid.uuid4().hex, company_id=company_id, nome=nome)
        db.session.add(profile)
    else:
        profile.company_id = company_id
        profile.nome = nome

    profile.data_criacao_laudo = data_criacao
    profile.ajuste_psicossocial = state.get("ajuste_psicossocial", "")
    profile.data_da_revisao = state.get("data_da_revisao", "")
    profile.state = state
    profile.updated_at = datetime.utcnow()
    db.session.commit()
    return profile


def _month_year_from_text(value: str) -> tuple[str, str]:
    """Extrai mês por extenso e ano de textos como 05/06/2026, 06/2026 ou Junho/2026."""
    months = {
        1: "JANEIRO", 2: "FEVEREIRO", 3: "MARÇO", 4: "ABRIL", 5: "MAIO", 6: "JUNHO",
        7: "JULHO", 8: "AGOSTO", 9: "SETEMBRO", 10: "OUTUBRO", 11: "NOVEMBRO", 12: "DEZEMBRO",
    }
    value = (value or "").strip()
    year_match = re.search(r"(20\d{2})", value)
    year = year_match.group(1) if year_match else str(datetime.now().year)
    nums = re.findall(r"\d+", value)
    month_num = None
    if len(nums) >= 2:
        # Em dd/mm/aaaa usa o segundo número; em mm/aaaa usa o primeiro.
        month_num = int(nums[1]) if len(nums[0]) <= 2 and len(nums[1]) <= 2 else int(nums[0])
    elif len(nums) == 1 and len(nums[0]) <= 2:
        month_num = int(nums[0])
    if month_num and 1 <= month_num <= 12:
        return months[month_num], year
    normalized = value.upper()
    for name in months.values():
        if name in normalized:
            return name, year
    return months[datetime.now().month], year


def _replace_simple_docx_text(doc, replacements: dict[str, str]) -> None:
    """Substituição simples preservando o estilo do primeiro run quando necessário."""
    from copy import deepcopy

    def replace_paragraph(paragraph):
        full_text = paragraph.text or ""
        if not any(key in full_text for key in replacements):
            return
        for run in paragraph.runs:
            text_value = run.text
            for key, val in replacements.items():
                text_value = text_value.replace(key, val)
            run.text = text_value
        full_text = paragraph.text or ""
        if any(key in full_text for key in replacements):
            for key, val in replacements.items():
                full_text = full_text.replace(key, val)
            rpr = deepcopy(paragraph.runs[0]._r.rPr) if paragraph.runs and paragraph.runs[0]._r.rPr is not None else None
            paragraph.clear()
            run = paragraph.add_run(full_text)
            if rpr is not None:
                run._r.insert(0, rpr)

    def walk_part(part):
        for paragraph in part.paragraphs:
            replace_paragraph(paragraph)
        for table in part.tables:
            for row in table.rows:
                for cell in row.cells:
                    walk_part(cell)

    walk_part(doc)
    for section in doc.sections:
        for part in [section.header, section.footer, section.first_page_header, section.first_page_footer, section.even_page_header, section.even_page_footer]:
            walk_part(part)


def _prepare_link_docx(template_path: Path, output_path: Path, empresa: str, data_criacao: str, mes_extenso: str | None = None) -> Path:
    from docx import Document
    doc = Document(str(template_path))
    mes, ano = _month_year_from_text(data_criacao)
    if mes_extenso:
        mes = mes_extenso.strip().upper()
    replacements = {
        "T E M NAKASHIMA - ME": empresa or "",
        "{{EMPRESA}}": empresa or "",
        "{{MES DE CRIAÇÃO POR EXTENSO}}": mes,
        "{{MES DE CRIAÇÃO POR\nEXTENSO}}": mes,
        "{{ANO DE CRIAÇÃO}}": ano,
    }
    _replace_simple_docx_text(doc, replacements)
    # Os modelos enviados tinham o ano 2026 fixo; ajusta apenas nas páginas de link.
    for paragraph in doc.paragraphs:
        if "INCLUSÃO NO PGR EM" in (paragraph.text or "") and f"DE {ano}" not in paragraph.text:
            for run in paragraph.runs:
                run.text = run.text.replace("DE 2026", f"DE {ano}")
    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(output_path))
    return output_path


def _clear_header_footer_part(part) -> None:
    """Remove cabeçalho/rodapé herdado ao inserir anexos no Word final."""
    try:
        part.is_linked_to_previous = False
    except Exception:
        pass
    element = getattr(part, "_element", None)
    if element is not None:
        for child in list(element):
            element.remove(child)
    try:
        part.add_paragraph()
    except Exception:
        pass


def _clear_section_headers_footers(section) -> None:
    try:
        section.different_first_page_header_footer = True
    except Exception:
        pass
    for part in [
        section.header,
        section.footer,
        section.first_page_header,
        section.first_page_footer,
        section.even_page_header,
        section.even_page_footer,
    ]:
        _clear_header_footer_part(part)


def _render_pdf_pages_to_docx_body(doc, pdf_path: Path, images_dir: Path) -> None:
    """Insere as páginas do PDF como imagens no corpo do DOCX atual.

    Usado para preservar o relatório psicossocial exatamente como o PDF.
    """
    try:
        import fitz  # PyMuPDF
        from docx.shared import Inches, Pt
        from docx.enum.text import WD_ALIGN_PARAGRAPH
    except Exception as exc:  # pragma: no cover
        raise RuntimeError("A biblioteca PyMuPDF/python-docx não está instalada. Verifique o requirements.txt no Render.") from exc

    images_dir.mkdir(parents=True, exist_ok=True)
    pdf = fitz.open(str(pdf_path))
    if pdf.page_count == 0:
        raise ValueError("O PDF do Relatório Psicossocial não possui páginas.")
    if pdf.page_count > PSICOSSOCIAL_MAX_PAGES:
        raise ValueError(f"O PDF possui {pdf.page_count} páginas. O limite atual é {PSICOSSOCIAL_MAX_PAGES} páginas para evitar travamentos no Render.")

    first_rect = pdf[0].rect
    section = doc.sections[-1]
    page_margin = 0.05
    section.page_width = Inches(first_rect.width / 72)
    section.page_height = Inches(first_rect.height / 72)
    section.top_margin = Inches(page_margin)
    section.bottom_margin = Inches(page_margin)
    section.left_margin = Inches(page_margin)
    section.right_margin = Inches(page_margin)
    section.header_distance = Inches(0)
    section.footer_distance = Inches(0)
    _clear_section_headers_footers(section)

    zoom = PSICOSSOCIAL_RENDER_DPI / 72
    matrix = fitz.Matrix(zoom, zoom)

    try:
        for page_index in range(pdf.page_count):
            page = pdf[page_index]
            image_path = images_dir / f"pagina_{page_index + 1:03d}.jpg"
            pix = page.get_pixmap(matrix=matrix, alpha=False, colorspace=fitz.csRGB)
            pix.save(str(image_path), output="jpeg", jpg_quality=PSICOSSOCIAL_JPEG_QUALITY)

            paragraph = doc.add_paragraph()
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            paragraph.paragraph_format.space_before = Pt(0)
            paragraph.paragraph_format.space_after = Pt(0)
            paragraph.paragraph_format.line_spacing = 1
            run = paragraph.add_run()

            rect = page.rect
            max_width_in = (first_rect.width / 72) - (page_margin * 2)
            max_height_in = (first_rect.height / 72) - (page_margin * 2) - 0.03
            width_in = rect.width / 72 if rect.width else max_width_in
            height_in = rect.height / 72 if rect.height else max_height_in
            scale = min(max_width_in / width_in, max_height_in / height_in, 1)
            run.add_picture(str(image_path), width=Inches(width_in * scale))
    finally:
        pdf.close()


def _append_pdf_as_images_to_docx(base_docx: Path, pdf_path: Path, output_docx: Path) -> Path:
    try:
        from docx import Document
        from docx.enum.section import WD_SECTION
    except Exception as exc:  # pragma: no cover
        raise RuntimeError("A biblioteca python-docx não está instalada. Verifique o requirements.txt no Render.") from exc

    doc = Document(str(base_docx))
    doc.add_section(WD_SECTION.NEW_PAGE)
    output_docx.parent.mkdir(parents=True, exist_ok=True)
    # As imagens temporárias são incorporadas ao DOCX no momento do save.
    # Depois disso, podem ser removidas para não encher o disco do Render.
    with tempfile.TemporaryDirectory() as img_tmp:
        images_dir = Path(img_tmp) / "psicossocial_paginas"
        _render_pdf_pages_to_docx_body(doc, pdf_path, images_dir)
        doc.save(str(output_docx))
    return output_docx


def _convert_pdf_to_docx(pdf_path: Path, output_docx: Path) -> Path:
    """Converte PDF para DOCX preservando o visual de cada página.

    A conversão antiga por extração de texto/tabelas quebrava o layout dos
    relatórios psicossociais, criando sobreposições, faixas pretas e tabelas
    deformadas. Para documentos que precisam apenas ser anexados ao Word final,
    a forma mais fiel é rasterizar cada página do PDF e inserir a página inteira
    como imagem em um DOCX. O conteúdo fica não editável, mas visualmente igual
    ao PDF original.
    """
    try:
        import fitz  # PyMuPDF
        from docx import Document
        from docx.shared import Inches, Pt
    except Exception as exc:  # pragma: no cover
        raise RuntimeError("A biblioteca PyMuPDF/python-docx não está instalada. Verifique o requirements.txt no Render.") from exc

    output_docx.parent.mkdir(parents=True, exist_ok=True)
    images_dir = output_docx.parent / f"{output_docx.stem}_paginas"
    images_dir.mkdir(parents=True, exist_ok=True)

    pdf = fitz.open(str(pdf_path))
    if pdf.page_count == 0:
        raise ValueError("O PDF do Relatório Psicossocial não possui páginas.")

    doc = Document()
    section = doc.sections[0]

    # Usa o tamanho real da primeira página do PDF. A maioria dos relatórios
    # psicossociais é gerada em tamanho único.
    first_rect = pdf[0].rect
    page_width = Inches(first_rect.width / 72)
    page_height = Inches(first_rect.height / 72)
    section.page_width = page_width
    section.page_height = page_height
    # Margem mínima para evitar que o Word crie páginas em branco por causa
    # da altura da imagem somada ao parágrafo. O PDF já possui suas próprias
    # margens visuais, então essa redução não altera a aparência do conteúdo.
    page_margin = 0.05
    section.top_margin = Inches(page_margin)
    section.bottom_margin = Inches(page_margin)
    section.left_margin = Inches(page_margin)
    section.right_margin = Inches(page_margin)
    section.header_distance = Inches(0)
    section.footer_distance = Inches(0)
    _clear_section_headers_footers(section)

    # 200 dpi é um bom equilíbrio entre fidelidade visual e tamanho do arquivo.
    zoom = PSICOSSOCIAL_RENDER_DPI / 72
    matrix = fitz.Matrix(zoom, zoom)

    try:
        for page_index in range(pdf.page_count):
            page = pdf[page_index]
            image_path = images_dir / f"pagina_{page_index + 1:03d}.jpg"
            pix = page.get_pixmap(matrix=matrix, alpha=False, colorspace=fitz.csRGB)
            pix.save(str(image_path), output="jpeg", jpg_quality=PSICOSSOCIAL_JPEG_QUALITY)

            paragraph = doc.add_paragraph()
            paragraph.paragraph_format.space_before = Pt(0)
            paragraph.paragraph_format.space_after = Pt(0)
            paragraph.paragraph_format.line_spacing = 1
            run = paragraph.add_run()

            # Ajusta pelo tamanho útil da página. A pequena redução evita páginas
            # em branco causadas pelo parágrafo que o Word mantém após imagens.
            rect = page.rect
            max_width_in = (first_rect.width / 72) - (page_margin * 2)
            max_height_in = (first_rect.height / 72) - (page_margin * 2) - 0.03
            width_in = rect.width / 72 if rect.width else max_width_in
            height_in = rect.height / 72 if rect.height else max_height_in
            scale = min(max_width_in / width_in, max_height_in / height_in, 1)
            run.add_picture(str(image_path), width=Inches(width_in * scale))

    finally:
        pdf.close()

    doc.save(str(output_docx))
    return output_docx


def _merge_docx_files(paths: list[Path], output_path: Path) -> Path:
    if not paths:
        raise ValueError("Envie pelo menos um arquivo Word para juntar.")
    try:
        from docx import Document
        from docxcompose.composer import Composer
    except Exception as exc:  # pragma: no cover
        raise RuntimeError("As bibliotecas docxcompose/python-docx não estão instaladas. Verifique o requirements.txt no Render.") from exc

    master = Document(str(paths[0]))
    composer = Composer(master)
    for path in paths[1:]:
        # Garante que cada anexo comece em uma nova página.
        composer.doc.add_page_break()
        composer.append(Document(str(path)))
    output_path.parent.mkdir(parents=True, exist_ok=True)
    composer.save(str(output_path))
    return output_path


def _save_uploaded_file(file_storage, folder: Path, allowed_exts: set[str], label: str) -> Path:
    if not file_storage or not getattr(file_storage, "filename", ""):
        raise ValueError(f"Envie o arquivo: {label}.")
    filename = secure_filename(file_storage.filename)
    ext = Path(filename).suffix.lower()
    if ext not in allowed_exts:
        raise ValueError(f"O arquivo {label} precisa estar em: {', '.join(sorted(allowed_exts))}.")
    folder.mkdir(parents=True, exist_ok=True)
    path = folder / f"{uuid.uuid4().hex}_{filename}"
    file_storage.save(path)
    return path


def _build_combined_pgr_aet_psychosocial(pgr_docx: Path, aet_docx: Path, psicossocial_pdf: Path, output_path: Path, empresa: str, data_criacao: str, mes_extenso: str | None = None) -> Path:
    workdir = output_path.parent / f"merge_{uuid.uuid4().hex}"
    workdir.mkdir(parents=True, exist_ok=True)
    try:
        link_pgr = workdir / "link_pgr_para_aet.docx"
        link_aet = workdir / "link_aet_para_psicossocial.docx"
        merged_without_psy = workdir / "pgr_aet_links.docx"

        _prepare_link_docx(LINK_PGR_AET_TEMPLATE, link_pgr, empresa, data_criacao, mes_extenso)
        _prepare_link_docx(LINK_AET_PSICOSSOCIAL_TEMPLATE, link_aet, empresa, data_criacao, mes_extenso)

        # Primeiro junta os arquivos editáveis. Depois adiciona o Relatório
        # Psicossocial como páginas-imagem em uma nova seção sem cabeçalho/rodapé.
        # Isso evita as distorções geradas por conversão PDF -> DOCX por texto.
        _merge_docx_files([pgr_docx, link_pgr, aet_docx, link_aet], merged_without_psy)
        return _append_pdf_as_images_to_docx(merged_without_psy, psicossocial_pdf, output_path)
    finally:
        shutil.rmtree(workdir, ignore_errors=True)


def _send_docx_with_cleanup(path: Path, download_name: str):
    """Envia o DOCX e limpa o arquivo temporário quando a resposta terminar.

    Evita acúmulo de arquivos grandes no disco do Render após juntar documentos.
    """
    response = send_file(
        path,
        as_attachment=True,
        download_name=download_name,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )

    @response.call_on_close
    def _cleanup() -> None:
        try:
            path.unlink(missing_ok=True)
        except Exception:
            pass

    return response


def _is_ajax_request() -> bool:
    return request.headers.get("X-Requested-With") == "XMLHttpRequest"


def _ajax_error(message: str, status_code: int = 400):
    return jsonify({"ok": False, "error": message}), status_code


def _gerar_context(form_state: dict[str, Any] | None = None) -> dict[str, Any]:
    today = datetime.now().strftime("%m/%Y")
    next_year = datetime.now().replace(year=datetime.now().year + 1).strftime("%m/%Y")
    risks = _sorted_risks()
    ltcat_risks = [risk for risk in risks if str(risk.get("tipo_risco", "")).strip().upper() in TIPOS_RISCO_LTCAT]
    return {
        "risks": risks,
        "ltcat_risks": ltcat_risks,
        "risk_groups": _sorted_risk_groups(),
        "sectors": _sorted_sectors(),
        "exams": _sorted_exams(),
        "groups": _sorted_groups(),
        "grouped_sectors": _grouped_sectors(),
        "companies": _sorted_companies(),
        "aet_presets": AET_CNAE_PRESETS,
        "report_profiles": _sorted_report_profiles(),
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




@app.post("/salvar-configuracao-laudo")
def save_report_profile():
    profile = _save_report_profile_from_form(auto=False)
    if profile:
        flash("Configuração da empresa salva. Você poderá carregar essa seleção novamente depois.", "success")
    else:
        flash("Selecione uma empresa antes de salvar a configuração da geração.", "error")
    return _render_gerar_with_current_form()


@app.post("/configuracao-laudo/<profile_id>/excluir")
def delete_report_profile(profile_id: str):
    profile = db.session.get(ReportProfile, profile_id)
    if profile:
        db.session.delete(profile)
        db.session.commit()
        flash("Configuração salva excluída.", "success")
    return redirect(url_for("gerar"))


@app.route("/juntar-arquivos")
def juntar_arquivos():
    return render_template("juntar.html", companies=_sorted_companies(), today=datetime.now().strftime("%d/%m/%Y"))


@app.route("/importar-laudo-antigo", methods=["GET", "POST"])
def importar_laudo_antigo():
    extracted = None
    if request.method == "POST":
        try:
            file_storage = request.files.get("arquivo")
            text_value = _extract_text_from_upload(file_storage)
            extracted = _smart_extract_laudo_data(text_value)
            flash("Leitura concluída. Revise os dados extraídos antes de salvar no sistema.", "success")
        except Exception as exc:
            flash(f"Erro ao importar laudo antigo: {exc}", "error")
    return render_template("importar_laudo.html", extracted=extracted or {}, groups=_sorted_groups(), risks=_sorted_risks(), exams=_sorted_exams())


@app.post("/importar-laudo-antigo/salvar")
def salvar_importacao_laudo_antigo():
    empresa = _field("empresa")
    cnpj = _field("cnpj")
    setores_text = _field("setores_extraidos")
    riscos_text = _field("riscos_extraidos")
    exames_text = _field("exames_extraidos")
    group_id = _field("grupo_id") or None

    if not empresa and not cnpj and not setores_text and not riscos_text and not exames_text:
        flash("Nenhum dado foi informado para salvar.", "error")
        return redirect(url_for("importar_laudo_antigo"))

    company_created = False
    if empresa or cnpj:
        company = None
        if cnpj:
            company = Company.query.filter(db.func.lower(Company.cnpj) == cnpj.lower()).first()
        if not company and empresa:
            company = Company.query.filter(db.func.lower(Company.nome) == empresa.lower()).first()
        if not company:
            company = Company(nome=empresa or "Empresa importada", cnpj=cnpj or "")
            db.session.add(company)
            company_created = True
        else:
            if empresa and not company.nome:
                company.nome = empresa
            if cnpj and not company.cnpj:
                company.cnpj = cnpj

    sector_count = 0
    for line in _unique_clean_lines(setores_text.splitlines()):
        existing = Sector.query.filter(db.func.lower(Sector.setor) == line.lower()).first()
        if not existing:
            db.session.add(Sector(setor=line, group_id=group_id, cargos=[{
                "id": uuid.uuid4().hex,
                "cargo": "A DEFINIR",
                "cbo": "A DEFINIR",
                "n_func": "1",
                "descricao": "Atividades importadas de laudo antigo; revisar e detalhar conforme função.",
            }]))
            sector_count += 1

    risk_count = 0
    for line in _unique_clean_lines(riscos_text.splitlines()):
        existing = Risk.query.filter(db.func.lower(Risk.risco) == line.lower()).first()
        if not existing:
            db.session.add(Risk(
                risco=line,
                acoes="Revisar ações preventivas/corretivas conforme atividade e atualizar com treinamentos NR aplicáveis.",
                indicador="Acompanhar implementação das medidas, registros de orientação e ausência de ocorrências relacionadas.",
                tipo_risco="ERGONÔMICO",
                possiveis_lesoes="Revisar possíveis lesões ou agravos conforme o risco importado.",
                fontes_circunstancias="Informação importada de laudo antigo; revisar fontes ou circunstâncias.",
                epis="A definir conforme avaliação técnica.",
                epcs="A definir conforme avaliação técnica.",
                grau_severidade="MÉDIO",
                grau_possibilidade="POSSÍVEL",
                grau_nivel_risco="MODERADO",
            ))
            risk_count += 1

    exam_count = 0
    for line in _unique_clean_lines(exames_text.splitlines()):
        existing = Exam.query.filter(db.func.lower(Exam.exame) == line.lower()).first()
        if not existing:
            db.session.add(Exam(exame=line, periodicidade="Conforme PCMSO"))
            exam_count += 1

    db.session.commit()
    flash(f"Importação salva: empresa {'criada' if company_created else 'atualizada/verificada'}, {sector_count} setor(es), {risk_count} risco(s) e {exam_count} exame(s) novo(s). Revise os cadastros importados antes de gerar laudos.", "success")
    return redirect(url_for("importar_laudo_antigo"))


@app.post("/juntar-arquivos")
def merge_files_avulso():
    try:
        with tempfile.TemporaryDirectory() as tmp:
            tmpdir = Path(tmp)
            pgr_docx = _save_uploaded_file(request.files.get("pgr_docx"), tmpdir, {".docx"}, "PGR em Word (.docx)")
            aet_docx = _save_uploaded_file(request.files.get("aet_docx"), tmpdir, {".docx"}, "AET em Word (.docx)")
            psicossocial_pdf = _save_uploaded_file(request.files.get("psicossocial_pdf"), tmpdir, {".pdf"}, "Relatório Psicossocial em PDF")
            company_id = _field("company_id")
            company = db.session.get(Company, company_id) if company_id else None
            empresa = _field("empresa_avulsa") or (company.nome if company else "")
            data_criacao = _field("data_criacao_laudo")
            mes_extenso = _field("mes_extenso")
            if not empresa:
                raise ValueError("Informe ou selecione a empresa para preencher as páginas intermediárias.")
            if not data_criacao and not mes_extenso:
                raise ValueError("Informe a data de criação do laudo ou o mês por extenso para preencher as páginas intermediárias.")
            stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
            output_path = OUTPUT_DIR / f"pgr_aet_psicossocial_{stamp}.docx"
            _build_combined_pgr_aet_psychosocial(pgr_docx, aet_docx, psicossocial_pdf, output_path, empresa, data_criacao, mes_extenso)
            return _send_docx_with_cleanup(output_path, "PGR_AET_RELATORIO_PSICOSSOCIAL.docx")
    except Exception as exc:
        message = f"Erro ao juntar arquivos: {exc}"
        if _is_ajax_request():
            return _ajax_error(message)
        flash(message, "error")
        return redirect(url_for("juntar_arquivos"))

@app.route("/grupos-riscos")
def risk_groups():
    return render_template("risk_groups.html", risk_groups=_sorted_risk_groups(), risks=_sorted_risks())


def _risk_group_form_data(existing_id: str | None = None) -> dict[str, Any]:
    return {
        "id": existing_id or uuid.uuid4().hex,
        "nome": _field("nome"),
        "descricao": _field("descricao"),
        "risk_ids": request.form.getlist("risk_ids"),
    }


def _validate_risk_group(data: dict[str, Any], existing_id: str | None = None) -> list[str]:
    errors: list[str] = []
    if not data.get("nome"):
        errors.append("Preencha o nome do grupo de riscos.")
    if not data.get("risk_ids"):
        errors.append("Selecione pelo menos um risco para o grupo.")
    if data.get("nome"):
        query = RiskGroup.query.filter(db.func.lower(RiskGroup.nome) == data["nome"].strip().lower())
        if existing_id:
            query = query.filter(RiskGroup.id != existing_id)
        if query.first():
            errors.append("Já existe um grupo de riscos com esse nome.")
    return errors


def _apply_risk_group_data(model: RiskGroup, data: dict[str, Any]) -> RiskGroup:
    model.nome = data["nome"].strip()
    model.descricao = data.get("descricao", "").strip()
    risk_ids = _dedupe_preserve_order(data.get("risk_ids", []))
    model.risks = Risk.query.filter(Risk.id.in_(risk_ids)).all() if risk_ids else []
    model.updated_at = datetime.utcnow()
    return model


@app.post("/grupo-risco/novo")
def create_risk_group():
    data = _risk_group_form_data()
    errors = _validate_risk_group(data)
    if errors:
        for error in errors:
            flash(error, "error")
        return redirect(url_for("risk_groups"))
    model = RiskGroup(id=data["id"])
    _apply_risk_group_data(model, data)
    db.session.add(model)
    db.session.commit()
    flash("Grupo de riscos cadastrado com sucesso.", "success")
    return redirect(url_for("risk_groups"))


@app.route("/grupo-risco/<group_id>/editar", methods=["GET", "POST"])
def edit_risk_group(group_id: str):
    model = db.session.get(RiskGroup, group_id)
    if not model:
        flash("Grupo de riscos não encontrado.", "error")
        return redirect(url_for("risk_groups"))
    if request.method == "POST":
        data = _risk_group_form_data(existing_id=group_id)
        errors = _validate_risk_group(data, existing_id=group_id)
        if errors:
            for error in errors:
                flash(error, "error")
            return render_template("risk_group_form.html", group={**model.to_dict(), **data}, risks=_sorted_risks(), title="Editar grupo de riscos")
        _apply_risk_group_data(model, data)
        db.session.commit()
        flash("Grupo de riscos atualizado com sucesso.", "success")
        return redirect(url_for("risk_groups"))
    return render_template("risk_group_form.html", group=model.to_dict(), risks=_sorted_risks(), title="Editar grupo de riscos")


@app.post("/grupo-risco/<group_id>/excluir")
def delete_risk_group(group_id: str):
    model = db.session.get(RiskGroup, group_id)
    if model:
        db.session.delete(model)
        db.session.commit()
        flash("Grupo de riscos excluído.", "success")
    return redirect(url_for("risk_groups"))


@app.post("/api/risco/novo")
def api_create_risk():
    """Cadastro rápido usado na tela Gerar laudos, sem recarregar a página."""
    new_risk = _form_to_risk()
    errors = _validate_risk(new_risk)

    if new_risk.get("risco"):
        existing = Risk.query.filter(db.func.lower(Risk.risco) == new_risk["risco"].strip().lower()).first()
        if existing:
            errors.append("Já existe um risco cadastrado com esse nome.")

    if errors:
        return jsonify({"ok": False, "errors": errors}), 400

    risk_model = _risk_from_dict(new_risk)
    db.session.add(risk_model)
    db.session.commit()
    return jsonify({"ok": True, "message": "Risco cadastrado com sucesso.", "risk": risk_model.to_dict()})


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
        for group in RiskGroup.query.all():
            if risk_model in (group.risks or []):
                group.risks.remove(risk_model)
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
        # A LTCAT volta a ser automática: todos os riscos ambientais
        # selecionados no bloco "Riscos por setor" entram no laudo.
        # Tipos não ambientais são ignorados, e setor sem risco ambiental
        # gera o bloco de AUSÊNCIA DE RISCOS.
        risk_ids = request.form.getlist(f"sector_risk_ids_{sector_id}")
        group_ids = request.form.getlist(f"sector_risk_group_ids_{sector_id}")
        risk_ids = _dedupe_preserve_order(risk_ids + _risk_ids_from_group_ids(group_ids))
        selected_risks = [risks[risk_id] for risk_id in risk_ids if risk_id in risks]
        selected_risks = [
            risk for risk in selected_risks
            if str(risk.get("tipo_risco", "")).strip().upper() in TIPOS_RISCO_LTCAT
        ]
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




@app.post("/api/sugerir-exames")
def api_sugerir_exames():
    groups, errors = _selected_sector_risk_groups()
    # Para sugestão, não bloqueia setor sem risco; só retorna avisos.
    suggestions = _exam_rule_suggestions_for_groups(groups)
    return jsonify({"ok": True, "suggestions": suggestions, "warnings": errors, "rules": EXAM_RULES})


@app.post("/api/previsualizar-geracao")
def api_previsualizar_geracao():
    groups, errors = _selected_sector_risk_groups()
    company = _company_payload_from_form()
    preview = _build_generation_preview(groups, company)
    preview["erros"] = errors
    return jsonify({"ok": True, "preview": preview})


@app.post("/gerar-aet-completa")
def generate_complete_aet():
    groups, errors = _selected_sector_risk_groups()
    company = _company_payload_from_form()
    errors.extend(_validate_complete_report_fields(company, "AET"))
    if errors:
        for error in errors:
            flash(error, "error")
        return _render_gerar_with_current_form()
    try:
        empresa = company.get("empresa") or company.get("nome", "")
        cnpj = company.get("cnpj", "")
        data_atual = company.get("data_atual", "")
        data_final = company.get("data_final", "")
        _save_report_profile_from_form(auto=True)
        return _send_generated_docx(
            generate_aet_docx,
            groups,
            "aet_completa",
            "AET_COMPLETA.docx",
            empresa,
            cnpj,
            data_atual,
            data_final,
            company,
        )
    except Exception as exc:
        flash(f"Erro ao gerar AET completa: {exc}", "error")
        return _render_gerar_with_current_form()


@app.post("/gerar-pacote-empresa")
def generate_company_zip_package():
    groups, errors = _selected_sector_risk_groups()
    ltcat_groups, ltcat_errors = _selected_sector_ltcat_groups()
    company = _company_payload_from_form()
    errors.extend(ltcat_errors)
    errors.extend(_validate_complete_report_fields(company, "pacote ZIP"))
    if errors:
        for error in errors:
            flash(error, "error")
        return _render_gerar_with_current_form()
    try:
        empresa = company.get("empresa") or company.get("nome", "")
        cnpj = company.get("cnpj", "")
        data_atual = company.get("data_atual", "")
        data_final = company.get("data_final", "")
        safe_empresa = _normalize_filename(empresa)
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        zip_path = OUTPUT_DIR / f"pacote_{safe_empresa}_{stamp}.zip"
        with tempfile.TemporaryDirectory() as tmp:
            tmpdir = Path(tmp)
            pgr_path = tmpdir / "PGR_COMPLETO.docx"
            pcmso_path = tmpdir / "PCMSO_COMPLETO.docx"
            ltcat_path = tmpdir / "LTCAT_COMPLETO.docx"
            aet_path = tmpdir / "AET_COMPLETA.docx"
            generate_complete_pgr_docx(groups, pgr_path, empresa, cnpj, data_atual, data_final, company)
            generate_complete_pcmso_docx(groups, pcmso_path, empresa, cnpj, data_atual, data_final, company)
            generate_complete_ltcat_docx(ltcat_groups, ltcat_path, empresa, cnpj, data_atual, data_final, company)
            generate_aet_docx(groups, aet_path, empresa, cnpj, data_atual, data_final, company)
            with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
                for path in [pgr_path, pcmso_path, ltcat_path, aet_path]:
                    zf.write(path, arcname=path.name)
                resumo = tmpdir / "RESUMO_DA_GERACAO.txt"
                preview = _build_generation_preview(groups, company)
                resumo.write_text(
                    f"Empresa: {empresa}\nCNPJ: {cnpj}\nVigência: {data_atual} a {data_final}\n"
                    f"Setores: {preview['setores']}\nRiscos: {preview['riscos']}\nExames: {preview['exames']}\n"
                    f"Riscos psicossociais: {preview['psicossociais']}\nRiscos ambientais para LTCAT: {preview['ambientais']}\n"
                    f"Avisos:\n- " + "\n- ".join(preview.get('avisos') or ['Nenhum aviso.']),
                    encoding="utf-8"
                )
                zf.write(resumo, arcname="RESUMO_DA_GERACAO.txt")
        _save_report_profile_from_form(auto=True)
        return _send_zip_with_cleanup(zip_path, f"PACOTE_LAUDOS_{safe_empresa}.zip")
    except Exception as exc:
        flash(f"Erro ao gerar pacote ZIP da empresa: {exc}", "error")
        return _render_gerar_with_current_form()


@app.post("/gerar-pgr-aet-psicossocial")
def generate_pgr_aet_psychosocial():
    groups, errors = _selected_sector_risk_groups()
    company = _company_payload_from_form()
    errors.extend(_validate_complete_report_fields(company, "PGR + AET + Relatório Psicossocial"))
    if errors:
        if _is_ajax_request():
            return _ajax_error("\n".join(errors))
        for error in errors:
            flash(error, "error")
        return _render_gerar_with_current_form()
    try:
        with tempfile.TemporaryDirectory() as tmp:
            tmpdir = Path(tmp)
            psicossocial_pdf = _save_uploaded_file(request.files.get("psicossocial_pdf"), tmpdir, {".pdf"}, "Relatório Psicossocial em PDF")
            empresa = company.get("empresa") or company.get("nome", "")
            cnpj = company.get("cnpj", "")
            data_atual = company.get("data_atual", "")
            data_final = company.get("data_final", "")
            pgr_docx = tmpdir / "pgr_gerado.docx"
            aet_upload = request.files.get("aet_docx")
            if aet_upload and getattr(aet_upload, "filename", ""):
                aet_docx = _save_uploaded_file(aet_upload, tmpdir, {".docx"}, "AET em Word (.docx)")
            else:
                aet_docx = tmpdir / "aet_gerada.docx"
                generate_aet_docx(groups, aet_docx, empresa, cnpj, data_atual, data_final, company)
            generate_complete_pgr_docx(groups, pgr_docx, empresa, cnpj, data_atual, data_final, company)
            _save_report_profile_from_form(auto=True)
            stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
            output_path = OUTPUT_DIR / f"pgr_aet_psicossocial_{stamp}.docx"
            _build_combined_pgr_aet_psychosocial(
                pgr_docx,
                aet_docx,
                psicossocial_pdf,
                output_path,
                empresa,
                company.get("data_criacao_laudo") or company.get("data_avaliacao") or company.get("data_atual", ""),
                _field("mes_extenso"),
            )
            return _send_docx_with_cleanup(output_path, "PGR_AET_RELATORIO_PSICOSSOCIAL.docx")
    except Exception as exc:
        message = f"Erro ao gerar PGR + AET + Relatório Psicossocial: {exc}"
        if _is_ajax_request():
            return _ajax_error(message)
        flash(message, "error")
        return _render_gerar_with_current_form()

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
        _save_report_profile_from_form(auto=True)
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
        _save_report_profile_from_form(auto=True)
        return _send_generated_docx(
            generate_complete_ltcat_docx,
    generate_aet_docx,
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
        _save_report_profile_from_form(auto=True)
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
