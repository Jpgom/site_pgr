# Site de Cadastro de Riscos SST - PGR Completo

Sistema em Flask para cadastrar riscos, setores e cargos e gerar o **PGR completo em Word** usando o modelo principal.

## O que esta versão faz

- Usa banco de dados para guardar os cadastros.
- Funciona localmente com SQLite automaticamente.
- Funciona no Render com PostgreSQL usando a variável `DATABASE_URL`.
- Mantém os riscos cadastrados no sistema.
- Permite cadastrar, editar e excluir setores/cargos.
- Permite apagar todos os setores/cargos de uma vez, sem apagar os riscos.
- Gera o PGR completo preenchendo:
  - Empresa;
  - CNPJ;
  - Data atual / início da vigência;
  - Data final da vigência;
  - Relação Função x Atividade;
  - Descritivo dos setores;
  - Inventário/Risco PGR por setor;
  - Plano de Ação.

## Regras mantidas

- Cada setor do Inventário/Risco PGR começa em nova página.
- Os riscos de cada setor saem um abaixo do outro.
- A frase “NENHUM FATOR DE RISCO PSICOSSOCIAL...” aparece somente quando o setor **não** tiver risco do tipo `ERGONÔMICO PSICOSSOCIAL`.
- Quando houver risco psicossocial no setor, essa frase é removida.
- No Plano de Ação:
  - Riscos comuns usam `Data atual / início da vigência` como **Prazo de Implantação**.
  - Riscos comuns usam `Data final da vigência` como **Prazo Reavaliação**.
  - Riscos do tipo `ERGONÔMICO PSICOSSOCIAL` usam **30 DIAS** e **180 DIAS**.
- As células coloridas continuam com texto preto.
- O PCMSO não entra no PGR completo nesta etapa.

## Como rodar localmente

```bash
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
python app.py
```

Acesse:

```text
http://127.0.0.1:5000
```

O banco local será criado automaticamente em:

```text
instance/sst_riscos.db
```

## Como subir no Render

### Opção 1 - Blueprint

1. Suba esta pasta para um repositório no GitHub.
2. No Render, escolha a opção de criar serviço via Blueprint.
3. Selecione o arquivo `render.yaml` deste projeto.
4. O Render criará o serviço web e o banco PostgreSQL conforme as configurações do arquivo.

### Opção 2 - Manual

1. Crie um Web Service no Render apontando para o repositório.
2. Configure:

```text
Build Command: pip install -r requirements.txt
Start Command: gunicorn app:app
```

3. Crie um banco PostgreSQL no Render.
4. Copie a connection string do banco para a variável de ambiente:

```text
DATABASE_URL
```

5. Crie também a variável:

```text
SECRET_KEY
```

com qualquer valor forte/aleatório.

## Estrutura principal

```text
app.py                  Rotas, banco de dados e fluxo do sistema
word_generator.py       Geração dos arquivos Word
modelos/                Modelos DOCX usados na geração
templates/              Telas HTML
static/                 CSS e JS
requirements.txt        Dependências do projeto
Procfile                Start command para deploy
render.yaml             Blueprint para Render
```
