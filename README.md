# Automação SST - PGR, PCMSO e LTCAT

Sistema Flask para cadastro de riscos, setores/cargos, empresas, exames e geração de laudos em Word.

## Funcionalidades

- Cadastro de riscos ocupacionais com dados para PGR, PCMSO, Plano de Ação e LTCAT.
- Importação de riscos em massa por modelo Excel.
- Cadastro de setores/cargos organizados por grupos.
- Importação de setores/cargos em massa por modelo Excel, vinculando os setores ao grupo escolhido.
- Cadastro completo de empresas para preencher a Identificação da Empresa nos laudos.
- Cadastro de exames para o PCMSO.
- Geração de PGR completo.
- Geração de PCMSO completo.
- Geração de LTCAT completo.
- Exportações avulsas: Plano de Ação, Inventário PGR, Riscos/Exames PCMSO, Relação Função x Atividade e Descritivo Setor.

## Rodar localmente

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

## Deploy no Render

O projeto está pronto para Render com PostgreSQL.

Arquivos importantes:

- `render.yaml`
- `Procfile`
- `runtime.txt`
- `requirements.txt`

No Render, use Blueprint apontando para o repositório GitHub. O sistema usa `DATABASE_URL` quando existir; localmente usa SQLite automático em `instance/sst_riscos.db`.

## Atualizações

Para atualizar no Render, substitua os arquivos, depois rode:

```bash
git add .
git commit -m "atualizacao do sistema"
git push
```

O banco existente é preservado. As novas tabelas/colunas são criadas automaticamente ao iniciar o sistema.
