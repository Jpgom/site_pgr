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
- Geração de LTCAT completo automática: todos os riscos ambientais (FÍSICO, QUÍMICO e BIOLÓGICO) selecionados nos setores entram no LTCAT.
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

## Atualização V20

- Cadastro rápido de risco dentro da tela **Gerar laudos**, sem recarregar a página.
- O novo risco é salvo via rota `/api/risco/novo` e aparece imediatamente nas listas de riscos por setor.
- Opção para marcar automaticamente o novo risco nos setores já selecionados.

## V21 - Grupos de riscos

- Nova aba **Grupos de riscos**.
- Permite criar pacotes de riscos, por exemplo: LIMPEZA, MANUTENÇÃO, ADMINISTRATIVO.
- Na tela **Gerar laudos**, cada setor agora permite aplicar um ou mais grupos de riscos.
- Ao marcar um grupo, todos os riscos daquele grupo são selecionados automaticamente para o setor.
- Mesmo usando um grupo, ainda é possível marcar riscos isolados ou desmarcar riscos específicos.
- A LTCAT continua automática: usa apenas os riscos ambientais selecionados no setor, ou seja, FÍSICO, QUÍMICO e BIOLÓGICO.
