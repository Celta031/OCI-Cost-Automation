# SDD: OCI Cost Updater Produtivo

## Objetivo

Este documento descreve o desenho tecnico do OCI Cost Updater, uma ferramenta Python para buscar custos mensais da Oracle Cloud Infrastructure (OCI), agrupar valores por compartment e atualizar uma planilha Excel operacional.

O projeto deve continuar simples para execucao mensal, mas com configuracao externa, logs claros, validacoes, testes e separacao suficiente para uso produtivo e publicacao segura no GitHub.

## Contexto Atual

A versao inicial concentrava configuracao, chamada OCI, atualizacao Excel e relatorio em um unico arquivo. Isso funcionava para uso local, mas exigia editar variaveis diretamente no codigo e dificultava testes, automacao e publicacao publica.

O fluxo produtivo passa a usar `oci_cost_updater.py` como entrada unica. O arquivo `atualizar_custos_oci.py` fica tratado como legado/local e nao deve ser necessario para novas execucoes.

## Arquitetura

O script fica organizado em camadas pequenas:

- Configuracao: combina argumentos CLI, variaveis do ambiente e `.env`, com precedencia CLI > ambiente > `.env`.
- Validacao: confirma arquivo Excel, colunas, linha inicial, datas UTC e periodo valido antes de chamar a OCI.
- Cliente OCI: usa o SDK oficial e o arquivo `~/.oci/config`, opcionalmente com profile informado.
- Agregacao: transforma os itens retornados pela OCI em `{compartment: custo}` e soma compartments repetidos.
- Excel: abre a planilha, mapeia compartments pela coluna de origem e escreve custos na coluna destino.
- Relatorio: registra linhas atualizadas, linhas zeradas, linhas ignoradas e compartments OCI nao mapeados.

## Interfaces Publicas

Execucao principal:

```bash
python oci_cost_updater.py --excel-path sua-planilha.xlsx --sheet "OCI CONSUMO" --start 2026-04-21T00:00:00Z --end 2026-05-21T00:00:00Z --compartment-column A --target-column S --start-row 5
```

Modo de validacao sem salvar:

```bash
python oci_cost_updater.py --dry-run
```

Configuracao local via `.env`:

```env
OCI_EXCEL_PATH=sua-planilha.xlsx
OCI_SHEET_NAME=OCI CONSUMO
OCI_START_ROW=5
OCI_COMPARTMENT_COLUMN=A
OCI_TARGET_COLUMN=S
OCI_BILLING_START=2026-04-21T00:00:00Z
OCI_BILLING_END=2026-05-21T00:00:00Z
OCI_PROFILE=DEFAULT
OCI_COMPARTMENT_DEPTH=6
```

## Tratamento de Erros

Erros de configuracao retornam exit code `2`, incluindo planilha inexistente, aba inexistente, colunas invalidas, datas invalidas ou periodo invertido.

Falhas operacionais retornam exit code `1`, incluindo erro de autenticacao OCI, permissao, indisponibilidade da API ou falha ao salvar a planilha.

Execucao bem-sucedida retorna exit code `0`.

## Seguranca

Credenciais OCI permanecem fora do repositorio e devem continuar no caminho padrao do SDK (`~/.oci/config`) ou equivalente suportado pela OCI. O projeto nao deve armazenar tenancy, fingerprint, chave privada, planilhas reais ou arquivos `.env` com dados produtivos.

O `.gitignore` deve bloquear `.env`, planilhas Excel, ambientes virtuais, caches e logs.

## Performance

O volume esperado e pequeno, com execucao mensal. A chamada OCI e feita uma vez por periodo e a planilha e percorrida uma vez a partir da linha inicial. A complexidade principal e linear em relacao ao numero de linhas da planilha.

A agregacao usa dicionarios para lookup O(1), evitando buscas repetidas por compartment.

## Testes

Os testes automatizados devem cobrir:

- Validacao de datas ISO UTC.
- Precedencia de configuracao CLI sobre `.env`.
- Agregacao de custos, incluindo `computed_amount=None`.
- Normalizacao de nomes de compartments.
- Calculo de compartments OCI nao mapeados.
- Atualizacao de workbook temporario.
- Modo `dry-run` sem persistir alteracoes.
- Erro claro para aba inexistente.

## Decisoes

- `oci_cost_updater.py` e a entrada principal publica.
- `.env` simplifica o uso mensal local.
- CLI permite sobrescrever valores sem editar codigo.
- A planilha e salva diretamente por padrao.
- `--dry-run` existe para conferencia antes do salvamento real.
