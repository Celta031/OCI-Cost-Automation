# OCI Cost Updater - Automação de Custos OCI

Script Python para automatizar a atualização de planilhas Excel com custos mensais da Oracle Cloud Infrastructure (OCI).

## 📋 Descrição

Este projeto busca automaticamente os custos de todos os compartments da OCI em um período específico e atualiza uma planilha Excel existente com os valores obtidos. Além disso, gera um relatório de discrepâncias para identificar compartments com custo na OCI que não estão mapeados na planilha.

## 📁 Estrutura do Projeto

- **`oci_cost_updater.py`** - Script principal com configurações personalizáveis
- **`.gitignore`** - Arquivo que protege dados sensíveis de serem versionados
- **`README.md`** - Documentação do projeto

## 🚀 Funcionalidades

- ✅ Conexão automática com OCI usando credenciais configuradas
- ✅ Busca de custos por compartment (incluindo subcompartments até 6 níveis de profundidade)
- ✅ Atualização automática de planilha Excel
- ✅ Inserção de valores zerados para compartments sem custo
- ✅ Relatório de discrepâncias entre OCI e planilha
- ✅ Período de faturamento configurável

## 📦 Pré-requisitos

- Python 3.7 ou superior
- Conta Oracle Cloud Infrastructure (OCI)
- Arquivo de configuração OCI (`~/.oci/config`) - [Como configurar](https://docs.oracle.com/en-us/iaas/Content/API/Concepts/sdkconfig.htm)
- Planilha Excel com estrutura definida (nomes de compartments em uma coluna)

## 🔧 Instalação

1. Clone o repositório:
```bash
git clone <url-do-repositorio>
cd automacao_custos_mensal
```

2. Crie e ative um ambiente virtual:
```bash
python -m venv venv
# Windows
venv\Scripts\activate
# Linux/Mac
source venv/bin/activate
```

3. Instale as dependências:
```bash
pip install oci openpyxl
```

4. Configure suas credenciais OCI:
   - Crie o arquivo `~/.oci/config` seguindo a [documentação oficial](https://docs.oracle.com/en-us/iaas/Content/API/Concepts/sdkconfig.htm)
   - Garanta que sua chave API esteja configurada corretamente

5. Edite as configurações no arquivo `oci_cost_updater.py` conforme sua necessidade (veja seção abaixo)

## ⚙️ Configuração

Edite as variáveis no início do arquivo `oci_cost_updater.py`:

```python
# Caminho da planilha Excel
ARQUIVO_EXCEL = 'sua-planilha.xlsx'

# Nome da aba da planilha
NOME_ABA = 'OCI CONSUMO'

# Linha inicial dos dados (após cabeçalho)
LINHA_INICIAL = 5

# Coluna com nomes dos compartments
COLUNA_COMPARTMENT = 'A'

# Coluna onde os valores serão inseridos
COLUNA_DESTINO = 'AB'

# Período de faturamento (formato YYYY-MM-DDTHH:MM:SSZ)
DATA_INICIO = "2025-12-21T00:00:00Z"
DATA_FIM    = "2026-01-21T00:00:00Z"
```

## 📊 Estrutura da Planilha

A planilha deve ter a seguinte estrutura:

- **Coluna A**: Nomes dos compartments OCI
- **Coluna de Destino** (configurável): Onde os valores de custo serão inseridos
- **Linha Inicial**: Primeira linha com dados (após cabeçalho)

## 🎯 Uso

Execute o script:

```bash
python oci_cost_updater.py
```

O script irá:
1. Conectar à OCI e buscar os custos do período configurado
2. Abrir a planilha Excel especificada
3. Atualizar os valores na coluna de destino
4. Salvar a planilha automaticamente
5. Exibir relatório de discrepâncias

## 📝 Exemplo de Saída

```
Conectando à OCI para buscar custos...
Sucesso! Encontrados custos para 15 compartments na OCI.
Abrindo planilha: SUPTEC - GINS - CONSUMO OCI 2025.xlsx...
Atualizando linhas...
Planilha atualizada e salva com sucesso!

----------------RELATÓRIO----------------
✅ Todos os compartments da OCI foram mapeados na planilha.
-----------------------------------------
```

## ⚠️ Observações Importantes

- O script utiliza `compartment_depth=6` para buscar subcompartments em até 6 níveis
- Compartments não encontrados na planilha receberão valor R$ 0,00
- Compartments da OCI não mapeados na planilha serão listados no relatório
- **NUNCA versione** seu arquivo `~/.oci/config` ou chaves privadas
- **NUNCA versione** planilhas Excel com dados reais (já incluído no `.gitignore`)
- Revise permissões de acesso aos compartments OCI
- Use políticas de IAM apropriadas para acesso read-only de custos

## 🔒 Segurança

### ⚠️ Antes de Fazer Commit

Certifique-se de que:
1. Você não adicionou dados sensíveis nas configurações do script
2. O arquivo `.gitignore` está protegendo planilhas Excel e credenciais
3. Suas chaves privadas OCI não estão no repositório

## 🤝 Contribuindo

Contribuições são bem-vindas! Sinta-se à vontade para:
- Reportar bugs
- Sugerir novas funcionalidades
- Enviar pull requests

## 📄 Licença

Este projeto é de uso livre. Adapte conforme suas necessidades.

## 👥 Autor

Desenvolvido para automatizar processos de gestão de custos OCI.
