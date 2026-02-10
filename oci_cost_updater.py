import oci
import openpyxl
from datetime import datetime

# ==========================================
# CONFIGURAÇÕES (EDITE AQUI)
# ==========================================
# Caminho da sua planilha Excel
ARQUIVO_EXCEL = 'sua-planilha.xlsx' 

# Nome da aba (sheet) que você quer editar (ex: 'Sheet1', 'Custos', 'OCI CONSUMO', etc)
# Se deixar None, ele pega a aba que estiver ativa/aberta ao salvar.
NOME_ABA = 'Sheet1' 

# Linha onde começam os dados (ignorando cabeçalho).
# Exemplo: Se seus dados começam na linha 5, use LINHA_INICIAL = 5
LINHA_INICIAL = 2 

# Coluna onde estão os nomes dos Compartments (ex: 'A', 'B', etc)
COLUNA_COMPARTMENT = 'A'

# Coluna onde você quer INSERIR os novos valores (ex: 'C', 'P', 'AB', etc)
COLUNA_DESTINO = 'B' 

# Período de Faturamento (Formato YYYY-MM-DDTHH:MM:SSZ)
# Exemplo: 1º de Janeiro de 2025 a 1º de Fevereiro de 2025
DATA_INICIO = "2025-01-01T00:00:00Z"
DATA_FIM    = "2025-02-01T00:00:00Z"

# ==========================================
# LÓGICA DO SCRIPT
# ==========================================

def get_oci_costs(start_time, end_time):
    """
    Busca os custos da OCI agrupados por compartment.
    
    Args:
        start_time (str): Data/hora inicial no formato ISO 8601
        end_time (str): Data/hora final no formato ISO 8601
    
    Returns:
        dict: Dicionário com {nome_compartment: valor_custo}
    """
    print("Conectando à OCI para buscar custos...")
    try:
        # Carrega configurações do arquivo ~/.oci/config
        config = oci.config.from_file()
        usage_client = oci.usage_api.UsageapiClient(config)
        
        # Monta os detalhes da requisição de custos
        request_details = oci.usage_api.models.RequestSummarizedUsagesDetails(
            tenant_id=config['tenancy'],
            time_usage_started=start_time,
            time_usage_ended=end_time,
            granularity='MONTHLY',
            query_type='COST',
            group_by=['compartmentName'],
            compartment_depth=6  # Define nível de profundidade dos subcompartments (6 geralmente é suficiente)
        )
        
        # Faz a requisição à API da OCI
        response = usage_client.request_summarized_usages(request_details)
        
        # Processa os resultados e agrupa por nome do compartment
        custos_dict = {}
        for item in response.data.items:
            comp_name = item.compartment_name if item.compartment_name else "Root"
            
            # Trata computed_amount None como 0.0
            valor = item.computed_amount if item.computed_amount is not None else 0.0
            
            # Se o compartment já existe, soma o valor
            if comp_name in custos_dict:
                custos_dict[comp_name] += valor
            else:
                custos_dict[comp_name] = valor
                
        print(f"Sucesso! Encontrados custos para {len(custos_dict)} compartments na OCI.")
        return custos_dict

    except Exception as e:
        print(f"Erro ao conectar na OCI: {e}")
        return None

def update_excel(custos_oci):
    """
    Atualiza a planilha Excel com os custos obtidos da OCI.
    
    Args:
        custos_oci (dict): Dicionário com {nome_compartment: valor_custo}
    
    Returns:
        list: Lista de compartments que foram encontrados e atualizados na planilha
    """
    print(f"Abrindo planilha: {ARQUIVO_EXCEL}...")
    try:
        # Abre a planilha Excel
        wb = openpyxl.load_workbook(ARQUIVO_EXCEL)
        if NOME_ABA:
            ws = wb[NOME_ABA]
        else:
            ws = wb.active
            
        # Lista para rastrear quais compartments da OCI foram usados na planilha
        compartments_encontrados = []
        
        print("Atualizando linhas...")
        
        # Itera sobre as linhas da planilha a partir da LINHA_INICIAL
        for row in range(LINHA_INICIAL, ws.max_row + 1):
            cell_name = ws[f"{COLUNA_COMPARTMENT}{row}"].value
            cell_dest = ws[f"{COLUNA_DESTINO}{row}"]
            
            if not cell_name:
                continue  # Pula linhas vazias
                
            # Normaliza o nome (remove espaços extras e converte para string)
            nome_planilha = str(cell_name).strip()
            
            # Verifica se o compartment tem custo na OCI
            if nome_planilha in custos_oci:
                valor = custos_oci[nome_planilha]
                cell_dest.value = float(valor)  # Insere o valor
                compartments_encontrados.append(nome_planilha)
            else:
                # Compartment existe na planilha, mas não veio custo da OCI (R$ 0)
                cell_dest.value = 0.0
        
        # Salva o arquivo
        wb.save(ARQUIVO_EXCEL)
        print("Planilha atualizada e salva com sucesso!")
        
        return compartments_encontrados

    except Exception as e:
        print(f"Erro ao manipular Excel: {e}")
        return []

def verificar_discrepancias(custos_oci, encontrados_planilha):
    """
    Verifica e exibe discrepâncias entre compartments na OCI e na planilha.
    
    Args:
        custos_oci (dict): Dicionário com custos obtidos da OCI
        encontrados_planilha (list): Lista de compartments encontrados na planilha
    """
    print("\n----------------RELATÓRIO----------------")
    
    # Verifica compartments que estão na OCI mas NÃO estão na planilha
    oci_keys = set(custos_oci.keys())
    planilha_keys = set(encontrados_planilha)
    
    sobra_oci = oci_keys - planilha_keys
    
    if sobra_oci:
        print("⚠️  ATENÇÃO: Os seguintes compartments geraram custo na OCI mas NÃO foram encontrados na planilha:")
        for comp in sobra_oci:
            print(f" - {comp}: R$ {custos_oci[comp]:.2f}")
    else:
        print("✅ Todos os compartments da OCI foram mapeados na planilha.")
        
    print("-----------------------------------------")

# ==========================================
# EXECUÇÃO
# ==========================================
if __name__ == "__main__":
    # 1. Busca dados na OCI
    dados_oci = get_oci_costs(DATA_INICIO, DATA_FIM)
    
    if dados_oci:
        # 2. Atualiza Excel e pega lista de processados
        processados = update_excel(dados_oci)
        
        # 3. Gera relatório de discrepâncias
        verificar_discrepancias(dados_oci, processados)
