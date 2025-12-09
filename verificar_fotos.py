import pandas as pd
import os

def encontrar_caminho_imagem(pasta_imagens, codigo_produto):
    """
    Tenta encontrar o arquivo de imagem para o código do produto, 
    verificando as extensões mais comuns e ignorando a caixa (case).
    """
    # Lista de extensões comuns que serão testadas
    extensoes = ['.jpg', '.jpeg', '.png']
    
    # Monta o nome base do arquivo (ex: '01.0000')
    nome_base = codigo_produto.strip()
    
    # Itera sobre as extensões para checar a existência do arquivo
    for ext in extensoes:
        # Tenta o nome do arquivo em minúsculas
        nome_arquivo = f"{nome_base}{ext}"
        caminho_completo = os.path.join(pasta_imagens, nome_arquivo)
        
        if os.path.exists(caminho_completo):
            return caminho_completo
            
        # Tenta o nome do arquivo em MAIÚSCULAS (caso o sistema seja sensível)
        nome_arquivo_upper = f"{nome_base}{ext.upper()}"
        caminho_completo_upper = os.path.join(pasta_imagens, nome_arquivo_upper)
        
        if os.path.exists(caminho_completo_upper):
            return caminho_completo_upper
            
    # Se não encontrou nenhuma das extensões/casos, retorna None
    return None

def verificar_produtos_sem_foto(caminho_excel, caminho_pasta_imagens, coluna_codigo='Código', coluna_estoque='Estoque'):
    """
    Lê o arquivo Excel e filtra APENAS os produtos para os quais 
    NENHUMA imagem foi encontrada na pasta.
    """
    print("Iniciando a leitura do arquivo Excel...")
    
    try:
        # Define a coluna 'Código' como string para preservar zeros à esquerda
        df = pd.read_excel(caminho_excel, dtype={coluna_codigo: str}) 
        print(f"Total de {len(df)} produtos encontrados no relatório.")
    except Exception as e:
        print(f"ERRO ao ler o Excel. Verifique o caminho ou nome das colunas. Detalhe: {e}")
        return

    produtos_sem_foto = []
    print("Iniciando a verificação de imagens...")

    # Itera sobre cada linha do DataFrame
    for index, row in df.iterrows():
        codigo_produto = str(row[coluna_codigo]).strip()
        
        # CHAMA A FUNÇÃO ROBUSTA para buscar a imagem
        caminho_da_imagem = encontrar_caminho_imagem(caminho_pasta_imagens, codigo_produto)
        
        # 🌟 O Filtro Mágico: Verifica se a imagem NÃO foi encontrada (é None)
        if caminho_da_imagem is None:
            # Se NENHUM caminho de imagem foi encontrado, adiciona o produto à lista de ausentes
            produtos_sem_foto.append(row.to_dict())

    # Cria e salva o novo DataFrame APENAS com os produtos sem foto
    if produtos_sem_foto:
        df_sem_foto = pd.DataFrame(produtos_sem_foto)
        nome_arquivo_saida = 'Produtos_Sem_Foto.xlsx'
        
        try:
            df_sem_foto.to_excel(nome_arquivo_saida, index=False) 
            print("-" * 50)
            print(f"✅ CONCLUÍDO! {len(df_sem_foto)} produtos foram encontrados sem foto.")
            print(f"✅ O relatório contendo **APENAS OS PRODUTOS SEM FOTO** foi salvo em: **{nome_arquivo_saida}**")
            
            # Verificação de estoque
            if coluna_estoque in df_sem_foto.columns:
                produtos_com_estoque_sem_foto = df_sem_foto[df_sem_foto[coluna_estoque].astype(float, errors='ignore') > 0]
                if not produtos_com_estoque_sem_foto.empty:
                    print(f"⚠️ Atenção: {len(produtos_com_estoque_sem_foto)} desses produtos **têm estoque** e estão sem foto. Priorize-os!")

        except Exception as e:
            print(f"ERRO ao salvar o arquivo Excel de saída: {e}")
    else:
        print("-" * 50)
        print("🎉 EXCELENTE! Todos os produtos no relatório têm uma imagem correspondente na pasta.")


# --- CONFIGURAÇÕES ---
# VERIFIQUE E AJUSTE ESTES VALORES!
CAMINHO_DO_RELATORIO = 'Relatorio_Produtos.xlsx' 
CAMINHO_DA_PASTA_DE_IMAGENS = './Imagens_Produtos' 

# Estes nomes DEVEM ser IGUAIS aos cabeçalhos da sua planilha
NOME_COLUNA_CODIGO = 'Código'   
NOME_COLUNA_ESTOQUE = 'Estoque' 

# --- EXECUTAR O SCRIPT ---
if __name__ == '__main__':
    verificar_produtos_sem_foto(
        caminho_excel=CAMINHO_DO_RELATORIO,
        caminho_pasta_imagens=CAMINHO_DA_PASTA_DE_IMAGENS,
        coluna_codigo=NOME_COLUNA_CODIGO,
        coluna_estoque=NOME_COLUNA_ESTOQUE
    )