from Classes.BAS_Generico import *
from Classes.BAS_Arquivo import *
from Classes.SIS_Conexao import clsConexaoBancoDados

# ************************************
# Sub Rotina para Executar os Scripts
# ************************************
def executaScript(caminho, nomeArqScript, bancoDados, nomeArqlogExec):
    # --- Verifica o Arquivo de Script
    if not ExisteArquivo(caminho, nomeArqScript):
        RegistraLinhaArquivo(nomeArqlogExec, f'O arquivo de script não foi localizado em: {caminho + nomeArqScript}.', True)
        return

    # --- CONEXÃO COM O BANDO DE DADOS
    RegistraLinhaArquivo(nomeArqlogExec, f'Conectando ao SQL Server banco de dados < {bancoDados} > para aplicar/executar o script...', True)

    conexao = clsConexaoBancoDados.ConexaoSQLServer()
    conexao.servidor = 'LAPTOP-NC2UQK3P\SQLEXPRESS'  # Pegar do Arquivo de Configuração
    conexao.bancoDados = bancoDados
    conexao.conectaAutentWindows()
    if conexao.status != StatusExecucao.Sucesso:
        RegistraLinhaArquivo(nomeArqlogExec, conexao.mensagem, True)
        return (conexao.mensagem)

    if conexao.conectado:
        RegistraLinhaArquivo(nomeArqlogExec, f'A conexão com o SQL Server foi bem sucedida ao banco de dados {bancoDados}!', True)
    else:
        RegistraLinhaArquivo(nomeArqlogExec, f'Não foi possível conectar ao SQL Server. Banco de dados {bancoDados}', True)

    # --- Executa/Aplica o Script
    RegistraLinhaArquivo(nomeArqlogExec, f'Preparando para executar/aplicar o script do arquivo: {caminho + nomeArqScript}...', True)

    conexao.executaArquivoScript(caminho, arquivo)
    if conexao.status != StatusExecucao.Sucesso:
        RegistraLinhaArquivo(nomeArqlogExec, conexao.mensagem, True)
        return(conexao.mensagem)

    RegistraLinhaArquivo(nomeArqlogExec,f'O arquivo de script: {caminho + nomeArqScript} foi executado/aplicado com sucesso no banco de dados {bancoDados}',True)

    return 'Script Executado com Sucesso!'

# ************************************
#         INÍCIO DA ROTINA
# ************************************

# ----- Arquivo de LOG
caminho = os.getcwd() + "\ConfigArquivos\LogExecucao\ ".strip()

dt = dtm.now()
dtf = dt.strftime('%Y%m%d_%H%M%S')
arquivoLogExecucao = f'LogScript_{dtf}.log'

logExec = CriaArquivo(caminho, arquivoLogExecucao)
if not logExec:
    sys.exit(f'Arquivo de LOG: {caminho}{arquivo}')

# ----- Executa/Aplica o Script (Cria a Tabela no Banco de Dados Caso Não Exista)
caminho = os.getcwd() + "\ ".strip()
arquivo = "ScriptEstruturaBD.sql"
bancoDados = 'BD_ADRRBR'
mensagem = executaScript(caminho, arquivo, bancoDados, logExec)
sys.exit(mensagem)






