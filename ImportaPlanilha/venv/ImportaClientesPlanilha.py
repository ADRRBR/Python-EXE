from Classes.BAS_Generico import *
from Classes.BAS_Arquivo import *
from Classes.SIS_Conexao import clsConexaoBancoDados
from Classes.APL_Planilha import clsPlanilha
from Classes.APL_Clientes import clsClientes

# -------------------------------- ARQUIVOS DE LOG

caminho = os.getcwd() + "\ConfigArquivos\LogExecucao\ ".strip()

dt = dtm.now()
dtf = dt.strftime('%Y%m%d_%H%M%S')
arquivoLogExecucao = f'LogExecucao_{dtf}.log'
arquivoLogManutencao =  f'LogManutencao_{dtf}.log'

logExec = CriaArquivo(caminho, arquivoLogExecucao)
if not logExec:
    sys.exit(f'Arquivo de LOG: {caminho}{arquivoLogExecucao}')

logManut = CriaArquivo(caminho, arquivoLogManutencao)
if not logExec:
    sys.exit(f'Arquivo de LOG: {caminho}{arquivoLogManutencao}')

# -------------------------------- CONEXÃO COM O BANDO DE DADOS

RegistraLinhaArquivo(logExec,'Preparando o Acesso à Classe < Conexao >...', True)
RegistraLinhaArquivo(logExec,'Conectando ao banco de dados...', True)

conexao = clsConexaoBancoDados.ConexaoSQLServer()
conexao.conectaArquivoConfig('', '')
if conexao.status != StatusExecucao.Sucesso:
    RegistraLinhaArquivo(logExec, conexao.mensagem, True)
    sys.exit(conexao.mensagem)

if conexao.conectado:
    RegistraLinhaArquivo(logExec, 'A conexão com o SQL Server foi bem sucedida!', True)
else:
    RegistraLinhaArquivo(logExec, 'Não foi possível conectar ao SQL Server.', True)

listaData = conexao.dataListar
listaDataHora = conexao.datahoraListar
listaHora = conexao.horaListar

gravaData = conexao.dataGravar
gravaDataHora = conexao.datahoraGravar
gravaHora = conexao.horaGravar

listaValor = conexao.valorListar
gravaValor = conexao.valorGravar

# --------------------------------

RegistraLinhaArquivo(logExec, 'Preparando o Acesso à Classe < Clientes >...', True)

clientes = clsClientes.Clientes()
clientes.conexao = conexao

colunas = clientes.colClientes
codigo                = colunas.codigo.value - 1
nome                  = colunas.nome.value - 1
descricao             = colunas.descricao.value - 1
data_primeiro_contato = colunas.data_primeiro_contato.value - 1
valor_faturamento     = colunas.valor_faturamento.value - 1
data_renovacao        = colunas.data_renovacao.value - 1
hora_diaria_ligacao   = colunas.hora_diaria_ligacao.value - 1
valor_primeira_compra = colunas.valor_primeira_compra.value - 1

# --------------------------------

RegistraLinhaArquivo(logExec, 'Preparando o Acesso à Classe < Planilha >...', True)

RegistraLinhaArquivo(logExec,'Preparando para ler a planilha...', True)

planilha = clsPlanilha.Planilha()
tabPlan = planilha.arquivoConfigPlanilha('', '')
if planilha.status != StatusExecucao.Sucesso:
    RegistraLinhaArquivo(logExec, planilha.mensagem, True)
    sys.exit(planilha.mensagem)

TipoAtualizacao = conexao.tipoAtualizacaoBD.Incluir

for linha in range(len(tabPlan)):
     RegistraLinhaArquivo(logManut, f'Linha: {linha+1} - {tabPlan.loc[linha][codigo]}, ', True)
     RegistraLinhaArquivo(logManut, f'Linha: {linha+1} - {tabPlan.loc[linha][nome]}, ', True)
     RegistraLinhaArquivo(logManut, f'Linha: {linha+1} - {tabPlan.loc[linha][descricao]}, ', True)
     RegistraLinhaArquivo(logManut, f'Linha: {linha+1} - {listaDataHora(tabPlan.loc[linha][data_primeiro_contato])}, ', True)
     RegistraLinhaArquivo(logManut, f'Linha: {linha+1} - {listaValor(tabPlan.loc[linha][valor_faturamento])}, ', True)
     RegistraLinhaArquivo(logManut, f'Linha: {linha+1} - {listaData(tabPlan.loc[linha][data_renovacao])}, ', True)
     RegistraLinhaArquivo(logManut, f'Linha: {linha+1} - {listaHora(tabPlan.loc[linha][hora_diaria_ligacao])}, ', True)
     RegistraLinhaArquivo(logManut, f'Linha: {linha+1} - {listaValor(tabPlan.loc[linha][valor_primeira_compra])}, ', True)

     # -------------------------------- Inclusão
     clientes.prpClientes = {
         colunas.pk_cliente:            0,
         colunas.codigo:                tabPlan.loc[linha][codigo],
         colunas.nome:                  tabPlan.loc[linha][nome],
         colunas.descricao:             tabPlan.loc[linha][descricao],
         colunas.data_primeiro_contato: gravaDataHora(tabPlan.loc[linha][data_primeiro_contato]),
         colunas.valor_faturamento:     gravaValor(tabPlan.loc[linha][valor_faturamento]),
         colunas.data_renovacao:        gravaData(tabPlan.loc[linha][data_renovacao]),
         colunas.hora_diaria_ligacao:   gravaHora(tabPlan.loc[linha][hora_diaria_ligacao]),
         colunas.valor_primeira_compra: gravaValor(tabPlan.loc[linha][valor_primeira_compra])
     }

     clientes.prpClientes_atualizaBD(TipoAtualizacao)
     if clientes.status == StatusExecucao.Sucesso:
         RegistraLinhaArquivo(logManut, 'Registro incluído com sucesso.', True)
     else:
         RegistraLinhaArquivo(logManut, clientes.mensagem, True)

# --------------------------------
