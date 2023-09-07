from Classes.BAS_Generico import *

# Colunas da Tabela
class colClientes(Enum):
    pk_cliente = 0
    codigo = 1
    nome = 2
    descricao = 3
    data_primeiro_contato = 4
    valor_faturamento = 5
    data_renovacao = 6
    hora_diaria_ligacao = 7
    valor_primeira_compra = 8

class Clientes:
    def __init__(self):
        self.__colClientes = colClientes
        self.__prpClientes = []
        self.__lstClientes = []
        self.__JSONClientes = ' '
        self.__status = StatusExecucao.SemRequisicao
        self.__mensagem = ' '

    # Lista de Propriedades (Colunas da Tabela)
    @property
    def prpClientes(self):
        return self.__prpClientes

    @prpClientes.setter
    def prpClientes(self, valor):
        self.__prpClientes = [] # Limpa para ter Sempre um Elemento (Dicionário na Lista)
        self.__prpClientes.append(valor)

    # Objeto da Conexão
    @property
    def conexao(self):
        return self.__conexao

    @conexao.setter
    def conexao(self, valor):
        self.__conexao = valor

    # Registros Resultado de Consulta
    @property
    def lstClientes(self):
        return self.__lstClientes

    # Registros Resultado de Consulta (JSON)
    @property
    def JSONClientes(self):
        return self.__JSONClientes

    # Colunas da Tabela
    @property
    def colClientes(self):
        return self.__colClientes

    # Status
    @property
    def status(self):
        return self.__status

    # Mensagem
    @property
    def mensagem(self):
        return self.__mensagem

    # (existeRegistroChave) Verifica a Integridade Através da Chave Exclusiva da Tabela
    def existeRegistroChave(self, codigo):
        self.__LimpaStatus()

        if not self.__VerificaConexao():
            return

        sWhere = f"""
                  WHERE CODIGO = '{codigo}'
                  """

        self.consulta(sWhere, ' ')

        if self.__status == StatusExecucao.Encontrado:
            return True

        return False

    # (existeRegistroChavePK) Verifica a Integridade Através da Chave Exclusiva da Tabela Combinado com a Chave Primária
    def existeRegistroChavePK(self, pk_cliente, codigo):
        self.__LimpaStatus()

        if not self.__VerificaConexao():
            return

        sWhere = f"""
                  WHERE PK_CLIENTE =  {pk_cliente}
                  AND   CODIGO     = '{codigo}'
                  """

        self.consulta(sWhere, ' ')

        if self.__status == StatusExecucao.Encontrado:
            return True

        return False

    # (prpClientesExecuta) Método que Executa/APLICA uma Consulta de Todas as Colunas da Tabela, Conforme o Critério Informado
    def prpClientes_atualizaBD(self, tipoAtualizacaoBD):
        self.__LimpaStatus()

        if not self.__VerificaConexao():
            return

        v_pk_cliente            = self.__prpClientes[0][self.__colClientes.pk_cliente]
        v_codigo                = self.__prpClientes[0][self.__colClientes.codigo]
        v_nome                  = self.__prpClientes[0][self.__colClientes.nome]
        v_descricao             = self.__prpClientes[0][self.__colClientes.descricao]
        v_data_primeiro_contato = self.__prpClientes[0][self.__colClientes.data_primeiro_contato]
        v_valor_faturamento     = self.__prpClientes[0][self.__colClientes.valor_faturamento]
        v_data_renovacao        = self.__prpClientes[0][self.__colClientes.data_renovacao]
        v_hora_diaria_ligacao   = self.__prpClientes[0][self.__colClientes.hora_diaria_ligacao]
        v_valor_primeira_compra = self.__prpClientes[0][self.__colClientes.valor_primeira_compra]

        if tipoAtualizacaoBD == self.__conexao.tipoAtualizacaoBD.Incluir:
            # Verifica a Integridade Através da Chave Exclusiva da Tabela
            if self.existeRegistroChave(v_codigo):
                self.__status = StatusExecucao.Erro
                self.__mensagem = f'Já existe registro cadastrado com o código {v_codigo}.'
                return

            sComando = f"""INSERT INTO Tab_Clientes
                           (
                             codigo               
                            ,nome                 
                            ,descricao            
                            ,data_primeiro_contato
                            ,valor_faturamento
                            ,data_renovacao
                            ,hora_diaria_ligacao
                            ,valor_primeira_compra
                           )
                           VALUES
                           (
                             '{v_codigo}'               
                            ,'{v_nome}'                 
                            ,'{v_descricao}'            
                            ,'{v_data_primeiro_contato}'
                            , {v_valor_faturamento}
                            ,'{v_data_renovacao}'
                            ,'{v_hora_diaria_ligacao}'
                            , {v_valor_primeira_compra}
                           )
                        """
            sMensagem = 'Registro inserido com sucesso!'

        elif tipoAtualizacaoBD == self.__conexao.tipoAtualizacaoBD.Alterar:
            # Verifica a Integridade Através da Chave Exclusiva da Tabela Combinado com a Chave Primária
            if not self.existeRegistroChavePK(v_pk_cliente, v_codigo):
                self.__status = StatusExecucao.Erro
                self.__mensagem = f'O registro com o código {v_codigo} - PK: {v_pk_cliente} não foi localizado para alteração.'
                return

            sComando = f"""UPDATE Tab_Clientes
                           SET    nome                   = '{v_nome}'
                                 ,descricao              = '{v_descricao}'
                                 ,data_primeiro_contato  = '{v_data_primeiro_contato}'
                                 ,valor_faturamento      =  {v_valor_faturamento}
                                 ,data_renovacao         = '{v_data_renovacao}'
                                 ,hora_diaria_ligacao    = '{v_hora_diaria_ligacao}'
                                 ,valor_primeira_compra  =  {v_valor_primeira_compra}
                           WHERE  pk_cliente             =  {v_pk_cliente}
                           AND    codigo                 = '{v_codigo}'
                        """
            sMensagem = 'Registro alterado com sucesso!'

        elif tipoAtualizacaoBD == self.__conexao.tipoAtualizacaoBD.Excluir:
            # Verifica a Integridade Através da Chave Exclusiva da Tabela Combinado com a Chave Primária
            if not self.existeRegistroChavePK(v_pk_cliente, v_codigo):
                self.__status = StatusExecucao.Erro
                self.__mensagem = f'O registro com o código {v_codigo} - PK: {v_pk_cliente} não foi localizado para exclusão.'
                return

            sComando = f"""DELETE 
                           FROM   Tab_Clientes
                           WHERE  pk_cliente   =  {v_pk_cliente}
                           AND    codigo       = '{v_codigo}'  
                        """
            sMensagem = 'Registro excluído com sucesso!'

        else:
            self.__status = StatusExecucao.Erro
            self.__mensagem = 'Tipo de atualização inválido.'
            return

        self.__LimpaStatus()

        self.__conexao.executaSQL(sComando)

        self.__status = self.__conexao.status
        self.__mensagem = self.__conexao.mensagem

        if self.__status != StatusExecucao.Sucesso:
            return

        self.__mensagem = sMensagem

    # (consulta) Método que Efetua uma Consulta de Todas as Colunas da Tabela, Conforme o Critério Informado
    def consulta(self, SQLWhere, SQLOrderBY):
        self.__LimpaStatus()

        if not self.__VerificaConexao():
            return

        if not self.__VerificaParametrosConsulta(SQLWhere):
            return

        sSQL = f"""SELECT A.*
                   FROM   TAB_CLIENTES A 
                """
        sSQL += SQLWhere + ' '
        sSQL += SQLOrderBY + ' '

        tabSQL = self.__conexao.consultaSQL(sSQL)

        self.__status = self.__conexao.status
        self.__mensagem = self.__conexao.mensagem

        # Inicia o Retorno JSON
        clientesJSON = []
        retornoJSON = []

        if self.__status != StatusExecucao.Encontrado:
            # Retorno JSON
            retornoJSON.append({'Mensagem:': self.__mensagem,
                                'Quantidade:': 0
                              })
            retornoJSON.append({'Registros:': clientesJSON})
            self.__JSONClientes = json.dumps(retornoJSON, cls=NpEncoder, indent=4)
            return

        listaData = self.__conexao.dataListar
        listaDataHora = self.__conexao.datahoraListar
        listaHora = self.__conexao.horaListar

        gravaData = self.__conexao.dataGravar
        gravaDataHora = self.__conexao.datahoraGravar
        gravaHora = self.__conexao.horaGravar

        listaValor = self.__conexao.valorListar
        gravaValor = self.__conexao.valorGravar

        for indice in range(len(tabSQL)):
            valClientes = {
                self.__colClientes.pk_cliente:            tabSQL.loc[indice, 'pk_cliente'],
                self.__colClientes.codigo:                tabSQL.loc[indice, 'codigo'],
                self.__colClientes.nome:                  tabSQL.loc[indice, 'nome'],
                self.__colClientes.descricao:             tabSQL.loc[indice, 'descricao'],
                self.__colClientes.data_primeiro_contato: tabSQL.loc[indice, 'data_primeiro_contato'],
                self.__colClientes.valor_faturamento:     tabSQL.loc[indice, 'valor_faturamento'],
                self.__colClientes.data_renovacao:        tabSQL.loc[indice, 'data_renovacao'],
                self.__colClientes.hora_diaria_ligacao:   tabSQL.loc[indice, 'hora_diaria_ligacao'],
                self.__colClientes.valor_primeira_compra: tabSQL.loc[indice, 'valor_primeira_compra']
            }
            self.__lstClientes.append(valClientes)

            dataAux = dtm.strptime(tabSQL.loc[indice, 'data_renovacao'], '%Y-%m-%d').date()
            horaAux = tabSQL.loc[indice, 'hora_diaria_ligacao'].split('.')
            horaAux = horaAux[0]

            # Acrescenta o Retorno JSON
            valClientesJSON = {
                'pk_cliente':            tabSQL.loc[indice, 'pk_cliente'],
                'codigo':                tabSQL.loc[indice, 'codigo'],
                'nome':                  tabSQL.loc[indice, 'nome'],
                'descricao':             tabSQL.loc[indice, 'descricao'],
                'data_primeiro_contato': listaDataHora(tabSQL.loc[indice, 'data_primeiro_contato']),
                'valor_faturamento':     listaValor(tabSQL.loc[indice, 'valor_faturamento']),
                'data_renovacao':        listaData(dataAux),
                'hora_diaria_ligacao':   listaHora(horaAux),
                'valor_primeira_compra': listaValor(tabSQL.loc[indice, 'valor_primeira_compra'])
            }
            clientesJSON.append(valClientesJSON)

        # Retorno JSON
        retornoJSON.append({'Mensagem:': 'Consulta efetuada com sucesso',
                            'Quantidade:': len(tabSQL)
                          })
        retornoJSON.append({'Registros:': clientesJSON})
        self.__JSONClientes = json.dumps(retornoJSON, cls=NpEncoder, indent=4)

    # Métodos Privados
    def __VerificaConexao(self):
        try:
            if not self.__conexao.conectado:
                self.__status = StatusExecucao.Erro
                self.__mensagem = 'O objeto não está conectado ao banco de dados'
                return False
            return True
        except AttributeError as erro:
            self.__status = StatusExecucao.Erro
            self.__mensagem = ('Informe a propriedade < .conexao > para conexão com o banco de dados.')
        except Exception as erro:
            self.__status = StatusExecucao.Erro
            self.__mensagem = (f'Ocorreu o erro {erro}.' + ' Ao tentar conexão com o banco de dados.')
        return False

    def __VerificaParametrosConsulta(self, CondicaoSQLWhere):
        if CondicaoSQLWhere == ' ':
            self.__status = StatusExecucao.Erro
            self.__mensagem = 'Informe a Condição SQL Where'
            return False
        return True

    def __LimpaStatus(self):
        self.__Status = StatusExecucao.SemRequisicao
        self.__mensagem = ' '




