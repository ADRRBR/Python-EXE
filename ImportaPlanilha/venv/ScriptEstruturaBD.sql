
IF OBJECT_ID ('dbo.TAB_Clientes') IS NOT NULL 
    RETURN

CREATE TABLE dbo.TAB_Clientes
(
  pk_cliente              int          identity(1,1) NOT NULL 
 ,codigo                  varchar(20)                NOT NULL
 ,nome                    varchar(100)               NOT NULL
 ,descricao               text
 ,data_primeiro_contato   datetime
 ,valor_faturamento	      float
 ,data_renovacao	      date
 ,hora_diaria_ligacao	  time
 ,valor_primeira_compra	  float
)

alter table dbo.TAB_Clientes add constraint PK_Tab_Clientes primary key(pk_cliente);

create unique index IND_TAB_Clientes_1 ON dbo.TAB_Clientes(codigo)
create unique index IND_TAB_Clientes_2 ON dbo.TAB_Clientes(nome)

