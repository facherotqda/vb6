use Consultas
go
alter procedure sp_AgregarCliente 
@cod_cliente as varchar(50),
@empresa as varchar(50),
@dirrecion as varchar(50),
@poblacion as varchar(50),
@telefono as varchar(50),
@responsable as varchar(50),
@msg as varchar(100) output

as
begin
SET NOCOUNT ON

 begin tran T_add
 
 begin try
	
	insert into dbo.Clientes ([CÓDIGO CLIENTE],dirección,empresa,población,responsable,teléfono)
	values(@cod_cliente,@empresa,@dirrecion,@poblacion,@telefono,@responsable)
	set @msg='El usuario se registro correctamente desde sql.'
	
	commit tran T_add
	 
 end try
 
 begin catch
 set @msg= 'Ocurrio un error: '+ERROR_MESSAGE() +' en la linea '+CONVERT(nvarchar(255),error_line() )+ '.'		 
	
	rollback tran T_add
 end catch
	 
end 
go

DECLARE @msg AS VARCHAR(100);
EXEC sp_AgregarCliente '1','hola Torres','clau@mail.com','a220109',1,'asdasd',@msg OUTPUT
SELECT @msg AS msg