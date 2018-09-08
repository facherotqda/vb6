use Consultas

create procedure sp_EliminarCliente
(

@codigo_cliente as varchar(50)
)
as
begin
SET NOCOUNT ON
delete from Clientes where [CÓDIGO CLIENTE]=@codigo_cliente

select 'el cliente a sido eliminado';
end

exec sp_EliminarCliente '14';