create database Consultas

use Consultas

create table Clientes
(
[C�DIGO CLIENTE] int primary key not null identity(1,1),
empresa varchar(50),
direcci�n varchar(50),
poblaci�n varchar(50),
tel�fono varchar(50),
responsable varchar(50)
)


drop table Clientes 

--creando una restriccion(constraint) para columna codigo cliente
alter table Clientes	
add constraint PK_NOREPEAT
unique ([C�DIGO CLIENTE] ) 

--quito primero la restriccion
alter table clientes drop constraint PK_NOREPEAT;
--quito regristo de restriccion
alter table clientes drop constraint PK__Clientes__14F572BB03317E3D

--modifico una vez sacada la restriccion
alter table clientes alter column [C�DIGO CLIENTE] varchar(50) not null 

--agrego Restriccion Unique para que no se repita La PRIMARY KEY
alter table clientes add constraint PK_UNIQUE unique ([C�DIGO CLIENTE] ) 

