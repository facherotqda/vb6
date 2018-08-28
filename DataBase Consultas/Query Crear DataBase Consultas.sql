create database Consultas

use Consultas

create table Clientes
(
[CÓDIGO CLIENTE] int primary key not null identity(1,1),
empresa varchar(50),
dirección varchar(50),
población varchar(50),
teléfono varchar(50),
responsable varchar(50)
)


drop table Clientes 

--creando una restriccion(constraint) para columna codigo cliente
alter table Clientes	
add constraint PK_NOREPEAT
unique ([CÓDIGO CLIENTE] ) 

--quito primero la restriccion
alter table clientes drop constraint PK_NOREPEAT;
--quito regristo de restriccion
alter table clientes drop constraint PK__Clientes__14F572BB03317E3D

--modifico una vez sacada la restriccion
alter table clientes alter column [CÓDIGO CLIENTE] varchar(50) not null 

--agrego Restriccion Unique para que no se repita La PRIMARY KEY
alter table clientes add constraint PK_UNIQUE unique ([CÓDIGO CLIENTE] ) 

