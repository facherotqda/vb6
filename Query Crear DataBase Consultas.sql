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


