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


