if not exists (select * from sys.databases s where s.name = 'ContactManager')
begin
	create database ContactManager
end
go

use ContactManager
go

if exists (select * from INFORMATION_SCHEMA.TABLES t where t.TABLE_NAME = 'contact')
begin
	drop table contact
end
go

if exists (select * from INFORMATION_SCHEMA.TABLES t where t.TABLE_NAME = 'addresses')
begin
	drop table addresses
end
go



create table dbo.addresses(
iAddressId int not null identity primary key,
vchStreet varchar(200) not null,
vchCity varchar(50) not null,
chState char(2) not null,
vchPostalCode varchar(10) not null
)
go

create table dbo.contact(
iContactId int not null identity primary key,
iAddressId int null foreign key references addresses(iAddressId),
vchFirstName varchar(100) not null,
vchLastName varchar(100) not null
)
go
--insert data into address table
insert addresses(vchStreet,vchCity,chState,vchPostalCode)
select m.street, m.city,m.state, m.zip
from Mock_Data m
--insert data into contacts table 
insert contact(vchFirstName,vchLastName)
select  m.first_name, m.last_Name
from Mock_data m

declare @i int = 1
while @i<1001
begin
	update contact 
	set iAddressId = @i
	from contact c
	where c.iContactId = @i  
	select @i += 1
end
alter table contact alter column iAddressId int not null 
select *from MOCK_DATA
select * from addresses
select * from contact



	