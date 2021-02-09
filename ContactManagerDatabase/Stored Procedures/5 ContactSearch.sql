create or alter procedure  dbo.ContactSearch(
@vchSearchCriteria varchar(1000)='',
@vchMessage varchar(1000) output
)
as
begin
declare @iReturn int = 0
	select 	@vchSearchCriteria = ISNULL(@vchSearchCriteria,''),
	@vchMessage =''
	
	declare @t table(iContactId int)

	insert @t (iContactId)
	select c.iContactId 
	from contact c
	join addresses a
	on a.iAddressId = c.iAddressId
	where 
	c.vchFirstName like '%' + @vchSearchCriteria+'%'
	or c.vchLastName like '%' + @vchSearchCriteria +'%'
	or a.vchStreet like  '%' + @vchSearchCriteria+'%'
	or a.vchCity like '%' + @vchSearchCriteria+'%'
	or a.chState like '%' + @vchSearchCriteria +'%'
	or a.vchPostalCode like '%' + @vchSearchCriteria + '%'


	Select 
	c.iContactId,c.vchFirstName, c.vchLastName,c.iAddressId,a.vchStreet, a.vchCity, a.chState, a.vchPostalCode
	from contact c
	join @t t
	on t.iContactId = c.iContactId
	join addresses a
	on a.iAddressId = c.iAddressId
	order by c.vchLastName, c.vchFirstName
	return @ireturn
end
go
grant execute to public