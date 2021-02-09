create or alter procedure dbo.ContactGet(
	@iContactId int = 0,
	@bAll bit = 0,
	@vchMessage varchar(1000) output
	)
as
begin
	declare @iReturn int = 0
	select @iContactId = ISNULL(@iContactId,0), @bAll =isnull(@bAll,0),@vchMessage = ''
	
	declare @t table(iContactId int)

	if @iContactId <> 0
	begin
		insert @t (iContactId)
		values (@iContactId)
	end
	else if @bAll = 1
	begin
		insert @t (iContactId)
		select c.iContactId
		from contact c
	end
	

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
/*grant execute on ContactGet to public
declare @icontacttId int = 0,@bAll int = 1, @iReturn int, @vchMessage varchar(1000)
exec @iReturn = ContactGet @iContactId = @icontacttId, @bAll = 1,@vchMessage = @vchMessage output*/
