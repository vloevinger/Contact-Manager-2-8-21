create or alter procedure dbo.ContactUpdate(
	@iContactId int output,
	@iAddressId int,
	@vchFirstName varchar(100),
	@vchLastName varchar(100),
	@vchStreet varchar(200),
	@vchCity varchar(50),
	@chState char(2),
	@vchPostalCode varchar(10),
    @vchMessage varchar(1000) output
)
as
begin
	declare @iReturn int = 0
	select @iContactId = isnull(@iContactId,0), @iAddressId = isnull(@iAddressId,0), @vchFirstName = isnull(@vchFirstName,''),@vchLastName = ISNULL(@vchLastName,''),@vchMessage = ''

	if @iContactId = 0
	
	begin
		insert addresses(vchStreet,vchCity,chState,vchPostalCode)
		values(@vchStreet,@vchCity,@chState,@vchPostalCode);
		select @iAddressId = SCOPE_IDENTITY()
		
		insert contact(iAddressId,vchFirstName,vchLastName)
		values(@iAddressId,@vchFirstName,@vchLastName)
		select @iContactId = SCOPE_IDENTITY()
	end
	else 
	begin
		update addresses
		set vchStreet = @vchStreet, vchCity =@vchCity, chState =@chState, vchPostalCode = @vchPostalCode
		where iAddressId = @iAddressId
		
		update contact
		set iAddressId = @iAddressId, vchFirstName= @vchFirstName, vchLastName = @vchLastName
		where iContactId = @iContactId
	end
	return @iReturn
end
go
grant execute on ContactUpdate to public
go