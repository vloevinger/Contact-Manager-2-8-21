create or alter procedure dbo.ContactDelete(@iContactId int, @vchMessage varchar(1000) output)
as
begin
	declare @iReturn int = 0
	
	select @iContactId = ISNULL(@iContactId,0), @vchMessage = ''
	
	if not exists (select * from contact c where c.iContactId = @iContactId)
	begin
		select @iReturn = 1, @vchMessage = 'Specified Contact Id does not exist'
		goto failed_validation
	end
	
	
	delete contact where iContactId = @iContactId
	delete addresses where iAddressId = (select iAddressId from contact where iContactId = @iContactId) 	

	failed_validation:
	return @iReturn
end
go
grant execute on ContactDelete to public