# Contact Manager 2-8-21
 

Contact Manager is a simple contact management app created using ASP.Net MVC and a SQL Server database.


## Database
The database consists of two tables: 
  - addresses
  - contact
 ### Database Set up 
 To set up the database I downloaded Mock Data and imported it to a dbo.Mock_Data table and then used that data to populate my tables.
 Once you import the Mock Data to a database, you can use the database files included to create a Contact Management Database.
 The files are numbered in the order that they should be rolled out.
  
  
## .NET Framework

### CPU Framework
This is a framework I developed while taking a full stack development course.
It contains all the business logic needed to get the data from the database to the front end.

### Contact Manager Biz Objects
This class library inherits from the Biz Objects Class in CPU Framework.  In it I create a custom class called BizContact which can retrieve all the data from the contact and addresses from the database. I use this as the Model in the MVC framework.

### Contact Manager CSharp
This project has contains the Views and Controllers for the Contact Manager Web App.
- Controllers: 
  -Home Controller
    - 3 methods and views:
      - view: allows the user to view a list of contacts. Also allows user to search for criteria.
      - edit: allows the user to update or create a new contact.
      - delete: allows the user to delete a contact

In order to work properly you must update the connection string to the server and database you are using in your computer. 
To do this, navigate to the web.config folder and update the server and database in the connection string.

 <connectionStrings>
    <add name="conn" connectionString="Server=DESKTOP-FA1UCE3;Database=ContactManager;Trusted_Connection=True;" />
  </connectionStrings>
