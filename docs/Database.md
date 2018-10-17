# [VBScript-CRUD](../README.md)
## Database Structure

This structure encapsulates the [Statement](Statement.md) Logic, extending it's functionality, also providing the logic to work with Entities.

All public methods and properties of Statement stuct are provided with bridges, now returning a Database instance to chained calls. The following new properties and methods are also providaded:

* Public interface
    * *ADODB.Connection* **Connection**  
        Standard ADO Connection to use.
    * *boolean* **MySQL_Date_Patch**  
        If the data-type correction for MySQL must be used.  
        (ADO have trouble to read DATE from MySQL, so using this provider requires the Patch)
    * *boolean* **UseFowardOnly**  
        If FowardOnly recordsets must be used in queries. Used internally for performance on Entity Read, but also avaliable to users.
    * *ADODB.Recordset* **ForwardOnlyRecordset**()  
        Creates a Recordset disconnected from database, allowing to use DB data without an active connection to it.  
        Uses Foward Only cursor type to maximize performance.
    * *ADODB.Recordset* **StaticRecordset**()  
        Creates a Recordset disconnected from database, allowing to use DB data without an active connection to it.  
        Uses static cursor type to allow moving in any direction.
    * *Scripting.Dictionary* **EntityFields**(*Object* Entity)  
        Recovers all the fields of given $Entity.
    * *Scripting.Dictionary* **EntityKeys**(*Object* Entity)  
        Recovers the fields of given $Entity registered as keys.
    * *array* **ParseEntities**(*Object* Entity *ADODB.Recordset* Recordset)  
        Converts $Recordset to an array of $Entity objects.  
        Kept public to allow use of custom queries for reading Entities from Database, allowing optimization outside the library.
* Connection-related functions
    * *self* **Connect**()  
        Connects to the database (if not already connected) and increment the connection counter.
    * *self* **Disconnect**()  
        Decrement the  connection counter, disconects from the database (if connected) when the counter reaches 0.
* Condition-related clauses
    * *self* **Where_Entity**(*Object* Entity)  
        Adds conditions on WHERE clause based on given $Entity.
* SQL Statement Execution
    * *int* **Run_Insert**()  
        Assembles the clauses from current statement in an INSERT statement.
    * *ADODB.Recordset* **Run_Select**()  
        Assembles the clauses from current statement in a SELECT statement.
    * *int* **Run_Update**()  
        Assembles the clauses from current statement in an UPDATE statement.
    * *int* **Run_Delete**()  
        Assembles the clauses from current statement in an DELETE statement.
* Generic Entity CRUD
    * *int* **Create**(*Object* Entity)  
        Creates $Entity's register on it's Database table.
    * *array&lt;Entity&gt;* **Read**(*Object* Entity)  
        Read registers from $Entity's table on Database using $Entity to filter records.
    * *int* **Update**(*Object* Entity)  
        Updates the $Entity's register on it's Database table.
    * *int* **Delete**(*Object* Entity)  
        Deletes the $Entity's register from it's Database table.

***Notice:** All methods related to entities require [VBScript-Reflect](https://github.com/the-linck/VBScript-Reflect) to be included on your project.*