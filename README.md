# VBScript-CRUD

This library gives an ASP Classic implementation of Object Oriented access to Databases.

The implentation is based in classes that abstracts the SQL logic out of your code, letting you deal only with classes representing your database's tables - wich it's far easier and allows better portability of code to modern languages.

# Requirements

This library *needs* [VBScript-Reflect](https://github.com/the-linck/VBScript-Reflect) to provide Object Oriented Database capability - wich is the focus of the project. Without it, the library provides several utilitary functions, but the database utilities related to Entites will not work as expected.  
_**Notice:** this core branch supposes that VBScript-Reflect folder will be at the same level of VBScript-CRUD - named exactly as the project._

Also, as VBScript-Reflect, this library optionaly depends on [ASPJson](https://github.com/rcdmk/aspJSON), to provide JSON Export capabilies.

# Project Structure

There are 6 code files on the project - but just 2 you need to care about.

* **Database.asp**  
    Main database class, wich encapsulates the access to the Connection, the SQL logic and Data to Entities conversion.  
    Already includes the files needed to work and provides a default instance named *DB_Instance*.
* **ADOConstants.asp**  
    All constants used by ADODB to configurate database access.
* **Functions.asp**  
    Functions used by the library and made avaliable to the user. Some aditional utilitary functions are provided too.
* **Conditions.asp**  
    Encapsulates conditions for Where and Join clauses.
* **Statement.asp**  
    Encapsulates the clauses of SQL logic, letting you build statements with commands in any order and have a standard SQL output.
* **DB_Entity.asp**  
    Extends VBScript-Reflect's *_Entity.asp*, adding properties and methods used to operante in database with Entities.

# Deconstructed SQL

A great resource of this library, that not only comes along OO Database but also provides it, is Deconstructed SQL: being able to use SQL clauses without having to care about their order and not needing to provide them in a single sentence.  
This allows a very flexible use of the library, letting you make statements in any way you want - ending with Command objects feeded with standard and valid SQL and already parametized to prevent SQL Injection.

As said before, this resource is the base of the Database access through Entities, because it really makes thing quite more straightfoward.

### Statement class

The *Statement* class provides Deconstructed SQL functionality by the following public properties and methods:

* *bool* **UseUTF**  
    If UTF8 charset must be used to deal with strings.
* *bool* **PrepareCommands**  
    If built SQL Statements are set as Prepared Statements.
* *self* **Clear**  
    Removes all caluses from this statement.
* Table-related clauses
    * *self* **From_Clause**(*string* Table)  
        Set the table for Select/Delete statements.
    * *self* **Into_Clause**(*string* Table)  
        Set the table for Insert statements.
    * *self* **Update_Clause**(*string* Table)  
        Set the table for Update statements.
* Join-related clauses
    * *self* **Join_Clause**(*string* JoinType, *string* Table, *string|array|Dictionary|DB_Condition* Conditions, *string* Operator)  
        Adds a JOIN clause of $JoinType for $Table, with $Conditions, using $Operator for them.
* Field-related clauses
    * *self* **Select_Clause**(*string|Array|Dictionary* Fields)  
        Stores the given $Fields in the fieldlist for Select statements.  
        To provide alisases, pass a dictionary in $Fields.
    * *self* **Set_Clause**(*string|Array|Dictionary* Fields)  
        Stores the given $Fields in the field/value list for Update statements.  
        To provide values to the fields pass a Dictionary to $Fields or they will be set to null.
    * *self* **Set_Field**(*string* Fields, *mixed* Value)  
        Stores $Field and $Value in the field/value list for Update statements.
    * *self* **Insert_Clause**(*string|Array|Dictionary* Fields)  
        Stores the given $Fields in the field/value list for Insert Update statements.  
        To provide values to the fields pass a Dictionary to $Fields or they will be set to null.
    * *self* **Insert_Field**(*string* Fields, *mixed* Value)  
        Stores $Field and $Value in the field/value list for Insert statements.
* Condition-related clauses
    * *self* **Where_Clause**(*string|array|Dictionary|DB_Condition* Conditions, *string* Operator)  
        Adds conditions on WHERE clause, using $Operator for them.
    * *self* **Where_In**(*string* Field, *string|array* Values, *string* Operator)  
        Adds a condition on WHERE clause to check if $Field is equal/in $Values, using $Operator for this condition.
* Order-related clauses
    * *self* **Order_Clause**(*string|array<string>* Fields, *string* Order)  
        Adds fields to ORDER BY clause.
* Group-related clauses
    * *self* **Group_Clause**(*string|array<string>* Fields)  
        Adds fields to GROUP BY clause.
* SQL Statement Assemble  
    Where the magic happens
    * *ADODB.Command* **Build_Insert**()  
        Assembles the clauses from this statement in an INSERT statement.  
        *After assembling the SQL, this DB_Statemet object is cleared.*
    * *ADODB.Command* **Build_Select**()  
        Assembles the clauses from this statement in an SELECT statement.  
        *After assembling the SQL, this DB_Statemet object is cleared.*
    * *ADODB.Command* **Build_Update**()  
        Assembles the clauses from this statement in an UPDATE statement.  
        *After assembling the SQL, this DB_Statemet object is cleared.*
    * *ADODB.Command* **Build_Delete**()  
        Assembles the clauses from this statement in an DELETE statement.  
        *After assembling the SQL, this DB_Statemet object is cleared.*


As you probabily noticed, most methods return the Statement object itself, allowing you to make chained calls on the object to build a SQL statement.  
The clauses names are not exactly equal as their SQL equivalent because some of them would collide with reserved words of the language, so they were kept standardized with sufixes.



# DB_Entity: the upgraded _Entity.asp include

This file is meant to be included in your ASP Classes along with VBScript-Reflect's _Entity.asp, to extend it's capabilities providing bidirectional encapsulated access for the Database class to the Entity fields, properties and methods - in short, the CRUD interface.

There's also two new methods for the optional JSON export freature - wich obviousy depends on ASPJson library.

The following properties and methods are provided by DB_Entity:

* Crud Interface
    * *string* KeyField  
        Gets the name of the field set to be used as primary key.  
        If none is set, returns the first registered field.
    * *int* Create()  
        Inserts this object on it's database table.
    * *array<self>* Read()  
        Querys this object's database table.
    * *int* Update()  
        Updates this object on it's database table.
    * *int* Delete()  
        Deletes this object from it's database table.
* Queryable Interface
    * *bool* Queryable  
        If this Entity is set to be used for queries (select statements).
    * *self* ToNonQueryable()  
        Marks this object to not be used for queries.
    * *self* ToQueryable()  
        Marks this object to be used in queries, setting all fields to empty.
* JSON export
    * *JSONobject* ToJSON()  
        Exports this Entity to a JSONobject, adding all registered Foreign entities to it.
    * *string* ToString()  
        Exports this Entity to a JSON string, adding all registered Foreign entities to it.
