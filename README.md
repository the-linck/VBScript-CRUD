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