# VBScript-CRUD

This library gives an ASP Classic implementation of Object Oriented access to Databases.

The implentation is based in classes that abstracts the SQL logic out of your code, letting you deal only with classes representing your database's tables - wich it's far easier and allows better portability of code to modern languages.

# Requirements

This library *needs* [VBScript-Reflect](https://github.com/the-linck/VBScript-Reflect) to provide Object Oriented Database capability - wich is the focus of the project. Without it, the library provides several utilitary functions, but the database utilities related to Entites will not work as expected.

Also, as VBScript-Reflect, this library optionaly depends on [ASPJson](https://github.com/rcdmk/aspJSON), to provide JSON Export capabilies.

# Project Structure

There are 6 code files on the project - but just 2 you need to care about.

* **[Database.asp](docs/Database.md)**  
    Main database structure, wich encapsulates the access to the Connection, the SQL logic and Data to Entities conversion.  
    Already includes the files needed to work and provides a default instance named *DB_Instance*.
* **[ADOConstants.asp](docs/ADOConstants.md)**  
    All constants used by ADODB to configurate database access.
* **[Functions.asp](docs/Functions.md)**  
    Functions used by the library and made avaliable to the user. Some aditional utilitary functions are provided too.
* **[Conditions.asp](docs/Conditions.md)**  
    Encapsulates conditions for Where and Join clauses.
* [**Statement.asp**](docs/Statement.md)  
    Encapsulates the clauses of SQL logic, letting you build statements with commands in any order and have a standard SQL output.
* [**DB_Entity.asp**](docs/DB_Entity.md)  
    Extends VBScript-Reflect's *_Entity.asp*, adding properties and methods used to operante in database with Entities.


## Why not whole evertything in classes?

Notice that we use 3 files with structures (*Database*, *Statement*, and *Conditions*), that are actually plain VBScript "classes", not using *VBScript-Reflect*'s extension include or even the one we provide. But why?

Native structures, without a extension include - and it's variables, properties and methods - are obviously far faster both in compilation and execution.
Also, this structs needn't any complex functionality, like reflection or static properties, so native implementation is all we need for the job.

So keep this performance tip: **restrict the use of the *DB_Entity.asp* extension include to the classes used to access your database**.

# Deconstructed SQL

A great resource of this library, that not only comes along OO Database but also provides it, is Deconstructed SQL: being able to use SQL clauses without having to care about their order and not needing to provide them in a single sentence.

This allows a very flexible use of the library, letting you make statements in any way you want - ending with Command objects feeded with standard and valid SQL and already parametized to prevent SQL Injection.


As said before, this resource is the base of the Database access through Entities, because it really makes thing quite more straightfoward.
Deconstructed SQL functionality is provided by instances of the **Statement** structure, wich has [it's own documentation](docs/Statement.md).



# DB_Entity: the upgraded _Entity.asp include

This file is meant to be included in your ASP Classes along with VBScript-Reflect's _Entity.asp, to extend it's capabilities providing bidirectional encapsulated access for the Database structure to the Entity fields, properties and methods - in short, the CRUD interface.

There's also two new methods for the optional JSON export freature - wich obviousy depends on ASPJson library.

The provided properties and methods are descripted on [DB_Entity own documentation](docs/DB_Entity.md).



## Static_Initialize: Declaring database-table properties

In addition to the new properties/methods provided for classes by the new include, you must also provide some metadata for the library know how to translate your Entites in SQL when accessing the database.

The *Static_Initialize()* method, wich is optional in VBScript-Reflect, fits live a glove for this purpose. Actually **it is essential** to do this task just once without having to deal with a initialization-lock logic that is already implemented on this lib - no need to reinvent the steel here.

You have to specify wich table this Entity represents on the database, also each field with name and data-type pretty much as done in VBScript-Reflect, but using ADO datatype contants (wich come shipped with this library).



### Implementation Example

Here's an example of how to do this on an dummy test page:

```ASP
<!--#include file="VBScript-Reflect/_Class.asp"-->
<!--#include file="VBScript-CRUD/Database.asp"-->
<%
Class PostTag
    %>
    <!--#include file="VBScript-Reflect/_Entity.asp"-->
    <!--#include file="VBScript-CRUD/DB_Entity.asp"-->
    <%
    Public ID
    Public Slug
    Public Name
    Public Description

    Sub Static_Initialize
        Self.Fields("TableName") = "post_tags"

        Self.Fields("ID") = adUnsignedInt
        Self.Fields("Slug") = adVarWChar
        Self.Fields("Name") = adVarWChar
        Self.Fields("Description") = adVarWChar
    End Sub

End Class
%>
```

If you want to setup default values for the Entity, just add an *Instance_Initialize()* method and set the fields there, so it will act as a parameter-less constructor.

### Foreign Entities

To go a step further, you can specify Foreign Entities on your entity. Those are  a direct derivation of foreign key relationship, that makes VBScript-CRUD automatically load on to a Entity all other entites marked as linked to them.

To enable this freature, a Dictionary must be stored on the *Foreign* static field. This dictionary's keys keep the fields of the Entity that will be used to store other entites, and the dictionary's values keep the class-name of this entities.

An example do clarify things:

```ASP
<%
Class Note
    %>
    <!--#include file="VBScript-Reflect/_Entity.asp"-->
    <!--#include file="VBScript-CRUD/DB_Entity.asp"-->
    <%
    Public ID
    Public Created
    Public AuthorID
    Public Title
    Public Content

    Public Author

    Sub Static_Initialize
        Self.Fields("TableName") = "post_tags"

        Self.Fields("ID") = adBigInt
        Self.Fields("Created") = adDate
        Self.Fields("AuthorID") = adUnsignedInt
        Self.Fields("Title") = adVarWChar
        Self.Fields("Content") = adLongVarWChar

        Self.Fields("Foreign") = Dictionary()
        Self.Fields("Foreign")("Author") = "User"
    End Sub

End Class
%>
```

With the above code you declare a **Note** entity, that registers **User** class as a Foreign Entity to be read to *Author* field. Quite simple, isn't?

There's no need to specify primary or foreign keys to be used in this relation - they will be automatically verified by the library, both entities must only share a common field, and this field must be the key-field of one of them.  
If no field-name is stored on *KeyField*static field, like on the previous example, the first field registered on the class is taken as the key-field of the Entity.
