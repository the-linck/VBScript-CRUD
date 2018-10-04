# [VBScript-CRUD](../README.md)
## DB_Entity include

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

### Implementation example

To use this include, it you must add the file to the class along with *_Entity* from VBScript-Reflect, as the following:

```ASP
<%
Class Char_Attribute
    %>
    <!--#include file="VBScript-Reflect/_Entity.asp"-->
    <!--#include file="VBScript-CRUD/DB_Entity.asp"-->
    <%
    Public IDChar
    Public Name
    Public Value
    Public Description

    Sub Static_Initialize
        Self.Fields("TableName") = "char_attributes"

        Self.Fields("IDChar") = adUnsignedInt
        Self.Fields("Name") = adVarWChar
        Self.Fields("Value") = adUnsignedTinyInt
        Self.Fields("Description") = adVarWChar
    End Sub

End Class
%>
```

### Foreign Entities

To enable Foreign Entities, a Dictionary must be stored on the *Foreign* static field. This dictionary's keys keep the fields of the Entity that will be used to store other entites, and the dictionary's values keep the class-name of this entities.


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

No primary or foreign key needs to be specified- they will be automatically verified by the library. The entities must only share a common field, that is the key-field of one of them.

If no field-name is stored on *KeyField* static field, the first field registered in Static_Initialize() is taken as the key-field of the Entity.
