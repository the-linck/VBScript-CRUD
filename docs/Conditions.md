# [VBScript-CRUD](../README.md)
## DB_Condition Structure

A simple structure used by the library to store a basic SQL Condition. You may use it outside the lib too, but is important to know how it is used internally to keep the standard.

All fields an methods of DB_Condition are public:

* *string* **Operator**  
    Local operator to insert before this condition in a conditions list.
* *string* **Field**  
    Field to check value(s) or simple fixed condition.  
    Wich one of this two depends on the presence of $Values in this struct.
* *mixed* **Values**  
    Value(s) to check against $Field.  
    For IN clauses will be an array
* *self* **Construct**(*string* Operator_, *string* Field_, *mixed* Values_)
    Real optional Constructor.

### How it works

When a DB_Condition is created, you may use the optional constructor to feed all struct fields at once. The *Field* field has two possible uses:

* hold a field name
* hold a simple textual condition

How this field will be used depends on the presence of *Values* field: if it's Empty, *Field* will be taken as a simple textual condition - behaviour used by Statement.Where_Clause() -, else *Field* will be intepreted as a value in an **equal as** comparision (for scalar values) or a list of values for a **IN** comparision (for array of values) - behaviour used by Statement.Where_In().  
***Object values are not allowed***

Finally, when Statement recovers this struct from it's storage to build the resulting command, *Operator* field is inserted before this condition on the current condition list - if this is not the first condition of the list, of course.

## DB_JoinClause Structure

This structure holds **n** DB_Condition instances to a join clause on Statements, allowing queries with complex join to be easily implemented.  
As DB_Condition, you may use this struct on your code outside the lib, but this one is far more specific to SQL utilities - wich make not useful for most othe purposes. But if you wish to use it, just understand how it is used to keep the pattern.

The following fields and methods are public:

* *Dictionary&lt;int, Condition&gt;* **Conditions**  
    Basic DB_Condition container.
* *string* **JoinType**  
    Join operator to insert before this join in a join-list.
* *string* **Table**  
    Table to join.
* *self* **Construct**(*string* JoinType_, *string* Table_)  
    Real optional Constructor.

When a instance is created, you may use the default constructor to pass the desired Join Type and the Table to join, but the conditions must be supplyed directly on the *Conditions* field - wich is an exposed dictionary.

Internally, the lib uses int keys on *Conditions* - adding each new condition to a key equal to Conditions.Count -, but the type of the keys is completely irrelevant as they are just ignored.  
The trick here is making the dictionary work as a list, adding elements without creating custom logic for memory (re)allocation.