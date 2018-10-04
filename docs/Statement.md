# [VBScript-CRUD](../README.md)
## Statement Structure

The *Statement* structure provides Deconstructed SQL functionality through the following public properties and methods:

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