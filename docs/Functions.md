# [VBScript-CRUD](../README.md)
## Provided Functions

* Type conversion/checking functions
    * *string|null* **AsString**(*mixed* Value)  
        Converts $Value to string if its not Null or Empty. Else convert it to Null.
    * *bool* **IsString**(*mixed* Value)  
        Checks if $Value is a string.
    * *bool* **isIsoDate**(*string* s_input)  
        Checks the given date-string is in ISO 8601.
    * *Date* **CIsoDate**(*string* s_input)  
        Parses the given date-string from ISO 8601.
    * *string* **CIsoDate**(*Date* d_input)  
        Parses the given date to an ISO 8601 string.
    * *Date|null* **AsDate**(*mixed* Value)  
        Converts $Value to Date if its a valid date value. Else convert it to Null.
    * *bool|null* **AsBool**(*mixed* Value)  
        Converts $Value to Bool if its a valid boolean value. Else convert it to Null.
    * *bool* **IsBool**(*mixed* Value)  
        Checks if $Value is a boolean.
    * *byte|null* **AsByte**(*mixed* Value)  
        Converts $Value to byte if its a valid byte value. Else convert it to Null.
    * *bool* **IsByte**(*mixed* Value)  
        Checks if $Value is a byte.
    * *int|null* **AsInt**(*mixed* Value)  
        Converts $Value to integer if its a valid integer value. Else convert it to Null.
    * *bool* **IsInt**(*mixed* Value)  
        Checks if $Value is a integer.
    * *long|null* **AsLong**(*mixed* Value)  
        Converts $Value to long if its a valid long value. Else convert it to Null.
    * *bool* **IsLong**(*mixed* Value)  
        Checks if $Value is a long.
    * *double|null* **AsDouble**(*mixed* Value)  
        Converts $Value to double if its a valid double value. Else convert it to Null.
    * *bool* **IsDouble**(*mixed* Value)  
        Checks if $Value is a double.
    * *single|null* **AsSingle**(*mixed* Value)  
        Converts $Value to single if its a valid single value. Else convert it to Null.
    * *bool* **IsSingle**(*mixed* Value)  
        Checks if $Value is a single.
    * *currency|null* **AsCurrency**(*mixed* Value)  
        Converts $Value to currency if its a valid currency value. Else convert it to Null.
    * *bool* **IsCurrency**(*mixed* Value)  
        Checks if $Value is a currency.
    * *bool* **IsVoid**(*mixed* Value)  
        Checks if $Value has no data. Empty, Nothing and Null are values not considered as data.
* Array functions 
    * *array* **EmptyArray**(*int* Size)  
        Creates an empty array with $Size elements allocated.
    * *Scripting.Dictionary* **Dictionary**()  
        Creates an empty Dictionary.
* SQL inspired functions
    * *mixed* **IIF**(*bool* Condition, *mixed* ValidReturn, *mixed* InvalidReturn)  
        Functional-equivalent of ternary operator.

Notice that as this library has [VBScript-Reflect](https://github.com/the-linck/VBScript-Reflect) as an optional dependency, the functions provided there may also be avaliable.