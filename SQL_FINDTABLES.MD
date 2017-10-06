## The VBA program will take input from the user and call a function, passing their input as args


*findTables(userInput)


## The SQL takes input and utilizes it in the query:

*SELECT DISTINCT table_name, column_name from all_tab_columns where owner = 'LDRDBA' and column_name like '%&userInput%';

### This needs to be a bit more advance to take as many fields as they want and then build the appropriate amounts of and operands for the WHERE clause