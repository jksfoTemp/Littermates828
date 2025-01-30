select table_schema as database_name,
    table_name
from information_schema.tables
-- where table_type = 'BASE TABLE'
--        and table_schema = database() 
-- order by database_name, table_name;userz