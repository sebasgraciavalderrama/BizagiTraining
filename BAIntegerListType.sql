IF NOT EXISTS (SELECT 1 FROM sys.types WHERE is_table_type = 1 AND name ='BAIntegerListType')
BEGIN
    CREATE TYPE BAIntegerListType AS TABLE
    (
        idValue INT NOT NULL,
        PRIMARY KEY CLUSTERED
        (
            idValue ASC
        )
        WITH (IGNORE_DUP_KEY = OFF)
    );
END;
