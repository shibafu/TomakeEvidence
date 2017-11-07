SELECT   T.name AS TableName,
          C.name AS ColumnName
FROM     sys.tables AS T
            INNER JOIN sys.columns AS C
             ON T.object_id = C.object_id
WHERE    T.name LIKE '%W30%'
ORDER BY T.name,
          C.name;