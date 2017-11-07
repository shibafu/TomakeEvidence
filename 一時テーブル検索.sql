SELECT   T.name AS TableName,
          C.name AS ColumnName,
		  TY.name   AS ColumnType,
		  C.max_length AS max_length
FROM     tempdb.sys.objects AS T
            INNER JOIN tempdb.sys.columns AS C
             ON T.object_id = C.object_id
			INNER JOIN tempdb.sys.types AS TY
			ON C.system_type_id = TY.system_type_id
WHERE    T.name LIKE '%W30%'
ORDER BY T.name,
          C.name;