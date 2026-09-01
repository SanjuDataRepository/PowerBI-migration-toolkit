WITH Legacy AS (
    SELECT
        TABLE_SCHEMA,
        TABLE_NAME,
        ORDINAL_POSITION,
        COLUMN_NAME,
        DATA_TYPE
    FROM INFORMATION_SCHEMA.COLUMNS
    WHERE TABLE_SCHEMA = 'LegacySchema' --Replace with your schema
      AND TABLE_NAME = 'LegacyTableName' --Replace with your table
),
Target AS (
    SELECT
        TABLE_SCHEMA,
        TABLE_NAME,
        ORDINAL_POSITION,
        COLUMN_NAME,
        DATA_TYPE
    FROM INFORMATION_SCHEMA.COLUMNS
    WHERE TABLE_SCHEMA = 'TargetSchema' --Replace with your schema
      AND TABLE_NAME = 'TargetTableName' --Replace with your table
)
SELECT
    COALESCE(l.ORDINAL_POSITION, t.ORDINAL_POSITION) AS [Position],

    CONCAT(l.TABLE_SCHEMA, '.', l.TABLE_NAME) AS [Legacy Table],
    CONCAT(t.TABLE_SCHEMA, '.', t.TABLE_NAME) AS [Target Table],

    l.COLUMN_NAME AS [Legacy Column Name],
    l.DATA_TYPE AS [Legacy Data Type],

    t.COLUMN_NAME AS [Target Column Name],
    t.DATA_TYPE AS [Target Data Type],

    CASE
        WHEN l.COLUMN_NAME IS NULL THEN 'Missing in Legacy'
        WHEN t.COLUMN_NAME IS NULL THEN 'Missing in Target'
        WHEN l.DATA_TYPE = t.DATA_TYPE THEN 'Type Match'
        ELSE 'Type Different'
    END AS [Comparison Flag]

FROM Legacy l
FULL OUTER JOIN Target t
    ON l.ORDINAL_POSITION = t.ORDINAL_POSITION

ORDER BY
    [Position];
