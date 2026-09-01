/*
===============================================================================
Power BI Query Metadata Extraction
===============================================================================

Purpose:
    Extract active and disabled Power Query metadata from a Power BI semantic
    model by using DAX Studio Dynamic Management Views.

Instructions:
    1. Open the PBIX file in Power BI Desktop.
    2. Connect DAX Studio to the open Power BI model.
    3. Run ONE query at a time.
    4. Export each result as a separate CSV file.
    5. Combine exported files and use as input to the Power BI Lineage Extractor.

Recommended output files:
    - active_queries.csv
    - disabled_queries.csv
Combine files

Important:
    Do not run both SELECT statements together. DAX Studio DMV queries should
    be executed individually.
===============================================================================
*/


/*
-------------------------------------------------------------------------------
Query 1: Active Queries
-------------------------------------------------------------------------------

Extracts partition metadata for Power Query queries that are loaded into the
semantic model.

Recommended export filename:
    active_queries.csv
*/

SELECT
    'Data Load Enabled' AS [ObjectType],
    [Name],
    [QueryDefinition]
FROM
    $SYSTEM.TMSCHEMA_PARTITIONS;


/*
-------------------------------------------------------------------------------
Query 2: Disabled Queries and Shared Expressions
-------------------------------------------------------------------------------

Extracts Power Query expressions that can include disabled-load queries,
shared expressions, parameters, functions, and staging queries.

Recommended export filename:
    disabled_queries.csv
*/

SELECT
    'Data Load Disabled' AS [ObjectType],
    [Name],
    [Expression] AS [QueryDefinition]
FROM
    $SYSTEM.TMSCHEMA_EXPRESSIONS;
