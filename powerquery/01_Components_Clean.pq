// Components_Clean
// Loads data/raw/components.csv, standardises the component code, flags issues
// and keeps one row per component (the first one, same rule as the workbook).

let
    Source = Csv.Document(
        File.Contents(ProjectFolder & "\data\raw\components.csv"),
        [Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]
    ),
    PromotedHeaders = Table.PromoteHeaders(Source, [PromoteAllScalars = true]),

    // Empty text becomes null so missing values are easy to detect
    BlanksToNull = Table.ReplaceValue(
        PromotedHeaders, "", null, Replacer.ReplaceValue,
        {"unit_cost", "origin_country", "hs_code", "supplier_id"}
    ),
    ChangedTypes = Table.TransformColumnTypes(
        BlanksToNull, {{"unit_cost", type number}}, "en-CA"
    ),

    // Same cleaning rule as the workbook: TRIM + UPPER
    AddedKey = Table.AddColumn(
        ChangedTypes, "comp_key", each Text.Upper(Text.Trim([comp_code])), type text
    ),

    // Count how many rows share the same key (duplicates)
    KeyCounts = Table.Group(
        AddedKey, {"comp_key"}, {{"key_count", each Table.RowCount(_), Int64.Type}}
    ),
    MergedCounts = Table.NestedJoin(
        AddedKey, {"comp_key"}, KeyCounts, {"comp_key"}, "Counts", JoinKind.LeftOuter
    ),
    ExpandedCounts = Table.ExpandTableColumn(MergedCounts, "Counts", {"key_count"}),

    AddedIssues = Table.AddColumn(
        ExpandedCounts, "issues",
        each Text.Combine(
            List.RemoveNulls({
                if [unit_cost] = null then "Missing cost" else null,
                if [origin_country] = null then "Missing origin" else null,
                if [key_count] > 1 then "Duplicate code" else null,
                if [hs_code] = null
                    or Text.Length([hs_code]) <> 7
                    or Text.Middle([hs_code], 4, 1) <> "."
                then "HS code format" else null
            }),
            "; "
        ),
        type text
    ),
    AddedStatus = Table.AddColumn(
        AddedIssues, "status", each if [issues] = "" then "OK" else "CHECK", type text
    ),

    // Keep the first row for each key. Table.Buffer fixes the row order first.
    Deduplicated = Table.Distinct(Table.Buffer(AddedStatus), {"comp_key"}),

    Final = Table.SelectColumns(
        Deduplicated,
        {"comp_key", "description", "uom", "unit_cost", "currency",
         "supplier_id", "origin_country", "hs_code", "issues", "status"}
    )
in
    Final
