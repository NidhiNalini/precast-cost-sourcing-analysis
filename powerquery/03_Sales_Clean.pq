// Sales_Clean
// Loads data/raw/sales_2025.csv, sets data types, adds month and category,
// and calculates revenue. Good source for a PivotTable (see docs/EXERCISES.md).

let
    Source = Csv.Document(
        File.Contents(ProjectFolder & "\data\raw\sales_2025.csv"),
        [Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]
    ),
    PromotedHeaders = Table.PromoteHeaders(Source, [PromoteAllScalars = true]),
    BlanksToNull = Table.ReplaceValue(
        PromotedHeaders, "", null, Replacer.ReplaceValue, {"unit_price_cad"}
    ),
    ChangedTypes = Table.TransformColumnTypes(
        BlanksToNull,
        {{"invoice_date", type date}, {"qty", Int64.Type}, {"unit_price_cad", type number}},
        "en-CA"
    ),
    AddedMonth = Table.AddColumn(
        ChangedTypes, "invoice_month", each Date.StartOfMonth([invoice_date]), type date
    ),
    AddedKey = Table.AddColumn(
        AddedMonth, "item_key", each Text.Upper(Text.Trim([item_no])), type text
    ),

    // Category from items.csv
    Items = Table.PromoteHeaders(
        Csv.Document(
            File.Contents(ProjectFolder & "\data\raw\items.csv"),
            [Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]
        ),
        [PromoteAllScalars = true]
    ),
    Merged = Table.NestedJoin(
        AddedKey, {"item_key"}, Items, {"item_no"}, "Item", JoinKind.LeftOuter
    ),
    Expanded = Table.ExpandTableColumn(Merged, "Item", {"category"}),
    FilledCategory = Table.ReplaceValue(
        Expanded, null, "UNKNOWN", Replacer.ReplaceValue, {"category"}
    ),

    AddedRevenue = Table.AddColumn(
        FilledCategory, "revenue_cad", each [qty] * [unit_price_cad], type number
    ),
    AddedStatus = Table.AddColumn(
        AddedRevenue, "status",
        each if [unit_price_cad] = null or [category] = "UNKNOWN" then "CHECK" else "OK",
        type text
    )
in
    AddedStatus
