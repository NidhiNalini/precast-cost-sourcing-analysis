// BOM_Clean
// Loads data/raw/bom.csv, joins it to Components_Clean, converts units,
// converts cost to CAD and calculates the extended cost of each BOM line.
// Requires the Components_Clean query.

let
    // Keep these in line with the Assumptions sheet
    FxRates = [CAD = 1.00, USD = 1.37, EUR = 1.50],
    UnitFactors = [#"LB>KG" = 0.45359237, #"G>KG" = 0.001, #"KG>T" = 0.001],

    Source = Csv.Document(
        File.Contents(ProjectFolder & "\data\raw\bom.csv"),
        [Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]
    ),
    PromotedHeaders = Table.PromoteHeaders(Source, [PromoteAllScalars = true]),
    ChangedTypes = Table.TransformColumnTypes(
        PromotedHeaders, {{"line_no", Int64.Type}, {"qty_per", type number}}, "en-CA"
    ),
    AddedKey = Table.AddColumn(
        ChangedTypes, "comp_key", each Text.Upper(Text.Trim([comp_code])), type text
    ),

    // Left join keeps BOM lines even when the component is not found
    Merged = Table.NestedJoin(
        AddedKey, {"comp_key"}, Components_Clean, {"comp_key"}, "Comp", JoinKind.LeftOuter
    ),
    Expanded = Table.ExpandTableColumn(
        Merged, "Comp",
        {"description", "uom", "unit_cost", "currency", "origin_country", "supplier_id"},
        {"comp_description", "comp_uom", "unit_cost", "currency", "origin_country", "supplier_id"}
    ),

    AddedFactor = Table.AddColumn(
        Expanded, "conv_factor",
        each if [comp_uom] = null then null
             else if [uom] = [comp_uom] then 1
             else Record.FieldOrDefault(UnitFactors, [uom] & ">" & [comp_uom], null),
        type number
    ),
    AddedQty = Table.AddColumn(
        AddedFactor, "qty_in_comp_uom", each [qty_per] * [conv_factor], type number
    ),
    AddedCostCad = Table.AddColumn(
        AddedQty, "unit_cost_cad",
        each if [currency] = null then null
             else [unit_cost] * Record.FieldOrDefault(FxRates, [currency], null),
        type number
    ),
    // null * number = null, so a missing cost stays missing (not zero)
    AddedExtended = Table.AddColumn(
        AddedCostCad, "extended_cost_cad", each [qty_in_comp_uom] * [unit_cost_cad], type number
    ),

    AddedIssues = Table.AddColumn(
        AddedExtended, "issues",
        each Text.Combine(
            List.RemoveNulls({
                if [comp_description] = null then "Unknown component" else null,
                if [comp_description] <> null and [conv_factor] = null then "No unit conversion" else null,
                if [comp_description] <> null and [unit_cost] = null then "Missing cost" else null,
                if [comp_description] <> null and [origin_country] = null then "Missing origin" else null
            }),
            "; "
        ),
        type text
    ),
    AddedStatus = Table.AddColumn(
        AddedIssues, "status", each if [issues] = "" then "OK" else "CHECK", type text
    ),

    Final = Table.SelectColumns(
        AddedStatus,
        {"parent_item", "line_no", "comp_code", "comp_key", "comp_description", "qty_per", "uom",
         "comp_uom", "conv_factor", "qty_in_comp_uom", "unit_cost_cad", "extended_cost_cad",
         "origin_country", "supplier_id", "issues", "status"}
    )
in
    Final
