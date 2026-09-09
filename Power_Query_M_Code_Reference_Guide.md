# Power Query M Code Reference Guide

This guide documents the main `power` query and the analysis queries used for the Excel dashboard.

> **Important:** All analysis queries below assume the main source query is named `power`.

---

# 1. Main Query — `power`

## Purpose
Connects to the folder, selects the most recently modified Excel file, opens the `Meter_Infrastructure_Data` sheet, promotes headers, and applies data types.

```powerquery
let
    Source = Folder.Files("C:\Users\User\OneDrive\Desktop\power"),

    #"Excel Files Only" = Table.SelectRows(
        Source,
        each [Extension] = ".xlsx"
    ),

    #"Sorted Files" = Table.Sort(
        #"Excel Files Only",
        {
            {"Date modified", Order.Descending}
        }
    ),

    #"Latest File" = Table.FirstN(
        #"Sorted Files",
        1
    ),

    #"Latest File Content" = #"Latest File"{0}[Content],

    #"Imported Excel Workbook" = Excel.Workbook(
        #"Latest File Content",
        null,
        true
    ),

    Meter_Infrastructure_Data_Sheet =
        #"Imported Excel Workbook"{
            [
                Item = "Meter_Infrastructure_Data",
                Kind = "Sheet"
            ]
        }[Data],

    #"Promoted Headers" = Table.PromoteHeaders(
        Meter_Infrastructure_Data_Sheet,
        [PromoteAllScalars = true]
    ),

    #"Changed Type" = Table.TransformColumnTypes(
        #"Promoted Headers",
        {
            {"Record_ID", type text},
            {"Reading_Date", type date},
            {"Region", type text},
            {"County", type text},
            {"Service_Area", type text},
            {"Feeder_ID", type text},
            {"Transformer_ID", type text},
            {"Meter_ID", type text},
            {"Customer_Type", type text},
            {"Connection_Type", type text},
            {"Meter_Type", type text},
            {"Meter_Age_Years", type number},
            {"Transformer_Rating_kVA", Int64.Type},
            {"Transformer_Load_pct", type number},
            {"Voltage_V", type number},
            {"Current_A", type number},
            {"Energy_Imported_kWh", type number},
            {"Energy_Billed_kWh", type number},
            {"Energy_Loss_pct", type number},
            {"Previous_Reading_kWh", type number},
            {"Current_Reading_kWh", type number},
            {"Consumption_kWh", type number},
            {"Target_Collection_pct", Int64.Type},
            {"Actual_Collection_pct", type number},
            {"Arrears_KES", type number},
            {"Outage_Count_30d", Int64.Type},
            {"Outage_Duration_Hrs_30d", type number},
            {"Meter_Communication_pct", type number},
            {"Last_Inspection_Date", type date},
            {"Days_Since_Inspection", Int64.Type},
            {"Tamper_Flag", type text},
            {"Estimated_Billing_Flag", type text},
            {"Meter_Status", type text},
            {"Transformer_Status", type text},
            {"Work_Order_Status", type text},
            {"Fault_Type", type text},
            {"Response_Time_Hrs", type number},
            {"Target_Response_Time_Hrs", Int64.Type},
            {"Priority", type text},
            {"Notes", type text}
        }
    )
in
    #"Changed Type"
```

---

# 2. `voltage status`

## Purpose
Classifies voltage readings as Low Voltage, Normal, or High Voltage.

```powerquery
let
    Source = power,

    #"Added VoltageStatus" = Table.AddColumn(
        Source,
        "VoltageStatus",
        each
            if [Voltage_V] <= 224 then "Low Voltage"
            else if [Voltage_V] >= 240 then "High Voltage"
            else "Normal",
        type text
    )
in
    #"Added VoltageStatus"
```

---

# 3. `counties`

## Purpose
Counts the number of records per county.

```powerquery
let
    Source = power,

    #"County Totals" = Table.Group(
        Source,
        {"County"},
        {
            {"Total", each Table.RowCount(_), Int64.Type}
        }
    )
in
    #"County Totals"
```

---

# 4. `CountyStats`

## Purpose
Creates a county performance summary for imported energy, billed energy, billing gap, billing performance, collection target, actual collection, and collection variance.

```powerquery
let
    Source = power,

    #"County Summary" = Table.Group(
        Source,
        {"County"},
        {
            {"Energy Imported", each List.Sum([Energy_Imported_kWh]), type number},
            {"Energy Billed", each List.Sum([Energy_Billed_kWh]), type number},
            {"Target Collection %", each List.Average([Target_Collection_pct]), type number},
            {"Actual Collection %", each List.Average([Actual_Collection_pct]), type number}
        }
    ),

    #"Added Billing Gap" = Table.AddColumn(
        #"County Summary",
        "Billing Gap",
        each [Energy Imported] - [Energy Billed],
        type number
    ),

    #"Added Billing Performance" = Table.AddColumn(
        #"Added Billing Gap",
        "Billing Performance %",
        each
            if [Energy Imported] = 0 then null
            else ([Energy Billed] / [Energy Imported]) * 100,
        type number
    ),

    #"Added Collection Variance" = Table.AddColumn(
        #"Added Billing Performance",
        "Collection Variance %",
        each [#"Actual Collection %"] - [#"Target Collection %"],
        type number
    )
in
    #"Added Collection Variance"
```

---

# 5. `Service_Area`

## Purpose
Summarizes records, arrears, outages, and average voltage by county and service area.

```powerquery
let
    Source = power,

    #"Service Area Summary" = Table.Group(
        Source,
        {"County", "Service_Area"},
        {
            {"Total Records", each Table.RowCount(_), Int64.Type},
            {"Total Arrears", each List.Sum([Arrears_KES]), type number},
            {"Total Outages", each List.Sum([Outage_Count_30d]), Int64.Type},
            {"Average Voltage", each List.Average([Voltage_V]), type number}
        }
    )
in
    #"Service Area Summary"
```

---

# 6. `County Voltage Summary`

## Purpose
Compares average, minimum, and maximum voltage by county and assigns an overall voltage status.

```powerquery
let
    Source = power,

    #"County Voltage Summary" = Table.Group(
        Source,
        {"County"},
        {
            {"Average Voltage", each List.Average([Voltage_V]), type number},
            {"Minimum Voltage", each List.Min([Voltage_V]), type number},
            {"Maximum Voltage", each List.Max([Voltage_V]), type number},
            {"Total Records", each Table.RowCount(_), Int64.Type}
        }
    ),

    #"Added Voltage Status" = Table.AddColumn(
        #"County Voltage Summary",
        "CountyVoltageStatus",
        each
            if [Average Voltage] <= 224 then "Low Voltage"
            else if [Average Voltage] >= 240 then "High Voltage"
            else "Normal",
        type text
    )
in
    #"Added Voltage Status"
```

---

# 7. `Outage Performance`

## Purpose
Summarizes outage count and outage duration by county.

```powerquery
let
    Source = power,

    #"County Outage Summary" = Table.Group(
        Source,
        {"County"},
        {
            {"Total Outages", each List.Sum([Outage_Count_30d]), Int64.Type},
            {"Total Outage Hours", each List.Sum([Outage_Duration_Hrs_30d]), type number},
            {"Average Outage Duration", each List.Average([Outage_Duration_Hrs_30d]), type number}
        }
    ),

    #"Sorted Counties" = Table.Sort(
        #"County Outage Summary",
        {
            {"Total Outages", Order.Descending}
        }
    )
in
    #"Sorted Counties"
```

---

# 8. `Collection Performance`

## Purpose
Compares average collection target against average actual collection by county.

```powerquery
let
    Source = power,

    #"County Collection Summary" = Table.Group(
        Source,
        {"County"},
        {
            {"Average Target Collection %", each List.Average([Target_Collection_pct]), type number},
            {"Average Actual Collection %", each List.Average([Actual_Collection_pct]), type number},
            {"Total Arrears KES", each List.Sum([Arrears_KES]), type number}
        }
    ),

    #"Added Collection Gap" = Table.AddColumn(
        #"County Collection Summary",
        "CollectionGap_pct",
        each [#"Average Actual Collection %"] - [#"Average Target Collection %"],
        type number
    ),

    #"Added Performance Status" = Table.AddColumn(
        #"Added Collection Gap",
        "PerformanceStatus",
        each
            if [CollectionGap_pct] >= 0 then "Target Achieved"
            else "Below Target",
        type text
    )
in
    #"Added Performance Status"
```

---

# 9. `High-Arrears Flag`

## Purpose
Classifies records into High, Medium, or Low arrears.

```powerquery
let
    Source = power,

    #"Added Arrears Status" = Table.AddColumn(
        Source,
        "ArrearsStatus",
        each
            if [Arrears_KES] >= 100000 then "High Arrears"
            else if [Arrears_KES] >= 50000 then "Medium Arrears"
            else "Low Arrears",
        type text
    )
in
    #"Added Arrears Status"
```

> Adjust the KES thresholds to match the organization's real business rules.

---

# 10. `CommunicationStatus`

## Purpose
Classifies meter communication performance.

```powerquery
let
    Source = power,

    #"Added Communication Status" = Table.AddColumn(
        Source,
        "CommunicationStatus",
        each
            if [Meter_Communication_pct] < 50 then "Poor"
            else if [Meter_Communication_pct] < 90 then "Needs Attention"
            else "Good",
        type text
    )
in
    #"Added Communication Status"
```

---

# 11. `InspectionStatus`

## Purpose
Classifies inspection recency.

```powerquery
let
    Source = power,

    #"Added Inspection Status" = Table.AddColumn(
        Source,
        "InspectionStatus",
        each
            if [Days_Since_Inspection] > 180 then "Overdue"
            else if [Days_Since_Inspection] > 90 then "Due Soon"
            else "Up to Date",
        type text
    )
in
    #"Added Inspection Status"
```

> The 90-day and 180-day thresholds are training examples.

---

# 12. `ResponseTimeVariance`

## Purpose
Calculates actual response time minus target response time.

```powerquery
let
    Source = power,

    #"Added Response Variance" = Table.AddColumn(
        Source,
        "ResponseVariance_Hrs",
        each [Response_Time_Hrs] - [Target_Response_Time_Hrs],
        type number
    ),

    #"Added Response Status" = Table.AddColumn(
        #"Added Response Variance",
        "ResponseStatus",
        each
            if [ResponseVariance_Hrs] <= 0 then "Within Target"
            else "Target Missed",
        type text
    )
in
    #"Added Response Status"
```

---

# 13. `CustomerType`

## Purpose
Shows the number of customers of each customer type by county, with customer types pivoted into columns.

```powerquery
let
    Source = power,

    #"Grouped Data" = Table.Group(
        Source,
        {"County", "Customer_Type"},
        {
            {"Total", each Table.RowCount(_), Int64.Type}
        }
    ),

    #"Pivoted Data" = Table.Pivot(
        #"Grouped Data",
        List.Distinct(#"Grouped Data"[Customer_Type]),
        "Customer_Type",
        "Total",
        List.Sum
    ),

    #"Replaced Nulls" = Table.ReplaceValue(
        #"Pivoted Data",
        null,
        0,
        Replacer.ReplaceValue,
        List.RemoveItems(
            Table.ColumnNames(#"Pivoted Data"),
            {"County"}
        )
    )
in
    #"Replaced Nulls"
```

---

# 14. `ConnectionType`

## Purpose
Shows connection type counts by county.

```powerquery
let
    Source = power,

    #"Grouped Data" = Table.Group(
        Source,
        {"County", "Connection_Type"},
        {
            {"Total", each Table.RowCount(_), Int64.Type}
        }
    ),

    #"Pivoted Data" = Table.Pivot(
        #"Grouped Data",
        List.Distinct(#"Grouped Data"[Connection_Type]),
        "Connection_Type",
        "Total",
        List.Sum
    ),

    #"Replaced Nulls" = Table.ReplaceValue(
        #"Pivoted Data",
        null,
        0,
        Replacer.ReplaceValue,
        List.RemoveItems(
            Table.ColumnNames(#"Pivoted Data"),
            {"County"}
        )
    )
in
    #"Replaced Nulls"
```

---

# 15. `Meter_Type`

## Purpose
Shows meter type counts by county.

```powerquery
let
    Source = power,

    #"Grouped Data" = Table.Group(
        Source,
        {"County", "Meter_Type"},
        {
            {"Total", each Table.RowCount(_), Int64.Type}
        }
    ),

    #"Pivoted Data" = Table.Pivot(
        #"Grouped Data",
        List.Distinct(#"Grouped Data"[Meter_Type]),
        "Meter_Type",
        "Total",
        List.Sum
    ),

    #"Replaced Nulls" = Table.ReplaceValue(
        #"Pivoted Data",
        null,
        0,
        Replacer.ReplaceValue,
        List.RemoveItems(
            Table.ColumnNames(#"Pivoted Data"),
            {"County"}
        )
    )
in
    #"Replaced Nulls"
```

---

# 16. `Meter_Status`

## Purpose
Shows meter status counts by county.

```powerquery
let
    Source = power,

    #"Grouped Data" = Table.Group(
        Source,
        {"County", "Meter_Status"},
        {
            {"Total", each Table.RowCount(_), Int64.Type}
        }
    ),

    #"Pivoted Data" = Table.Pivot(
        #"Grouped Data",
        List.Distinct(#"Grouped Data"[Meter_Status]),
        "Meter_Status",
        "Total",
        List.Sum
    ),

    #"Replaced Nulls" = Table.ReplaceValue(
        #"Pivoted Data",
        null,
        0,
        Replacer.ReplaceValue,
        List.RemoveItems(
            Table.ColumnNames(#"Pivoted Data"),
            {"County"}
        )
    )
in
    #"Replaced Nulls"
```

---

# 17. `TransformerStatus`

## Purpose
Shows transformer status counts by county.

```powerquery
let
    Source = power,

    #"Grouped Data" = Table.Group(
        Source,
        {"County", "Transformer_Status"},
        {
            {"Total", each Table.RowCount(_), Int64.Type}
        }
    ),

    #"Pivoted Data" = Table.Pivot(
        #"Grouped Data",
        List.Distinct(#"Grouped Data"[Transformer_Status]),
        "Transformer_Status",
        "Total",
        List.Sum
    ),

    #"Replaced Nulls" = Table.ReplaceValue(
        #"Pivoted Data",
        null,
        0,
        Replacer.ReplaceValue,
        List.RemoveItems(
            Table.ColumnNames(#"Pivoted Data"),
            {"County"}
        )
    )
in
    #"Replaced Nulls"
```

---

# 18. `WorkStatus`

## Purpose
Shows work order status counts by county.

```powerquery
let
    Source = power,

    #"Grouped Data" = Table.Group(
        Source,
        {"County", "Work_Order_Status"},
        {
            {"Total", each Table.RowCount(_), Int64.Type}
        }
    ),

    #"Pivoted Data" = Table.Pivot(
        #"Grouped Data",
        List.Distinct(#"Grouped Data"[Work_Order_Status]),
        "Work_Order_Status",
        "Total",
        List.Sum
    ),

    #"Replaced Nulls" = Table.ReplaceValue(
        #"Pivoted Data",
        null,
        0,
        Replacer.ReplaceValue,
        List.RemoveItems(
            Table.ColumnNames(#"Pivoted Data"),
            {"County"}
        )
    )
in
    #"Replaced Nulls"
```

---

# 19. `Fault_Type`

## Purpose
Shows fault type counts by county.

```powerquery
let
    Source = power,

    #"Grouped Data" = Table.Group(
        Source,
        {"County", "Fault_Type"},
        {
            {"Total", each Table.RowCount(_), Int64.Type}
        }
    ),

    #"Pivoted Data" = Table.Pivot(
        #"Grouped Data",
        List.Distinct(#"Grouped Data"[Fault_Type]),
        "Fault_Type",
        "Total",
        List.Sum
    ),

    #"Replaced Nulls" = Table.ReplaceValue(
        #"Pivoted Data",
        null,
        0,
        Replacer.ReplaceValue,
        List.RemoveItems(
            Table.ColumnNames(#"Pivoted Data"),
            {"County"}
        )
    )
in
    #"Replaced Nulls"
```

---

# 20. `ServiceAreaPriority`

## Purpose
Shows priority counts by service area. This is useful for identifying service areas carrying the most high-priority work.

```powerquery
let
    Source = power,

    #"Grouped Data" = Table.Group(
        Source,
        {"Service_Area", "Priority"},
        {
            {"Total", each Table.RowCount(_), Int64.Type}
        }
    ),

    #"Pivoted Data" = Table.Pivot(
        #"Grouped Data",
        List.Distinct(#"Grouped Data"[Priority]),
        "Priority",
        "Total",
        List.Sum
    ),

    #"Replaced Nulls" = Table.ReplaceValue(
        #"Pivoted Data",
        null,
        0,
        Replacer.ReplaceValue,
        List.RemoveItems(
            Table.ColumnNames(#"Pivoted Data"),
            {"Service_Area"}
        )
    )
in
    #"Replaced Nulls"
```

---

# 21. `Tamper_Flag`

## Purpose
Shows tamper flag counts by county.

```powerquery
let
    Source = power,

    #"Grouped Data" = Table.Group(
        Source,
        {"County", "Tamper_Flag"},
        {
            {"Total", each Table.RowCount(_), Int64.Type}
        }
    ),

    #"Pivoted Data" = Table.Pivot(
        #"Grouped Data",
        List.Distinct(#"Grouped Data"[Tamper_Flag]),
        "Tamper_Flag",
        "Total",
        List.Sum
    ),

    #"Replaced Nulls" = Table.ReplaceValue(
        #"Pivoted Data",
        null,
        0,
        Replacer.ReplaceValue,
        List.RemoveItems(
            Table.ColumnNames(#"Pivoted Data"),
            {"County"}
        )
    )
in
    #"Replaced Nulls"
```

---

# 22. `EstimatedBillingFlag`

## Purpose
Shows estimated billing flag counts by county.

```powerquery
let
    Source = power,

    #"Grouped Data" = Table.Group(
        Source,
        {"County", "Estimated_Billing_Flag"},
        {
            {"Total", each Table.RowCount(_), Int64.Type}
        }
    ),

    #"Pivoted Data" = Table.Pivot(
        #"Grouped Data",
        List.Distinct(#"Grouped Data"[Estimated_Billing_Flag]),
        "Estimated_Billing_Flag",
        "Total",
        List.Sum
    ),

    #"Replaced Nulls" = Table.ReplaceValue(
        #"Pivoted Data",
        null,
        0,
        Replacer.ReplaceValue,
        List.RemoveItems(
            Table.ColumnNames(#"Pivoted Data"),
            {"County"}
        )
    )
in
    #"Replaced Nulls"
```

---

# Suggested Dashboard Use

| Query | Suggested Visual |
|---|---|
| counties | Bar chart: records by county |
| CountyStats | KPI cards + county performance table |
| Service_Area | Service area comparison |
| County Voltage Summary | Column chart |
| Outage Performance | Bar chart |
| Collection Performance | Target vs Actual chart |
| High-Arrears Flag | KPI / status chart |
| CommunicationStatus | Communication quality chart |
| InspectionStatus | Inspection status chart |
| ResponseTimeVariance | Response performance chart |
| CustomerType | Stacked column chart |
| ConnectionType | Stacked column chart |
| Meter_Type | Stacked column chart |
| Meter_Status | Stacked bar chart |
| TransformerStatus | Column chart |
| WorkStatus | Work order status chart |
| Fault_Type | Fault distribution chart |
| ServiceAreaPriority | Priority by service area |
| Tamper_Flag | Tamper analysis chart |
| EstimatedBillingFlag | Estimated billing chart |

---

# Refresh Workflow

When a new complete Excel file is downloaded into:

```text
C:\Users\User\OneDrive\Desktop\power
```

the workflow is:

```text
New Excel File
      ↓
power selects newest file
      ↓
Analysis queries refresh
      ↓
Excel tables refresh
      ↓
Charts refresh
      ↓
Dashboard updates
```

Use:

**Excel → Data → Refresh All**

to refresh the entire solution.
