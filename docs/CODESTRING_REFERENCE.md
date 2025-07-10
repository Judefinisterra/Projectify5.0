# Codestring Reference Guide

## Overview

Codestrings are the proprietary language used by Projectify5.0 to define financial model calculations. Each codestring represents a specific calculation or assumption that feeds into the larger financial model.

## Structure

Every codestring follows this basic structure:

```
<CODETYPE; parameters; rowX="Column1|Column2|...|ColumnN|">
```

### Components

1. **Code Type**: Defines the calculation function (e.g., `CONST-E`, `MULT2-S`)
2. **Parameters**: Control behavior and references (e.g., `driver1="V1"`)
3. **Row Data**: 16 columns of data separated by pipes (`|`)

## Column Structure

Each row contains exactly 16 columns (15 pipes + ending):

```
"V1(D)|Label(L)|FinCode(F)|C1(C1)|C2(C2)|C3(C3)|C4(C4)|C5(C5)|C6(C6)|Y1(Y1)|Y2(Y2)|Y3(Y3)|Y4(Y4)|Y5(Y5)|Y6(Y6)|"
```

### Column Mapping

| Position | Label | Excel Column | Purpose |
|----------|-------|--------------|---------|
| 1 | (D) | A | Output Driver - Unique identifier |
| 2 | (L) | B | Label - Human readable description |
| 3 | (F) | C | FinCode - Financial statement mapping |
| 4 | (C1) | D | Fixed Assumption 1 |
| 5 | (C2) | E | Fixed Assumption 2 |
| 6 | (C3) | F | Fixed Assumption 3 |
| 7 | (C4) | G | Fixed Assumption 4 |
| 8 | (C5) | H | Fixed Assumption 5 |
| 9 | (C6) | I | Fixed Assumption 6 |
| 10 | (Y1) | K | Year 1 Assumption |
| 11 | (Y2) | L | Year 2 Assumption |
| 12 | (Y3) | M | Year 3 Assumption |
| 13 | (Y4) | N | Year 4 Assumption |
| 14 | (Y5) | O | Year 5 Assumption |
| 15 | (Y6) | P | Year 6 Assumption |
| 16 | - | Q | Empty (ends with `|`) |

## Code Types

### Seed Codes

These create the foundation assumptions of your model:

#### CONST-E (Constant)
Holds values constant across each year.

```javascript
<CONST-E; row1="V1(D)|Price per Unit(L)|(F)|(C1)|(C2)|(C3)|(C4)|(C5)|$10(C6)|~$10(Y1)|~$10(Y2)|~$10(Y3)|~$10(Y4)|~$10(Y5)|~$10(Y6)|">
```

**Use for:**
- Unit prices
- Percentage rates
- Cost per unit
- Monthly recurring values

#### SPREAD-E (Spread)
Divides annual values across 12 months.

```javascript
<SPREAD-E; row1="V2(D)|Annual Revenue(L)|(F)|(C1)|(C2)|(C3)|(C4)|(C5)|(C6)|$120000(Y1)|$150000(Y2)|$180000(Y3)|$210000(Y4)|$240000(Y5)|$270000(Y6)|">
```

**Use for:**
- Total annual revenues
- Annual expense amounts
- Total units sold per year
- Any aggregate annual figure

#### ENDPOINT-E (Endpoint)
Scales linearly between year-end values.

```javascript
<ENDPOINT-E; row1="V3(D)|Growing Price(L)|(F)|(C1)|(C2)|(C3)|(C4)|(C5)|$10(C6)|~$10(Y1)|~$12(Y2)|~$14(Y3)|~$16(Y4)|~$18(Y5)|~$20(Y6)|">
```

**Use for:**
- Gradually increasing prices
- Scaling rates over time
- Linear growth assumptions

### Math Operator Codes

These perform calculations using other codestrings:

#### MULT2-S (Multiply Two)
Multiplies two drivers together.

```javascript
<MULT2-S; driver1="V1"; driver2="V2"; row1="V3(D)|Total Revenue(L)|IS: revenue(F)|(C1)|(C2)|(C3)|(C4)|(C5)|(C6)|$F(Y1)|$F(Y2)|$F(Y3)|$F(Y4)|$F(Y5)|$F(Y6)|">
```

#### MULT3-S (Multiply Three)
Multiplies three drivers together.

```javascript
<MULT3-S; driver1="V1"; driver2="V2"; driver3="V3"; row1="V4(D)|Complex Calculation(L)|(F)|(C1)|(C2)|(C3)|(C4)|(C5)|(C6)|$F(Y1)|$F(Y2)|$F(Y3)|$F(Y4)|$F(Y5)|$F(Y6)|">
```

#### DIVIDE2-S (Divide Two)
Divides first driver by second driver.

```javascript
<DIVIDE2-S; driver1="V1"; driver2="V2"; row1="V3(D)|Revenue per Unit(L)|(F)|(C1)|(C2)|(C3)|(C4)|(C5)|(C6)|$F(Y1)|$F(Y2)|$F(Y3)|$F(Y4)|$F(Y5)|$F(Y6)|">
```

#### SUBTOTAL2-S (Add Two)
Adds two drivers together.

```javascript
<SUBTOTAL2-S; driver1="V1"; driver2="V2"; row1="V3(D)|Total Costs(L)|IS: direct costs(F)|(C1)|(C2)|(C3)|(C4)|(C5)|(C6)|$F(Y1)|$F(Y2)|$F(Y3)|$F(Y4)|$F(Y5)|$F(Y6)|">
```

#### SUBTOTAL3-S (Add Three)
Adds three drivers together.

```javascript
<SUBTOTAL3-S; driver1="V1"; driver2="V2"; driver3="V3"; row1="V4(D)|Total Expenses(L)|IS: corporate overhead(F)|(C1)|(C2)|(C3)|(C4)|(C5)|(C6)|$F(Y1)|$F(Y2)|$F(Y3)|$F(Y4)|$F(Y5)|$F(Y6)|">
```

#### SUBTRACT2-S (Subtract Two)
Subtracts second driver from first driver.

```javascript
<SUBTRACT2-S; driver1="V1"; driver2="V2"; row1="V3(D)|Net Profit(L)|IS: net income(F)|(C1)|(C2)|(C3)|(C4)|(C5)|(C6)|$F(Y1)|$F(Y2)|$F(Y3)|$F(Y4)|$F(Y5)|$F(Y6)|">
```

#### SUMLIST-S (Sum List)
Sums from specified driver to the row above the SUMLIST.

```javascript
<SUMLIST-S; driver1="V1"; row1="V5(D)|Total Revenue(L)|IS: revenue(F)|(C1)|(C2)|(C3)|(C4)|(C5)|(C6)|$F(Y1)|$F(Y2)|$F(Y3)|$F(Y4)|$F(Y5)|$F(Y6)|">
```

#### DIRECT-S (Direct Reference)
Directly references another driver.

```javascript
<DIRECT-S; driver1="V1"; row1="V2(D)|Revenue Copy(L)|(F)|(C1)|(C2)|(C3)|(C4)|(C5)|(C6)|$F(Y1)|$F(Y2)|$F(Y3)|$F(Y4)|$F(Y5)|$F(Y6)|">
```

### Advanced Codes

#### GROWTH-S (Growth)
Applies growth rate to previous period.

```javascript
<GROWTH-S; driver1="V1"; row1="V2(D)|Growing Revenue(L)|IS: revenue(F)|(C1)|(C2)|(C3)|(C4)|(C5)|(C6)|$F(Y1)|$F(Y2)|$F(Y3)|$F(Y4)|$F(Y5)|$F(Y6)|">
```

#### CHANGE-S (Change)
Calculates period-over-period change.

```javascript
<CHANGE-S; driver1="V1"; row1="V2(D)|Revenue Change(L)|(F)|(C1)|(C2)|(C3)|(C4)|(C5)|(C6)|$F(Y1)|$F(Y2)|$F(Y3)|$F(Y4)|$F(Y5)|$F(Y6)|">
```

#### ANNUALIZE-S (Annualize)
Multiplies by 12 for monthly to annual conversion.

```javascript
<ANNUALIZE-S; driver1="V1"; row1="V2(D)|Annual Revenue(L)|IS: revenue(F)|(C1)|(C2)|(C3)|(C4)|(C5)|(C6)|$F(Y1)|$F(Y2)|$F(Y3)|$F(Y4)|$F(Y5)|$F(Y6)|">
```

#### DEANNUALIZE-S (Deannualize)
Divides by 12 for annual to monthly conversion.

```javascript
<DEANNUALIZE-S; driver1="V1"; row1="V2(D)|Monthly Revenue(L)|(F)|(C1)|(C2)|(C3)|(C4)|(C5)|(C6)|$F(Y1)|$F(Y2)|$F(Y3)|$F(Y4)|$F(Y5)|$F(Y6)|">
```

### Reconciliation Codes

For balance sheet reconciliation tables:

#### BEGIN-S (Beginning Balance)
References ending balance from previous period.

```javascript
<BEGIN-S; driver1="V4"; row1="V1(D)|Beginning Balance(L)|(F)|(C1)|(C2)|(C3)|(C4)|(C5)|(C6)|$F(Y1)|$F(Y2)|$F(Y3)|$F(Y4)|$F(Y5)|$F(Y6)|">
```

#### END-S (Ending Balance)
Calculates ending balance for reconciliation.

```javascript
<END-S; driver1="V1"; row1="V4(D)|Ending Balance(L)|BS: current assets(F)|(C1)|(C2)|(C3)|(C4)|(C5)|(C6)|$F(Y1)|$F(Y2)|$F(Y3)|$F(Y4)|$F(Y5)|$F(Y6)|">
```

### Custom Formula Codes

#### FORMULA-S (Custom Formula)
Allows custom Excel formulas.

```javascript
<FORMULA-S; customformula="rd{V1}*cd{6-V2}"; row1="V3(D)|Custom Calculation(L)|(F)|(C1)|(C2)|(C3)|(C4)|(C5)|(C6)|$F(Y1)|$F(Y2)|$F(Y3)|$F(Y4)|$F(Y5)|$F(Y6)|">
```

**Custom Formula Functions:**
- `rd{V1}`: Reference row with driver V1
- `cd{6-V2}`: Reference column 6 of driver V2
- `BEG(cd{6})`: Start calculation from specified date
- `END(cd{6})`: End calculation at specified date
- `SPREAD(1000)`: Spread amount across 12 months
- `RANGE(rd{V1})`: Create range across entire time series

### Organizational Codes

#### LABELH1, LABELH2, LABELH3 (Headers)
Create section headers in your model.

```javascript
<LABELH1; row1="|Revenue Assumptions:(L)|(F)|(C1)|(C2)|(C3)|(C4)|(C5)|(C6)|(Y1)|(Y2)|(Y3)|(Y4)|(Y5)|(Y6)|">
<LABELH2; row1="|Unit Economics:(L)|(F)|(C1)|(C2)|(C3)|(C4)|(C5)|(C6)|(Y1)|(Y2)|(Y3)|(Y4)|(Y5)|(Y6)|">
<LABELH3; row1="|Pricing Assumptions:(L)|(F)|(C1)|(C2)|(C3)|(C4)|(C5)|(C6)|(Y1)|(Y2)|(Y3)|(Y4)|(Y5)|(Y6)|">
```

#### BR (Line Break)
Creates visual separation between sections.

```javascript
<BR; row1="|BR(L)|(F)|(C1)|(C2)|(C3)|(C4)|(C5)|(C6)|(Y1)|(Y2)|(Y3)|(Y4)|(Y5)|(Y6)|">
```

#### COLUMNHEADER-E (Column Header)
Creates column headers for assumption tables.

```javascript
<COLUMNHEADER-E; row1="|Employee Details:(L)|(F)|(C1)|(C2)|(C3)|Department(C4)|Role(C5)|Salary(C6)|(Y1)|(Y2)|(Y3)|(Y4)|(Y5)|(Y6)|">
```

## Parameters

### Common Parameters

#### driver1, driver2, driver3
Reference other codestrings by their output driver.

```javascript
driver1="V1"  // References the codestring with output driver V1
driver2="V2"  // References the codestring with output driver V2
```

#### sumif
Controls how annual columns are calculated from monthly data.

```javascript
sumif="Year"     // Sums all 12 months
sumif="Yearend"  // Takes last month value (December)
sumif="Offsetyear" // Takes first month value (January)
```

#### negative
Makes all values negative (for expenses).

```javascript
negative="true"  // Converts all values to negative
```

#### customformula
Defines custom Excel formulas (FORMULA-S only).

```javascript
customformula="rd{V1}*cd{6-V2}"  // Custom calculation
```

### Formatting Parameters

#### bold
Makes text bold.

```javascript
bold="true"   // Bold formatting
bold="false"  // Normal formatting
```

#### italic
Makes text italic.

```javascript
italic="true"   // Italic formatting
italic="false"  // Normal formatting
```

#### topborder
Adds top border (usually with bold).

```javascript
topborder="true"   // Add top border
topborder="false"  // No top border
```

#### indent
Indents the row.

```javascript
indent="1"  // Indent level 1
indent="2"  // Indent level 2
```

## FinCodes

FinCodes map calculations to financial statements:

### Income Statement
```javascript
"IS: revenue"           // Revenue line items
"IS: direct costs"      // Cost of goods sold
"IS: corporate overhead" // Operating expenses
"IS: d&a"              // Depreciation & amortization
"IS: interest expense"  // Interest payments
"IS: other income"      // Non-operating income
"IS: net income"        // Net income (rarely used directly)
```

### Balance Sheet
```javascript
"BS: current assets"     // Cash, receivables, inventory
"BS: fixed assets"       // Property, plant, equipment
"BS: current liabilities" // Short-term debt, payables
"BS: lt liabilities"     // Long-term debt
"BS: equity"            // Shareholders' equity
```

### Cash Flow Statement
```javascript
"CF: wc"       // Working capital changes
"CF: non-cash" // Non-cash items (depreciation, etc.)
"CF: cfi"      // Capital expenditures
"CF: cff"      // Financing activities
```

## Formatting Rules

### Dollar Signs
- Use `$` for currency values: `$1000`
- Use `~$` for italicized currency rates: `~$50`

### Tilde (~)
- Use `~` for italicized non-currency values: `~100`
- Use `~` for percentage labels: `~5%`

### Formula Placeholder
- Use `$F` in Y1-Y6 columns when the value is calculated
- Use `~F` for italicized calculated values

## Example Model Structure

```javascript
// Section header
<LABELH1; row1="|Revenue Model:(L)|(F)|(C1)|(C2)|(C3)|(C4)|(C5)|(C6)|(Y1)|(Y2)|(Y3)|(Y4)|(Y5)|(Y6)|">

// Line break
<BR; row1="|BR(L)|(F)|(C1)|(C2)|(C3)|(C4)|(C5)|(C6)|(Y1)|(Y2)|(Y3)|(Y4)|(Y5)|(Y6)|">

// Unit assumptions
<SPREAD-E; row1="V1(D)|Units Sold(L)|(F)|(C1)|(C2)|(C3)|(C4)|(C5)|(C6)|~1000(Y1)|~1200(Y2)|~1440(Y3)|~1728(Y4)|~2074(Y5)|~2488(Y6)|">

// Price assumptions
<CONST-E; row1="V2(D)|Price per Unit(L)|(F)|(C1)|(C2)|(C3)|(C4)|(C5)|$10(C6)|~$10(Y1)|~$10.50(Y2)|~$11.03(Y3)|~$11.58(Y4)|~$12.16(Y5)|~$12.77(Y6)|">

// Revenue calculation
<MULT2-S; driver1="V1"; driver2="V2"; bold="true"; topborder="true"; row1="V3(D)|Total Revenue(L)|IS: revenue(F)|(C1)|(C2)|(C3)|(C4)|(C5)|(C6)|$F(Y1)|$F(Y2)|$F(Y3)|$F(Y4)|$F(Y5)|$F(Y6)|">

// Cost assumptions
<CONST-E; row1="V4(D)|Cost per Unit(L)|(F)|(C1)|(C2)|(C3)|(C4)|(C5)|$4(C6)|~$4(Y1)|~$4.20(Y2)|~$4.41(Y3)|~$4.63(Y4)|~$4.86(Y5)|~$5.11(Y6)|">

// COGS calculation
<MULT2-S; driver1="V1"; driver2="V4"; negative="true"; row1="V5(D)|Cost of Goods Sold(L)|IS: direct costs(F)|(C1)|(C2)|(C3)|(C4)|(C5)|(C6)|$F(Y1)|$F(Y2)|$F(Y3)|$F(Y4)|$F(Y5)|$F(Y6)|">

// Gross profit
<SUBTRACT2-S; driver1="V3"; driver2="V5"; bold="true"; topborder="true"; row1="V6(D)|Gross Profit(L)|(F)|(C1)|(C2)|(C3)|(C4)|(C5)|(C6)|$F(Y1)|$F(Y2)|$F(Y3)|$F(Y4)|$F(Y5)|$F(Y6)|">
```

## Best Practices

### 1. Use Appropriate Seed Codes
- **CONST-E**: For rates, prices, per-unit values
- **SPREAD-E**: For annual totals that should be divided monthly
- **ENDPOINT-E**: For values that grow linearly between years

### 2. Maintain Logical Flow
- Place assumptions before calculations
- Use SUMLIST-S for totaling lists
- Apply proper formatting (bold, topborder for totals)

### 3. Expense Sign Convention
- All expenses must be negative
- Use `negative="true"` parameter or negative values in assumptions
- Ensures proper financial statement flow

### 4. Driver Management
- Use unique driver codes (V1, V2, V3, etc.)
- Reference drivers logically (calculations should reference their inputs)
- Avoid circular references

### 5. FinCode Usage
- Only final calculations should have FinCodes
- Intermediate calculations should have blank FinCodes
- Avoid double-counting in financial statements

### 6. Visual Organization
- Use BR codes liberally for visual separation
- Apply consistent indentation
- Group related calculations together

## Common Patterns

### Revenue Model
```javascript
// Units × Price = Revenue
<SPREAD-E; row1="V1(D)|Units Sold(L)|(F)|...">
<CONST-E; row1="V2(D)|Price per Unit(L)|(F)|...">
<MULT2-S; driver1="V1"; driver2="V2"; row1="V3(D)|Revenue(L)|IS: revenue(F)|...">
```

### Expense Model
```javascript
// Employees × Salary = Payroll
<CONST-E; row1="V1(D)|Number of Employees(L)|(F)|...">
<CONST-E; row1="V2(D)|Average Salary(L)|(F)|...">
<MULT2-S; driver1="V1"; driver2="V2"; negative="true"; row1="V3(D)|Payroll Expense(L)|IS: corporate overhead(F)|...">
```

### Balance Sheet Reconciliation
```javascript
// Beginning + Additions - Reductions = Ending
<BEGIN-S; driver1="V4"; row1="V1(D)|Beginning Balance(L)|(F)|...">
<SPREAD-E; row1="V2(D)|Additions(L)|(F)|...">
<SPREAD-E; row1="V3(D)|Reductions(L)|(F)|...">
<END-S; driver1="V1"; row1="V4(D)|Ending Balance(L)|BS: current assets(F)|...">
```

This reference guide provides comprehensive coverage of the codestring language. For additional help, consult the full documentation or contact support. 