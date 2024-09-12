# ExcelFlow: A Language for Excel File Transformation

ExcelFlow is a domain-specific language designed for filter editing, transforming, and modulating Excel files processed with the xlsx.full.min.js library, ultimately producing CSV exports. This language emphasizes readability, flexibility, and power through a unique syntax.

## Language Syntax

### Basic Structure

An ExcelFlow script consists of a series of operations, each on a new line. Operations are grouped into blocks using indentation (2 spaces).

### Variables

Variables are denoted by `$` prefix and can store cell values, ranges, or intermediate results.

```
$var = A1
$range = A1:C10
```

### Cell References

- Single cell: `A1`, `B2`, etc.
- Cell range: `A1:C10`
- Entire column: `A:A`, `B:B`, etc.
- Entire row: `1:1`, `2:2`, etc.

### Operations

1. Filter: `filter <condition>`
2. Transform: `transform <operation>`
3. Modulate: `modulate <function>`
4. Export: `export <options>`

### Conditions

Conditions use a prefix notation:

```
== (equal)
!= (not equal)
> (greater than)
< (less than)
>= (greater than or equal)
<= (less than or equal)
&& (and)
|| (or)
! (not)
```

### Functions

Built-in functions:

- `sum()`: Sum of values
- `avg()`: Average of values
- `count()`: Count of cells
- `concat()`: Concatenate strings
- `split()`: Split string into array
- `join()`: Join array into string
- `map()`: Apply function to each element
- `reduce()`: Reduce array to single value
- `pivot()`: Create pivot table

### Syntax Features

1. Pipe operator `|>` for chaining operations
2. Spread operator `...` for expanding ranges
3. Destructuring assignment for working with cell ranges
4. Pattern matching for complex conditionals
5. Partial application of functions using `_` placeholder

## Example of an Enterprise Report
Here's an advanced example demonstrating the power and flexibility of ExcelFlow:

```javascript
# Load the Excel file
load "sales_data.xlsx"

# Define variables
$sales_range = A2:E1000
$date_col = A:A
$product_col = B:B
$quantity_col = C:C
$price_col = D:D
$total_col = E:E

# Filter out rows with zero quantity
filter $sales_range |> != $quantity_col 0

# Add a new column for categorizing products
transform $sales_range |>
  add_column F "Category"
  map F:F {
    pattern_match $product_col
      case /^Electronics/ => "Tech"
      case /^Clothing/ => "Apparel"
      case /^Books/ => "Literature"
      case _ => "Other"
  }

# Calculate total revenue and add to a new column
transform $sales_range |>
  add_column G "Revenue"
  map G:G {
    $quantity, $price = destructure $quantity_col, $price_col
    * $quantity $price
  }

# Modulate the date format
modulate $date_col |>
  map { format_date "YYYY-MM-DD" }

# Create a pivot table for sales by category and month
$pivot_table = pivot {
  source: $sales_range
  rows: [F, month($date_col)]
  columns: []
  values: [
    sum(G) as "Total Revenue"
    avg(G) as "Average Revenue"
    count(A) as "Number of Sales"
  ]
}

# Export the results
export {
  filename: "sales_analysis.csv"
  sheets: [
    {
      name: "Filtered Sales"
      range: $sales_range
    },
    {
      name: "Sales by Category and Month"
      range: $pivot_table
    }
  ]
  options: {
    delimiter: ","
    include_headers: true
    date_format: "YYYY-MM-DD"
  }
}
```

This example demonstrates the following features:

1. Loading an Excel file
2. Defining variables for easy reference
3. Filtering out rows with zero quantity
4. Adding a new column with categorization using pattern matching
5. Calculating revenue using destructuring assignment
6. Modulating date format
7. Creating a pivot table for sales analysis
8. Exporting multiple sheets to a CSV file with custom options

The language allows for complex operations to be expressed concisely and readably, making it powerful for Excel file transformation tasks.
