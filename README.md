# Simple SupplyChain — Access Inventory Tracker

A small inventory/order tracker built in MS Access, demonstrating relational
schema design, forms, and VBA automation.

## Features
- Linked Products, Suppliers, Categories, and Orders tables
- Data entry forms with navigation, add, and delete
- VBA-driven data quality check (duplicate orders, suspicious quantities)
- One-click export to Excel

## Security
- All data-insertion logic uses parameterized queries (DAO 'QueryDef'+ 'Parameters') rather than string=concatenated SQL to avoid injection risks

## Schema
![Relationships](docs/Relationships.png)

## Known limitations
- Form and subform design needs work and added functionality 
- Needs more records for better query visualization
- Some automatization for inserting more records (maybe with randomization) would be nice
- Need to work with reports (generate, add to forms)
- Create more queries to demonstrate functionality with entry parameters
- Create main Form which opens on entry to navigate different windows and macros/modules
- Create more tables in the future to demonstrate 1-1, many-many relationships

## How to run
1. Open `SimpleSupplyChain.accdb`
2. Enable content/macros if prompted
3. Use the switchboard buttons to navigate