# Simple Inventory Tracker

**Veda Technology Internship · Data Analytics Track · Task 30 (Level 1)**

A rule-based inventory tracking sheet built in Excel. It shows current stock, reorder level, and an automatic stock status for each item, plus a low-stock report for items that need restocking.

## Files

| File | Description |
|------|-------------|
| `inventory_tracker.xlsx` | Workbook with the `Inventory` and `Low-Stock Report` sheets |
| `README.md` | Approach and outcome (this file) |

## Approach

### Sheet 1: Inventory
Columns: Item ID, Item Name, Category, Current Stock, Reorder Level, Supplier, **Stock Status**, **Units to Order**.

Stock Status uses nested `IF` formulas:

```excel
=IF(D2=0,"Out of Stock",IF(D2<=E2,"Reorder Now","In Stock"))
```

Units to Order restocks up to twice the reorder level:

```excel
=IF(D2<=E2,E2*2-D2,0)
```

Conditional formatting highlights each row:
- Red: Out of Stock
- Yellow: Reorder Now
- Green: In Stock

### Sheet 2: Low-Stock Report
Lists every item at or below its reorder level, pulled from the Inventory sheet by formula so it stays in sync. It also shows the total count of items to reorder.

## Outcome
- 12 sample items tracked
- 5 items flagged for reorder (1 out of stock, 4 at or below reorder level)
- Status updates automatically when Current Stock or Reorder Level changes

## Key Concepts
- **Reorder level:** the stock quantity at which a new order should be placed, so the item does not run out before the new stock arrives.
- **Why track stock status:** it turns raw counts into a clear action (order now or not), prevents stockouts and overstocking, and lets managers prioritise.

## Tools
Microsoft Excel (formulas: `IF`, `COUNTIF`; conditional formatting; filters)

## Notes
- The dataset is sample data created for this exercise.
- The Low-Stock Report rows are linked to the items flagged at build time. If stock changes so a different item becomes low, filter the Inventory sheet on Stock Status.
