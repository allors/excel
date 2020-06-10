# Allors Excel

Allors.Excel is a c# Excel VSTO AddIn. It speeds up access to excel by using a virtual DOM to update cells.
It contains useful features to programmatically manage workbooks, worksheets, cells.

# Installing via Nuget
	Install-Package Allors.Excel
 
# Features

## Workbook

### Properties
- IsActive will activate that workbook
- Worksheets contains the worksheets inside the workbook

### Methods
- GetNamedRanges(string refersToSheetName)
	
	Return a list of Excel.Ranges
- SetNamedRange(string name, Excel.Range range)

	Adds or updates the namedRange

- Copy(IWorksheet source, IWorksheet beforeWorksheet)

	Copies the source workbook to this workbook

## Worksheet

### Properties


### Methods


### Indexers


### Events