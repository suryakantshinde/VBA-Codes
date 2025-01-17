'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
''************************************************************************Create Pivot Table
Sub Create_Pivot_Table()

'[Source Sheet Name][sheet with Rawdata]
'[OutPut Sheet Name][Where we need Pivot to get generated]
'[Pivot Table Name] [Pivot Table Name]

Dim DataRange As Range
Dim Destination As Range
Worksheets("Day_wise_summary").Select
'----------------------------------------------------------
'Set data range for pivot table
Set DataRange = Worksheets("Monthly_Data").Range("A1:AF1045756")    '} Change here
'----------------------------------------------------------
'----------------------------------------------------------
'Set destination for 'Pivot Table - ""
Set Destination = Worksheets("Day_wise_summary").Range("A1")    '} Change this Code
'----------------------------------------------------------
'Create Pivot Table
Worksheets("Monthly_Data").Select
ActiveSheet.PivotTableWizard SourceType:=xlDatabase, _
SourceData:=DataRange, TableDestination:=Destination, TableName:="Pivot Table Day wise sum"  'Change the  Pivot Table Name
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
      ''''' Need to run till here then again record further steps
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
End Sub          
