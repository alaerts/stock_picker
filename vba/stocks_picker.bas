Attribute VB_Name = "stocks_picker"
' VBA module for the Stock Picker workbook.
'
' Import into your workbook once via the Excel VBA editor:
'   1. Press Alt+F11 to open the VBA editor.
'   2. File -> Import File -> select this stocks_picker.bas.
'   3. Save the workbook as .xlsm if prompted.
'
' Then in Excel, assign these macros to the "Rebuild Inventory" and
' "Get Quotes" button shapes (already placed on the Main sheet by
' init-workbook):
'   - Right-click button -> Assign Macro -> RebuildInventory.
'   - Right-click button -> Assign Macro -> GetQuotes.
'
' RunPython comes from the xlwings Excel add-in (install once via
' `xlwings addin install` from the command line).

Sub RebuildInventory()
    RunPython "import stocks_report; stocks_report.button_rebuild_inventory()"
End Sub

Sub GetQuotes()
    RunPython "import stocks_report; stocks_report.button_get_quotes()"
End Sub
