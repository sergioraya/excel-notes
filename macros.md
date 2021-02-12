## Setup Excel Macros
* Customize ribbon -> turn on developer ribbon
* Name macro after pressing macro
* Optional: add a shortcut key
* Go to tools -> record macros
* If possible, create macros and store in personal workbook 
  * Not sure if UC is disabling
  * May need to run Excel locally and not on onedrive

## VBA code
* Go to module1 to see code 
* You can make edits right on the code
  
### Workflow
* Macro name can't have spaces - use camelCase
* Store macro in -> Personal Macro Workbook
* This doesn't work with UC license

### Project
* Create new worksheets and name them week1, 2, 3 and 4
  * add vd code to delete sheet1
* Add a Macro to quick access ribbon
  * press the macro button once saved

### Actual VBA Code
``
Sub FourWeekly()

Dim i As Integer

i = 1

Do While i <= 5
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Name = "Week " & i
    
    i = i + 1
Loop

Worksheets(1).Delete
End Sub
``

## Rescources 
[link to Master Excel Macros and VBA] https://www.youtube.com/watch?v=_tPY5BGIsJQ&t=837s
* F8 for debugg menu. Use this to understand someone else's vba code
