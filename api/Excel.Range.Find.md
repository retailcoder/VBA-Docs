---
title: Range.Find method (Excel)
keywords: vbaxl10.chm144128
f1_keywords:
- vbaxl10.chm144128
ms.prod: excel
api_name:
- Excel.Range.Find
ms.assetid: d9585265-8164-cb4d-a9e3-262f6e06b6b8
ms.date: 04/19/2019
localization_priority: Priority
---


# Range.Find method (Excel)

Finds specific information in a range.

## Syntax

_expression_.**Find** (_What_, _After_, _LookIn_, _LookAt_, _SearchOrder_, _SearchDirection_, _MatchCase_, _MatchByte_, _SearchFormat_)

_expression_ A variable that represents a **[Range](excel.range(object).md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _What_|Required| **Variant**|The data to search for. Can be a string or any Microsoft Excel data type.|
| _After_|Optional| **Variant**|The cell after which you want the search to begin. This corresponds to the position of the active cell when a search is done from the user interface.<br/><br/>Notice that _After_ must be a single cell in the range. Remember that the search begins after this cell; the specified cell isn't searched until the method wraps back around to this cell.<br/><br/>If you do not specify this argument, the search starts after the cell in the upper-left corner of the range.|
| _LookIn_|Optional| **Variant**|Can be one of the following **[XlFindLookIn](excel.xlfindlookin.md)** constants: **xlFormulas**, **xlValues**, or **xlComments**.|
| _LookAt_|Optional| **Variant**|Can be one of the following **[XlLookAt](excel.xllookat.md)** constants: **xlWhole** or **xlPart**.|
| _SearchOrder_|Optional| **Variant**|Can be one of the following **[XlSearchOrder](excel.xlsearchorder.md)** constants: **xlByRows** or **xlByColumns**.|
| _SearchDirection_|Optional| **[XlSearchDirection](Excel.xlSearchDirection.md)** |The search direction.|
| _MatchCase_|Optional| **Variant**| **True** to make the search case-sensitive. The default value is **False**.|
| _MatchByte_|Optional| **Variant**|Used only if you have selected or installed double-byte language support. **True** to have double-byte characters match only double-byte characters. **False** to have double-byte characters match their single-byte equivalents.|
| _SearchFormat_|Optional| **Variant**|The search format.|

## Return value

A **Range** object that represents the first cell where that information is found.


## Remarks

This method returns **Nothing** if no match is found. The **Find** method does not affect the selection or the active cell.

The settings for _LookIn_,  _LookAt_,  _SearchOrder_, and _MatchByte_ are saved each time you use this method. If you do not specify values for these arguments the next time you call the method, the saved values are used. Setting these arguments changes the settings in the **Find** dialog box, and changing the settings in the **Find** dialog box changes the saved values that are used if you omit the arguments. To avoid problems, set these arguments explicitly each time you use this method.

You can use the **[FindNext](Excel.Range.FindNext.md)** and **[FindPrevious](Excel.Range.FindPrevious.md)** methods to repeat the search.

When the search reaches the end of the specified search range, it wraps around to the beginning of the range. To stop a search when this wraparound occurs, save the address of the first found cell, and then test each successive found-cell address against this saved address.

To find cells that match more complicated patterns, use a **For Each...Next** statement with the **Like** operator. For example, the following code searches for all cells in the range A1:C5 (in the active worksheet) that use a font whose name starts with the letters Cour. When a match is found, it changes the font to Times New Roman.

```vb
Dim c As Range
For Each c In ActiveSheet.Range("A1:C5")
    If c.Font.Name Like "Cour*" Then 
        c.Font.Name = "Times New Roman" 
    End If 
Next
```

## Example

This example finds all cells in the range A1:A500 on worksheet one that contain the value 2, and changes it to 5.

```vb
With ActiveWorkbook.Worksheets(1).Range("A1:A500") 
    Dim findResult As Range
    Set findResult = .Find(2, lookin:=xlValues) 
    If Not findResult Is Nothing Then 
        firstAddress = findResult.Address 
        Do 
            findResult.Value = 5 
            Set findResult = .FindNext(c) 
        Loop While Not findResult Is Nothing
    End If 
End With
```

<!-- is this attribution/advertisement still relevant?

**Sample code provided by:** Holy Macro! Books, [Holy Macro! It's 2,500 Excel VBA Examples](https://www.mrexcel.com/store/index.php?l=product_detail&p=1).
-->

This example takes a path and name of a workbook and a search term, and searches the active sheet in the specified workbook for the search term. If the search term is found, the address of the result is stored in cell D10 of the current workbook.

```vb
Option Explicit

Public Sub FindAddress()
    On Error GoTo CleanFail
    
    Dim searchText As String
    searchText = "Hello"
    
    'Use the current sheet as the place to store the data for which to search.
    Dim resultSheet As Worksheet
    Set resultSheet = ActiveSheet

    'The path & file name for the workbook in which to search.
    Const wbPath As String = "C:\Your\File\Path\"
    Const wbName As String = "YourFileName.xls"
    
    Dim wb As Workbook
    If Not TryGetWorkbookToSearch(wbPath & wbName, wb) Then
        MsgBox "The workbook '" & wbName & "' could not be found in" & vbNewLine & "'" & wbPath & "'.", vbExclamation
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    'Search for the specified text in whichever sheet is active in that workbook
    Dim findResult As Range
    Set findResult = wb.ActiveSheet.UsedRange.Find(searchText)
    If Not findResult Is Nothing Then
        'Record the address of the data, along with the date, in the current workbook.
        With resultSheet.Range("D10")
            .Value = "Address of " & searchText & ":"
            .Offset(0, 1).Value = "Date:"
            .Offset(1, 0).Value = findResult.Address
            .Offset(1, 1).Value = Date
            .Columns.AutoFit
            .Offset(1, 1).Columns.AutoFit
        End With
    Else
        MsgBox "The value '" & searchText & "' was not found in the active sheet of the specified workbook."
    End If
        
CleanExit:
    'Close the data workbook, without saving any changes, and turn screen updating back on.
    If Not wb Is Nothing Then wb.Close savechanges:=False
    Application.ScreenUpdating = True
    Exit Sub

CleanFail:
    MsgBox Err.Description, vbExclamation
    Resume CleanExit 'debug: place a breakpoint here (F9) and set the next statement to the "Resume" instruction underneath.
    Resume 'set as next statement and press F8 to jump to the statement that raised the error.
End Sub

Private Function TryGetWorkbookToSearch(ByVal wbFullPath As String, ByRef outWorkbook As Workbook) As Boolean
    On Error GoTo CleanFail
    Dim result As Boolean
    Set outWorkbook = Application.Workbooks.Open(wbFullPath)
    result = True
CleanExit:
    TryGetWorkbookToSearch = result
    Exit Function
CleanFail:
    Set outWorkbook = Nothing
    result = False
    Resume CleanExit
End Function
```

<!-- I count 8 contributors to this file, excluding myself.
### About the contributor

Holy Macro! Books publishes entertaining books for people who use Microsoft Office. See the complete catalog at MrExcel.com. 
-->


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
