---
title: Range.FindNext method (Excel)
keywords: vbaxl10.chm144129
f1_keywords:
- vbaxl10.chm144129
ms.prod: excel
api_name:
- Excel.Range.FindNext
ms.assetid: 308c6241-2398-13e6-ba68-17ec713376f6
ms.date: 06/08/2017
localization_priority: Normal
---


# Range.FindNext method (Excel)

Continues a search that was begun with the  **[Find](Excel.Range.Find.md)** method. Finds the next cell that matches those same conditions and returns a **[Range](Excel.Range(object).md)** object that represents that cell. This does not affect the selection or the active cell.


## Syntax

_expression_. `FindNext`( `_After_` )

_expression_ A variable that represents a **[Range](excel.range(object).md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _After_|Optional| **Variant**|The cell after which you want to search. This corresponds to the position of the active cell when a search is done from the user interface. Be aware that  _After_ must be a single cell in the range. Remember that the search begins after this cell; the specified cell is not searched until the method wraps back around to this cell. If this argument is not specified, the search starts after the cell in the upper-left corner of the range.|

## Return value

Range


## Remarks

When the search reaches the end of the specified search range, it wraps around to the beginning of the range. To stop a search when this wraparound occurs, save the address of the first found cell, and then test each successive found-cell address against this saved address.


## Example

This example finds all cells in the range A1:A500 that contain the value 2 and changes their values to 5.


```vb
With Worksheets(1).Range("a1:a500")
     Set c = .Find(2, lookin:=xlValues)
     If Not c Is Nothing Then
        firstAddress = c.Address
        Do
            c.Value = 5
            Set c = .FindNext(c)
            If c is Nothing Then Exit Do
        Loop While c.Address <> firstAddress
      End If
End With
```

This example finds all the cells in the first four columns that have a constant "X" in them and hides the column that contains the X.




```vb
Public Sub HideMarkedColumns()
    
    ' Sheet1 is the code name identifier referring to the ThisWorkbook.Worksheets("Sheet1") object.
    ' there is no need to declare a new variable for a sheet that exists in ThisWorkbook at compile-time.
    ' Set the "(Name)" property of each worksheet module in the Properties toolwindow (F4)    
    With Sheet1.Range("A1:D1").SpecialCells(xlCellTypeConstants)
        Dim findResult As Range
        Set findResult = .Find(What:="X") ' note: omitted optional parameter values are carried from any previous search
        If Not findResult Is Nothing Then ' Range.Find returns Nothing if there's no match
            Dim firstAddress As String
            firstAddress = findResult.Address
             
            Do
                findResult.EntireColumn.Hidden = True
                Set findResult = .FindNext(findResult)
            Loop While findResult.Address <> firstAddress
        End If
    End With

End Sub
```

<!-- note: removed 3rd example, which added nothing relevant to the topic. -->

<!-- is this section relevant? Or do I need to add my bio as well?
### About the contributor

Dennis Wallentin is the author of VSTO & .NET & Excel, a blog that focuses on .NET Framework solutions for Excel and Excel Services. Dennis has been developing Excel solutions for over 20 years and is also the coauthor of "Professional Excel Development: The Definitive Guide to Developing Applications Using Microsoft Excel, VBA and .NET (2nd Edition)." 
-->

## See also


[Range Object](Excel.Range(object).md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
