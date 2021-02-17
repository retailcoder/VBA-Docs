---
title: WorksheetFunction.Vdb method (Excel)
keywords: vbaxl10.chm137161
f1_keywords:
- vbaxl10.chm137161
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Vdb
ms.assetid: 601a57eb-56da-c3e5-4e6c-3029202c317d
ms.date: 05/25/2019
localization_priority: Normal
---


# WorksheetFunction.Vdb method (Excel)

Returns the depreciation of an asset for any period that you specify, including partial periods, by using the double-declining balance method or some other method that you specify. **Vdb** stands for variable declining balance.


## Syntax

_expression_.**Vdb** (_Arg1_, _Arg2_, _Arg3_, _Arg4_, _Arg5_, _Arg6_, _Arg7_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|Cost - the initial cost of the asset.|
| _Arg2_|Required| **Double**|Salvage - the value at the end of the depreciation (sometimes called the salvage value of the asset). This value can be 0.|
| _Arg3_|Required| **Double**|Life - the number of periods over which the asset is depreciated (sometimes called the useful life of the asset).|
| _Arg4_|Required| **Double**|Start_period - the starting period for which you want to calculate the depreciation. Start_period must use the same units as life.|
| _Arg5_|Required| **Double**|End_period - the ending period for which you want to calculate the depreciation. End_period must use the same units as life.|
| _Arg6_|Optional| **Variant**|Factor - the rate at which the balance declines. If factor is omitted, it is assumed to be 2 (the double-declining balance method). Change factor if you do not want to use the double-declining balance method. For a description of the double-declining balance method, see **[Ddb](excel.worksheetfunction.ddb.md)**.|
| _Arg7_|Optional| **Variant**|No_switch - a logical value specifying whether to switch to straight-line depreciation when depreciation is greater than the declining balance calculation.|

## Return value

**Double**


## Remarks

If no_switch is **True**, Microsoft Excel does not switch to straight-line depreciation even when the depreciation is greater than the declining balance calculation.
    
If no_switch is **False** or omitted, Excel switches to straight-line depreciation when depreciation is greater than the declining balance calculation.
    
All arguments except no_switch must be positive numbers.



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]