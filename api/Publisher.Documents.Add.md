---
title: Documents.Add method (Publisher)
keywords: vbapb10.chm8650756
f1_keywords:
- vbapb10.chm8650756
ms.prod: publisher
api_name:
- Publisher.Documents.Add
ms.assetid: 1e3536c8-8fc0-8c95-3a4c-b16fe8a99098
ms.date: 06/06/2019
localization_priority: Normal
---


# Documents.Add method (Publisher)

Adds a new **[Document](Publisher.Document.md)** object that represents a new publication to the **Documents** collection.


## Syntax

_expression_.**Add** (_PbWizard_, _desid_)

_expression_ An expression that returns a **[Documents](Publisher.Documents.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_PbWizard_ |Optional| **[PbWizard](Publisher.PbWizard.md)** |The wizard to use to create the new publication. Can be a **PbWizard** constant.|
|_desid_ |Optional| **Long**|The ID of the design to apply to the new publication.|

## Return value

Document


## Remarks

The _desid_ parameter value should be the ID of the design to apply. You can determine the design ID by creating a new publication that uses the wizard and design that you want in the Publisher user interface and then running the following Visual Basic for Applications (VBA) macro.

```vb
Public Sub FindDesignID() 
 
 Dim pbWizard As Wizard 
 Dim pbWizardProperty As WizardProperty 
 
 Set pbWizard = ThisDocument.Wizard 
 Set pbWizardProperty = pbWizard.Properties(1) 
 
 Debug.Print pbWizardProperty.CurrentValueId 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]