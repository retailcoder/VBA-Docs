---
title: Options.CommentsColor property (Word)
keywords: vbawd10.chm162988486
f1_keywords:
- vbawd10.chm162988486
ms.prod: word
api_name:
- Word.Options.CommentsColor
ms.assetid: 5c2861d6-7933-3e77-ba55-c7bfd174f44a
ms.date: 06/08/2017
localization_priority: Normal
---


# Options.CommentsColor property (Word)

Returns or sets a  **WdColorIndex** constant that represents the color of comments in a document. Read/write.

> [!IMPORTANT]
> This property has changed. The `Options.CommentsColor` property is still available but there will be no visible effect in the Redesigned Comments experience. However, the command will apply the `CommentsColor` property to the Word options so if the current user reverts to the previous commenting experience, the comment thread outline color will change based on the previous setting.

## Syntax

_expression_. `CommentsColor`

_expression_ Required. A variable that represents an **[Options](Word.Options.md)** object.


## Example

This example sets the global option for Microsoft Word to color comments made in documents according to the author of the comment.


```vb
Sub ColorCodeComments() 
 Options.CommentsColor = wdByAuthor 
End Sub
```


## See also


[Options Object](Word.Options.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
