---
title: ScopeFolder object (Office)
keywords: vbaof11.chm259000
f1_keywords:
- vbaof11.chm259000
ms.prod: office
api_name:
- Office.ScopeFolder
ms.assetid: fe46c1ad-fd60-a698-23dd-04d0631ac403
ms.date: 01/23/2019
localization_priority: Normal
---


# ScopeFolder object (Office)

Corresponds to a searchable folder. **ScopeFolder** objects are intended for use with the **[SearchFolders](office.searchfolders.md)** collection.


## Remarks

When you want to search specific folders, you can use the methods and properties of the **[SearchScope](office.searchscope.md)** object and **[ScopeFolders](office.scopefolders.md)** collection to retrieve **ScopeFolder** objects and add them to the **SearchFolders** collection.

In each **ScopeFolder** object, there is a **ScopeFolders** collection that contains the subfolders of the parent **ScopeFolder** object. You can traverse the entire folder structure of a search scope (for example, all local drives) by looping through these **ScopeFolders** collections and returning all of the lower-level **ScopeFolder** objects. A **ScopeFolder** object with no subfolders contains an empty **ScopeFolders** collection.

For an example that demonstrates how to loop through all of the **ScopeFolder** objects in a search scope, see the **SearchFolders** collection topic.

You can use the **[Add](office.searchfolders.add.md)** method of the **SearchFolders** collection to add a **ScopeFolder** object to the **SearchFolders** collection; however, it is usually simpler to use the **[AddToSearchFolders](office.scopefolder.addtosearchfolders.md)** method of the **ScopeFolder** that you want to add because there is only one **SearchFolders** collection for all searches.

For an example that demonstrates how to add a **ScopeFolder** to the **SearchFolders** collection, see the **SearchFolders** collection topic.


## Example

Use the **ScopeFolder** property of the **SearchScope** object to return the root **ScopeFolder** object of a search scope; for example:


```vb
Set sf = SearchScopes.Item(1).ScopeFolder
```

<br/>

Use the **[Item](office.scopefolders.item.md)** property of the **ScopeFolders** collection to return a subfolder of a root **ScopeFolder** object; for example:

```vb
Set sf = SearchScopes.Item(1).ScopeFolder.ScopeFolders.Item(1)
```

<br/>

The following example displays the root path of each directory in My Computer. To retrieve this information, the example first gets the **ScopeFolder** object at the root of My Computer. The path of this **ScopeFolder** object will always be "*". As with all **ScopeFolder** objects, the root object contains a **ScopeFolders** collection. This example loops through this **ScopeFolders** collection and displays the path of each **ScopeFolder** object in it. The paths of these **ScopeFolder** objects will be `A:\`, `C:\`, etc.

```vb
Sub DisplayRootScopeFolders() 
 
 'Declare variables that reference a 
 'SearchScope and a ScopeFolder object. 
 Dim ss As SearchScope 
 Dim sf As ScopeFolder 
 
 'Loop through the SearchScopes collection 
 'and display all of the root ScopeFolders collections in 
 'the My Computer scope. 
 For Each ss In SearchScopes 
 Select Case ss.Type 
 Case msoSearchInMyComputer 
 
 'Loop through each ScopeFolder object in 
 'the ScopeFolders collection of the 
 'SearchScope object and display the path. 
 For Each sf In ss.ScopeFolder.ScopeFolders 
 MsgBox "ScopeFolder object's path: " & sf.Path 
 Next sf 
 
 Case Else 
 End Select 
 Next 
 
End Sub
```


## See also

- [ScopeFolder object members](overview/Library-Reference/scopefolder-members-office.md)
- [Object Model Reference](overview/Library-Reference/reference-object-library-reference-for-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]