---
title: Permission.DocumentAuthor property (Office)
keywords: vbaof11.chm261013
f1_keywords:
- vbaof11.chm261013
ms.prod: office
api_name:
- Office.Permission.DocumentAuthor
ms.assetid: d756c476-8adf-a302-9356-e491b0ae9bf7
ms.date: 01/22/2019
localization_priority: Normal
---


# Permission.DocumentAuthor property (Office)

Gets or sets the name in email form of the author of the active document. Read/write.


## Syntax

_expression_.**DocumentAuthor**

_expression_ A variable that represents a **[Permission](Office.Permission.md)** object.


## Remarks

The **DocumentAuthor** property returns or sets the author of the active document. The author always has non-expiring owner rights to the document, whether owner permission is granted explicitly (through a **[UserPermission](Office.userPermission.md)** object) or not.

The **DocumentAuthor** property can only be changed to a different account that has been certified through the permissions user interface to open restricted content on the local computer. In most cases, users who have a single Windows account can only choose between their Windows and Passport identities.

If the user's Windows and Passport identities use the same email address, use the format `passport:someone@example.com` to specify the Passport identity as the **DocumentAuthor** property.


## Example

The following example displays information about the permissions settings of the active document, including the document author.


```vb
 Dim irmPermission As Office.Permission 
 Dim strIRMInfo As String 
 Set irmPermission = ActiveWorkbook.Permission 
 If irmPermission.Enabled Then 
 strIRMInfo = "Permissions are enabled on this document." & vbCrLf 
 strIRMInfo = strIRMInfo & " View in trusted browser: " & _ 
 irmPermission.EnableTrustedBrowser & vbCrLf & _ 
 " Document author: " & irmPermission.DocumentAuthor & vbCrLf & _ 
 " Users with rights: " & irmPermission.Count & vbCrLf & _ 
 " Cache licenses locally: " & irmPermission.StoreLicenses & vbCrLf & _ 
 " Request permission URL: " & irmPermission.RequestPermissionURL & vbCrLf 
 If irmPermission.PermissionFromPolicy Then 
 strIRMInfo = strIRMInfo & " Permissions applied from policy:" & vbCrLf & _ 
 " Policy name: " & irmPermission.PolicyName & vbCrLf & _ 
 " Policy description: " & irmPermission.PolicyDescription 
 Else 
 strIRMInfo = strIRMInfo & " Default permissions applied." 
 End If 
 Else 
 strIRMInfo = "Permissions are NOT enabled on this document." 
 End If 
 MsgBox strIRMInfo, vbInformation + vbOKOnly, "IRM Information" 
 Set irmPermission = Nothing 

```


## See also

- [Permission object members](overview/library-reference/permission-members-office.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]