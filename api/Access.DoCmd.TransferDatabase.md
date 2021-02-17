---
title: DoCmd.TransferDatabase method (Access)
keywords: vbaac10.chm4188
f1_keywords:
- vbaac10.chm4188
ms.prod: access
api_name:
- Access.DoCmd.TransferDatabase
ms.assetid: 7eff4d0c-f660-72db-ee99-b6a3158f01de
ms.date: 03/07/2019
localization_priority: Priority
---


# DoCmd.TransferDatabase method (Access)

The **TransferDatabase** method carries out the TransferDatabase action in Visual Basic.


## Syntax

_expression_.**TransferDatabase** (_TransferType_, _DatabaseType_, _DatabaseName_, _ObjectType_, _Source_, _Destination_, _StructureOnly_, _StoreLogin_)

_expression_ A variable that represents a **[DoCmd](Access.DoCmd.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _TransferType_|Optional|**[AcDataTransferType](Access.AcDataTransferType.md)**|The type of transfer you want to make.|
| _DatabaseType_|Optional|**Variant**|A string expression that's the name of one of the types of databases that you can use to import, export, or link data. The _DatabaseType_ parameter is required for exporting and link data actions but not required for importing actions. The types or databases are:<ul><li><p>Microsoft Access (default)  </p></li><li><p>Jet 2.x</p></li><li><p>Jet 3.x</p></li><li><p>dBase III</p></li><li><p>dBase IV</p></li><li><p>dBase 5.0</p></li><li><p>Paradox 3.x</p></li><li><p>Paradox 4.x</p></li><li><p>Paradox 5.x</p></li><li><p>Paradox 7.x</p></li><li><p>ODBC Database</p></li><li><p>WSS (SharePoint)</p></li></ul>|
| _DatabaseName_|Optional|**Variant**|A string expression that's the full name, including the path (for WSS, Windows SharePoint Services, the URL), of the database that you want to use to import, export, or link data.|
| _ObjectType_|Optional|**[AcObjectType](Access.AcObjectType.md)**|The type of object to import or export.|
| _Source_|Optional|**Variant**|A string expression that's the name of the object whose data you want to import, export, or link.|
| _Destination_|Optional|**Variant**|A string expression that's the name of the imported, exported, or linked object in the destination database.|
| _StructureOnly_|Optional|**Variant**|Use **True** (1) to import or export only the structure of a database table. Use **False** (0) to import or export the structure of the table and its data. If you leave this argument blank, the default (**False**) is assumed.|
| _StoreLogin_|Optional|**Variant**|Use **True** to store the sign-in identification (ID) and password for an ODBC database in the connection string for a linked table from the database. If you do this, you don't have to sign in each time you open the table. Use **False** if you don't want to store the sign-in ID and password. If you leave this argument blank, the default (**False**) is assumed. This argument is available only in Visual Basic.|

## Remarks

You can use the TransferDatabase action to import or export data between the current Microsoft Access database or Access project (.adp) and another database. For Access databases, you can also link a table to the current Access database from another database. With a linked table, you have access to the table's data while the table itself remains in the other database.

You can import and export tables between Access and other types of databases. You can also export Access select queries to other types of databases. Access exports the result set of the query in the form of a table. You can import and export any Access database object if both databases are Access databases.

If you import a table from another Access database that's a linked table in that database, it will still be linked after you import it. That is, the link is imported, not the table itself.

The administrator of an ODBC database can disable the feature provided by the _SaveLoginId_ argument, requiring all users to enter the sign-in ID and password each time they connect to the ODBC database.

> [!NOTE] 
> You can also use ActiveX Data Objects (ADO) to create a link by using the **ActiveConnection** property for the **Recordset** object.

## Example

The following example imports the Monthly Sales Report from the Access database Northwind.accdb into the Corporate Sales Report in the current database.

```vb
DoCmd.TransferDatabase acImport, "Microsoft Access", _ 
    "C:\Users\Public\Northwind.accdb", acReport, "Monthly Sales Report", _ 
    "Corporate Sales Report"
```

<br/>

The following example links the ODBC database table Authors to the current database.

```vb
DoCmd.TransferDatabase acLink, "ODBC Database", _ 
    "ODBC;DSN=DataSource1;UID=User2;PWD=www;LANGUAGE=us_english;" & _ 
    "DATABASE=pubs", acTable, "Authors", "dboAuthors"
```

<br/>

The following example imports a list in SharePoint to a table in the current database:

```vb
DoCmd.TransferDatabase acImport, "WSS", _
    "WSS;DATABASE=https://company-my.sharepoint.com/personal/username_domain_com/express;" & _
    "LIST=NameOfListToImport;RetrieveIds=Yes", _
    acTable, , "NameOfLocalTable", False
```

<br/>

The following example exports a table in the current database to a list in SharePoint:

```vb
DoCmd.TransferDatabase acExport, "WSS", _
    "https://company-my.sharepoint.com/personal/username_domain_com/express", _
    acTable, "NameOfLocalTable", "NameOfListInSharePoint", False
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
