<div align="center">

## Export from any Database to any database in code\!


</div>

### Description

This article explains how to, using pure code, without having to reference Excel, Access, dBase libraries, etc, export data from one format to another.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Jamey Church](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jamey-church.md)
**Level**          |Intermediate
**User Rating**    |4.7 (42 globes from 9 users)
**Compatibility**  |VB 5\.0, VB 6\.0, VB Script, ASP \(Active Server Pages\) , VBA MS Access, VBA MS Excel
**Category**       |[Databases/ Data Access/ DAO/ ADO](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases-data-access-dao-ado__1-6.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/jamey-church-export-from-any-database-to-any-database-in-code__1-31286/archive/master.zip)





### Source Code

<br><Br><br>:::Updated 2/1/02:::<B>I have received some corrected information thanks to Stephen Kent. This project requires MDAC, which apparently contains the Office (Excel, Access, etc) Libraries in it. Thus technically you do need the libraries/drivers for this to work, but you don't need the whole programs. To get Jet / MDAC, go to microsoft.com and search for MDAC 2.5 (I think 2.6 still has Jet...) or just MDAC, then Jet and download and install both.</b><br><br>
:::Updated 1/31/02:::<B>NOTICE: <i>Some</i> of the DAO code originated from <a href="http://www.smith-voice.com">Smith-Voice.com</a>. All other code is from one of my programs. <br><br>
Also note that you do need MDAC AND MS Jet installed for THIS example to work properly.(MSJet is in either 2.6 or 2.5 or lower (can't remember for sure if 2.6 includes it). Or you can download them seperately from the MS Website.) You may also install/use a different provider by modifying the "Provider=" part of the SQL Statements. You do not need MS Office/Excel/etc if you have an appropriate Provider/Driver(i.e. - Jet) installed and MDAC.</b>
<br><br>
If you have Excel 97, you will use the "Excel 8.0" definition in the SQL "INTO" Statements below. Excel 2000 is 9.0, and I believe 2002(XP) is 10.0. Replace "Excel 8.0" in the example code with the database name you wish to export to (I.E. - "Access 8.0", "Access 9.0", "dBase III", etc). Also in the project you must "Reference" (not add a component) One of the following:
<br><Br>
<u>For DAO:</u><br><br>
Microsoft DAO 3.6 Object Library <br>
Microsoft DAO 3.51 Object Library <br><br>
<u>For ADO:</u><br><br>
Microsoft Activex Data Objects 2.1(or higher) Library <br>
Microsoft ADO Ext 2.6 for DDL & Security <br><br>
Use one of the ADO/DAO references depending on what method you use. You should only need one. <br><br>
You need these for the ADO and DAO examples here. <br><br>
<br><Br>
Here is some code that illustrates how to export, for example, from Access to Excel without having to have either product installed on your computer, or the libraries.
<br><Br>
(This example is ADO. you can also do this in code with DAO, altho it is an outdated method.)
Create a new project with a command button and a DAO (or ADO2.1sp2 & higher) reference, then copy this code to the button's Click event. This assumes there is a database at C:\WINDOWS\Desktop\Master Database\master.mdb.
<br><br><Br>
=========================================<br>
'<br>
'This will demonstrate how to export from <br>
'MS Access to Excel, without either product <br>
'installed. This example uses ADO. It can also be <br>
'done in DAO with a little modification.<br>
'NOTE: Must have ADO(Activex Data Objects)<br>
'or DAO Referenced to use these examples<br>
<br><br>
'Define the variables/objects <br>
Dim conn as new ADODB.Connection <br>
Dim SQL, ConnectString as String <br><br>
'Assign SQL And ConnectString --- Notice the <br>
'SQL "INTO" Statement---explained below: <br><br>
SQL = "SELECT * INTO [Excel 8.0;DATABASE=F:\Master\Exported.xls].[Master] From [Complexes]" <br><br>
CS = "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=C:\WINDOWS\Desktop\Master Database\master.mdb" <br><br>
'Open the Connection and export the database. <br>
conn.Open ConnectString <br>
conn.Execute SQL <br><br>
'Close the connection.<br>
conn.close<br>
set conn = Nothing<br><br>
======================================<br><br>
This example uses DAO:<br><br>
======================================<br><br>
Dim db as database<br><br>
On Error Resume Next<br><br>
Set db = Workspaces(0).OpenDatabase("C:\WINDOWS\Desktop\Master Database\master.mdb") <br><br>
db.Execute "SELECT * INTO [dBase III;DATABASE=C:\My Documents].[testb] FROM [Authors]" <br><br>
If Err.Number <> 0 then 'Always check this!!!<br>
 Msgbox Err.Number & vbcr & Err.Description<br>
End If <br><br>
=========================================<br><br>
The generic layout for the SQL Statement:<br><br>
SELECT tbl.fields INTO
[dbms specifier;DATABASE="path"].[unqualified
filename; may be tablename or sheetname(in excel)] FROM [table or tables] <br><br>
=========================================<br><br>
Using the brackets and dot operator, you get a proper output in the database type of your choice. You can customize the SQL statement to your needs. (Such as ordering, limiting to certain columns, adding columns, etc.) For some SQL help, go to <A href="www.visualbasicforum.com">Visual Basic Forum.com</a>.<br><br>
========================================<br><br>
<u>{Explanation of the ADO Example (Applies to DAO Example also)}</u><br><br>
The important part of this, which does the exporting is the SQL Statement. The "INTO [Excel 8.0;DATABASE=F:\Master\Exported.xls].[Master]" is the part that does all the work. You can replace "Excel 8.0" with other databases such as dBase III, Access 8.0, etc. DATABASE="" is where you specify what file to export to, and if exporting to Excel, for example, the .[Master] is the sheet name to export to in the workbook. With Access it would refer to a table name. Every section of this part of the SQL Statement is required. Theoretically you could export to any database, altho I have not tested this beyond Access->Excel myself.
<br><br>
I hope this will help some of you who do not wish to install Access/Excel/dBase III on every clients computer just for your program to run, as it has been a GREAT help to me.
FYI: The reasoning behind MS etc not making any fuss about this method is simple. They want you to have to install their software to perform this kind of functionality. In fact they say it is NOT possible. Well, now you know that it IS possible :)<br><br>

