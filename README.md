<html lang="en" xmlns="http://www.w3.org/1999/xhtml">

<img align="left" src="Images/ReadMe/App.png" width="64px" >

# Microsoft Access Examples
Various examples of VBA, queries, macros, forms, reports and ribbon XML in an Microsoft Access database file

<!--[![Donate](https://img.shields.io/badge/Donate-PayPal-green.svg)](https://www.paypal.me/AnthonyDuguid/1.00)-->
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE "MIT License Copyright © Anthony Duguid")
![current_build Office_2016](https://img.shields.io/badge/current_build-Office_2016-red.svg)
[![Latest Release](https://img.shields.io/github/release/Access-projects/Access-examples.svg?label=latest%20release)](https://github.com/Access-projects/Access-examples/releases)
[![Github commits (since latest release)](https://img.shields.io/github/commits-since/Access-projects/Access-examples/latest.svg)](https://github.com/Access-projects/Access-examples/commits/master)
[![GitHub issues](https://img.shields.io/github/issues/Access-projects/Access-examples.svg)](https://github.com/Access-projects/Access-examples/issues)
<!--[![Github All Releases](https://img.shields.io/github/downloads/Access-projects/Access-examples/total.svg)](https://github.com/Access-projects/Access-examples/releases)-->

## Table of Contents
 - <a href="#references">References</a>
 - <a href="#cmd-line">Command Line Options</a>
 - <a href="#object-list">Object Listing Reference</a>

<a id="user-content-references" class="anchor" href="#references" aria-hidden="true"> </a>
### References
|Link                        |Type                 |
|:-------------------------------|:--------------------------|
|[Microsoft Access Find & Replace Add-in](http://www.rickworld.com/products.html)|Software|
|[Microsoft Access Merge & Diff](http://www.accdbmerge.net/download)|Software|
|[O'Reilly Access Database Design & Programming, 3rd Edition](http://shop.oreilly.com/product/9780596002732.do)|Book|

<a id="user-content-cmd-line" class="anchor" href="#cmd-line" aria-hidden="true"> </a>

<kbd>
<table>
    <caption>
        <h2>
            Command Line Options
        </h2>
    </caption>
<tr>
<th>OPTION</th>
<th>DESCRIPTION</th>
</tr>
<tr> 
<td>/decompile</td> 
<td>Undocumented command line. Will sometimes remove old code and objects from the database and make it faster.</td>
</tr> 
<tr> 
<td>/excl</td> 
<td>Opens the specified Access database for exclusive access. To open the database for shared access in a multiuser environment, omit this option. Applies to Access databases only.</td>
</tr> 
<tr> 
<td>/ro</td> 
<td>Opens the specified Access database or Access project for read-only access.</td>
</tr> 
<tr> 
<td>/user user name</td> 
<td>Starts Access by using the specified user name. Applies to Access databases only.</td>
</tr> 
<tr> 
<td>/pwd password</td> 
<td>Starts Access by using the specified password. Applies to Access databases only.</td>
</tr> 
<tr> 
<td>/profile user profile</td> 
<td>Starts Access by using the options in the specified user profile instead of the standard Windows Registry settings created when you installed Microsoft Access. This replaces the /ini option used in versions of Microsoft Access prior to Access 97 to specif</td>
</tr> 
<tr> 
<td>/compact target database or target Access project</td> 
<td>Compacts and repairs the Access database, or compacts the Access project that was specified before the /compact option, and then closes Access. If you omit a target file name following the /compact option, the file is compacted to the original name and fo</td>
</tr> 
<tr> 
<td>/repair</td> 
<td>Repairs the Access database that was specified before the /repair option, and then closes Microsoft Access. In Microsoft Access 2000 or later, compact and repair functionality is combined under /compact. The /repair option is supported for backward compat</td>
</tr> 
<tr> 
<td>/convert target database</td> 
<td>Converts a previous-version Access database or Access project to the default file format, renames the new file, and then closes Access. You must specify the source database before you use the /convert option. To view the default file format, click Options</td>
</tr> 
<tr> 
<td>/x macro</td> 
<td>Starts Access and runs the specified macro. Another way to run a macro when you open a database is to use an AutoExec macro.</td>
</tr> 
<tr> 
<td>/cmd</td> 
<td>Specifies that what follows on the command line is the value that will be returned by the Command function. This option must be the last option on the command line. You can use a semicolon (;) as an alternative to /cmd. 
Use this option to specify a comma</td>
</tr> 
<tr> 
<td>/nostartup</td> 
<td>Starts Access without displaying the task pane (the second dialog box that you see when you start Access).</td>
</tr> 
<tr> 
<td>/wrkgrp workgroup information file</td> 
<td>Starts Access by using the specified workgroup information file.  Applies to Access databases only.

C:\Program Files\Access 2K Runtime\Office\MSACCESS.EXE P:\general\Database\Facilities_Management_System\v_11_1\FMS.mde /user username /pwd /wrkgrp P:\gene</td>
</tr> 
<tr> 
<td>/runtime</td> 
<td>Starts Access in runtime mode for testing.  C:\Program Files\Office\MSACCESS.EXE C:\Database\database.mde /runtime</td>
</tr> 
</table>
</kbd>

<br>
<br>
<br>

<a id="user-content-object-list" class="anchor" href="#object-list" aria-hidden="true"> </a>

<kbd>
<table>
    <caption>
        <h2>
            Object Listing Reference
        </h2>
    </caption>
<tr>
<th>Type</th>
<th>Nbr</th>
<th>Name</th>
<th>Description</th>
<th>Items</th>
</tr>
<tr> 
<td>Form (-32768)</td> 
<td></td> 
<td></td> 
<td></td> 
<td></td>
</tr> 
<tr> 
<td></td> 
<td>1</td> 
<td>001_About_frm</td> 
<td>A description of the Microsoft Access Database.</td> 
<td></td>
</tr> 
<tr> 
<td></td> 
<td>2</td> 
<td>002_Splash_Screen_frm</td> 
<td>A splash screen with the company logo and the application information.</td> 
<td></td>
</tr> 
<tr> 
<td></td> 
<td>3</td> 
<td>003_Buttons_frm</td> 
<td>Command Button example with code</td> 
<td></td>
</tr> 
<tr> 
<td></td> 
<td>4</td> 
<td>004_DSN_Password_frm</td> 
<td>A form to login to a DSN database connection</td> 
<td></td>
</tr> 
<tr> 
<td></td> 
<td>5</td> 
<td>005_Access_Security_frm</td> 
<td>Allow the user to login, change password, edit permission; where granted, and view list of current users.</td> 
<td></td>
</tr> 
<tr> 
<td></td> 
<td>6</td> 
<td>006_Dialogbox_Examples_frm</td> 
<td>Dialog box examples</td> 
<td></td>
</tr> 
<tr> 
<td></td> 
<td>7</td> 
<td>010_Date_Calendar_frm</td> 
<td>Uses the calendar dll from Microsoft (only needed previous to 2007)</td> 
<td></td>
</tr> 
<tr> 
<td></td> 
<td>8</td> 
<td>301_Roulette_frm</td> 
<td>Roulette game</td> 
<td></td>
</tr> 
<tr> 
<td></td> 
<td>9</td> 
<td>700_Create_Filelist_frm</td> 
<td>Recursive code to go through complete folder structure and catalog a file list</td> 
<td></td>
</tr> 
<tr> 
<td></td> 
<td>10</td> 
<td>800_Link_Manager_frm</td> 
<td>Manage the links from server objects</td> 
<td></td>
</tr> 
<tr> 
<td></td> 
<td>11</td> 
<td>900_Check_Runtime_frm</td> 
<td>Check performance of queries and reports</td> 
<td></td>
</tr> 
<tr> 
<td></td> 
<td>12</td> 
<td>950_Leszynski_Conventions_frm</td> 
<td>Leszynski Naming Conventions for Microsoft Solution Developers</td> 
<td></td>
</tr> 
<tr> 
<td></td> 
<td>13</td> 
<td>980_Data_Dictionary_frm</td> 
<td>A form that documents the columns, column descriptions, and the rest of the attributes of each table within the database whether the table is a Access or Server based.</td> 
<td></td>
</tr> 
<tr> 
<td></td> 
<td></td> 
<td></td> 
<td></td> 
<td></td>
</tr> 
<tr> 
<td>Macro (-32766)</td> 
<td></td> 
<td></td> 
<td></td> 
<td></td>
</tr> 
<tr> 
<td></td> 
<td>1</td> 
<td>AutoKeys</td> 
<td>Hot Key designation</td> 
<td></td>
</tr> 
<tr> 
<td></td> 
<td>2</td> 
<td>Ribbon</td> 
<td>Access 2007-2016 Ribbon example</td> 
<td></td>
</tr> 
<tr> 
<td></td> 
<td></td> 
<td></td> 
<td></td> 
<td></td>
</tr> 
<tr> 
<td>Report (-32764)</td> 
<td></td> 
<td></td> 
<td></td> 
<td></td>
</tr> 
<tr> 
<td></td> 
<td>1</td> 
<td>001_Organization_Logo_srp</td> 
<td>Primary company logo.</td> 
<td></td>
</tr> 
<tr> 
<td></td> 
<td>2</td> 
<td>099_Object_Listing_Report_rpt</td> 
<td>A listing of all the objects within the database</td> 
<td></td>
</tr> 
<tr> 
<td></td> 
<td>3</td> 
<td>803_Object_Calendar_List_rpt</td> 
<td>A calendar style report</td> 
<td></td>
</tr> 
<tr> 
<td></td> 
<td>4</td> 
<td>803a_Calendar_List_srp</td> 
<td>A sub-report to 803_Object_Calendar_List_rpt</td> 
<td></td>
</tr> 
<tr> 
<td></td> 
<td>5</td> 
<td>804_Weekly_Report_rpt</td> 
<td>A 52 week style report</td> 
<td></td>
</tr> 
<tr> 
<td></td> 
<td>6</td> 
<td>805_Timeline_Report_rpt</td> 
<td>A MS Project like time line report example</td> 
<td></td>
</tr> 
<tr> 
<td></td> 
<td>7</td> 
<td>950_Leszynski_Conventions_rpt</td> 
<td>Leszynski Naming Conventions for Microsoft Solution Developers</td> 
<td></td>
</tr> 
<tr> 
<td></td> 
<td>8</td> 
<td>980_Data_Dictionary_rpt</td> 
<td>A data dictionary listing of all the tables in this database</td> 
<td></td>
</tr> 
<tr> 
<td></td> 
<td></td> 
<td></td> 
<td></td> 
<td></td>
</tr> 
<tr> 
<td>Module (-32761)</td> 
<td></td> 
<td></td> 
<td></td> 
<td></td>
</tr> 
<tr> 
<td></td> 
<td>1</td> 
<td>clsMonthCal</td> 
<td>Calendar Class</td> 
<td>CreateDTPControl
<br>
ReDraw
<br>
MultiSelect
<br>
MaxSelectRangeofDays
<br>
CalendarYOffset
<br>
MonthRows
<br>
MonthColumns
<br>
WindowLocation
<br>
PositionAtCursor
<br>
CursorXinit
<br>
CursorX
<br>
CursorY
<br>
FontSize
<br>
FontName
<br>
ShowWeekNumbers
<br>
NoTodayCircle
<br>
NoToday
<br>
hwnd
hWndForm
<br>
hWndCal
<br>
OneClick
<br>
SelectedDate
<br>
SetSelectedDateRange
<br>
StartSelectedDate
<br>
EndSelectedDate
<br>
GetSelectedDates
<br>
SetViewableMonths
<br>
SetBoldDayState
<br>
Class_Initialize
<br>
Class_Terminate
<br>
GetProperty
<br>
SetProperty
<br>
IsCalendar
</td>
</tr> 
<tr> 
<td></td> 
<td>2</td> 
<td>modCalendar</td> 
<td>Calendar Functions</td> 
<td>ShowMonthCalendar
<br>
WindowProc
<br>
GetFuncPtr
<br>
SetSelectedDate
<br>
SetMonths
<br>
sClick
<br>
ShowWeekNums
<br>
sShowToday
<br>
sShowcircleToday
<br>
sWindowPosition
<br>
UpdateCursor
<br>
LocationCursorOnCalendar
<br>
ReleaseClass
<br>
LoWord
<br>
MakeDWord
</td>
</tr> 
<tr> 
<td></td> 
<td>3</td> 
<td>modDate</td> 
<td>Date functions</td> 
<td>Timedelay
<br>
DaysInMonth
<br>
Age
<br>
DaysInMonthMS
<br>
DaysInMonth2
<br>
EndOfMonth
<br>
EndOfWeek
<br>
FormatInterval
<br>
LastBusDay
<br>
LeapYear2
<br>
NextDay
<br>
NextDay1
<br>
Num2Date
<br>
PriorDay
<br>
PriorDay1
<br>
StartOfMonth
<br>
StartOfWeek
<br>
String2Date
<br>
WorkDays
</td>
</tr> 
<tr> 
<td></td> 
<td>4</td> 
<td>modFileFunctions</td> 
<td>Contains procedures that check and refresh the links to Northwind tables.</td> 
<td>RefreshLinks
<br>
GetDirectory
<br>
FindFile
<br>
ReturnAllFiles
<br>
RecursiveDir
<br>
TrailingSlash
<br>
MSA_CreateFilterString
<br>
MSA_ConvertFilterString
<br>
MSA_GetSaveFileName
<br>
MSA_SimpleGetSaveFileName
<br>
MSA_GetOpenFileName
<br>
MSA_SimpleGetOpenFileName
<br>
OF_to_MSAOF
<br>
MSAOF_to_OF
</td>
</tr> 
<tr> 
<td></td> 
<td>5</td> 
<td>modMain</td> 
<td>Subroutines that run objects and procedures in the database</td> 
<td>OSUserID
<br>
ErrorMsg
<br>
GetObjectDescription
<br>
GetObjectTypeName
</td>
</tr> 
<tr> 
<td></td> 
<td>6</td> 
<td>modMathAreaVolume</td> 
<td>Area and Volume</td> 
<td>ACircle
<br>
ARect
<br>
ARing
<br>
ASphere
<br>
ASquare
<br>
ASquare2
<br>
ATrap
<br>
ATriangle2
<br>
RectDiag
<br>
SquareDiag
<br>
VCone
<br>
VCylinder
<br>
VPipe
<br>
VPyramid
<br>
VTruncPyramid
</td>
</tr> 
<tr> 
<td></td> 
<td>7</td> 
<td>modMathCumulative</td> 
<td>Cumulative Value</td> 
<td>CumulativeValue
<br>
DeCumulativeValue
</td>
</tr> 
<tr> 
<td></td> 
<td>8</td> 
<td>modMathStatistics</td> 
<td>Statistics</td> 
<td>Combin
<br>
Factorial
<br>
FactorialR
<br>
Permut
<br>
Regress
<br>
TestRegress
</td>
</tr> 
<tr> 
<td></td> 
<td>9</td> 
<td>modMathTrig</td> 
<td>Trigonometry</td> 
<td>ArcCos
<br>
Arccosec
<br>
Arccotan
<br>
ArcSec
<br>
ArcSin
<br>
ATan2
<br>
Cotan
<br>
Deg2Rad
<br>
HArccos
<br>
HArccosec
<br>
HArcsec
<br>
HArcsin
<br>
HArctan
<br>
HCos
<br>
HCosec
<br>
HSec
<br>
HSin
<br>
HTan
<br>
pi
<br>
Rad2Deg
<br>
Sec
</td>
</tr> 
<tr> 
<td></td> 
<td>10</td> 
<td>modMathXYZ</td> 
<td>Latitude/Longitude</td> 
<td>DegToDMS
<br>
DegToDMSStr
<br>
DMSStrToDeg
<br>
DMSToDeg
<br>
GreatArcDistance
<br>
LatLongToXYZ
<br>
testxyz
<br>
XYZToLatLong
</td>
</tr> 
<tr> 
<td></td> 
<td>11</td> 
<td>modOutlook</td> 
<td>Microsoft Outlook Functions</td> 
<td>ImportContactsFromOutlook
<br>
ImportCalendarFromOutlook
<br>
PushAppointments
</td>
</tr> 
<tr> 
<td></td> 
<td>12</td> 
<td>modString</td> 
<td>String Manipulation/Parsing</td> 
<td>LowerCC
<br>
ParseCSZ
<br>
ParseName
<br>
Proper
<br>
ProperEx
<br>
ProperWord
<br>
StrToHex
<br>
TestParseName
<br>
UpperCC
<br>
CountCSVWords
<br>
CountWords
<br>
CutFirstWord
<br>
CutLastWord
<br>
ReplaceStr
<br>
GetCSVWord
<br>
GetWord
<br>
Like2
<br>
LPad
<br>
ParseItemsToArray
<br>
RPad
<br>
ParseArticle
</td>
</tr> 
<tr> 
<td></td> 
<td></td> 
<td></td> 
<td></td> 
<td></td>
</tr> 
<tr> 
<td>Access Table (1)</td> 
<td></td> 
<td></td> 
<td></td> 
<td></td>
</tr> 
<tr> 
<td></td> 
<td>1</td> 
<td>CMD_LINE_TB</td> 
<td>Command line options for Access</td> 
<td></td>
</tr> 
<tr> 
<td></td> 
<td>2</td> 
<td>CUM_VAL_TB</td> 
<td>Used to test the function in modMathCumulative</td> 
<td></td>
</tr> 
<tr> 
<td></td> 
<td>3</td> 
<td>DATABASE_STRUCTURE_TB</td> 
<td>Used to store the database table structure</td> 
<td></td>
</tr> 
<tr> 
<td></td> 
<td>4</td> 
<td>DECUM_VAL_TB</td> 
<td>Used to test the function in modMathCumulative</td> 
<td></td>
</tr> 

<tr> 
<td></td> 
<td>5</td> 
<td>ROULETTE_TB</td> 
<td>Numbers for roulette</td> 
<td></td>
</tr> 
<tr> 
<td></td> 
<td>6</td> 
<td>TAG_GRP_TB</td> 
<td>Leszynski Naming Conventions for Microsoft Solution Developers</td> 
<td></td>
</tr> 
<tr> 
<td></td> 
<td>7</td> 
<td>TAG_NME_TB</td> 
<td>Leszynski Naming Conventions for Microsoft Solution Developers</td> 
<td></td>
</tr> 
<tr> 
<td></td> 
<td>8</td> 
<td>tblContacts</td> 
<td>Used with the Microsoft Outlook Functions</td> 
<td></td>
</tr> 
<tr> 
<td></td> 
<td>9</td> 
<td>tblDefaults</td> 
<td>Used to store defaults used in the this database</td> 
<td></td>
</tr> 
<tr> 
<td></td> 
<td>10</td> 
<td>tblFileList</td> 
<td>Used to store a list of files</td> 
<td></td>
</tr> 
<tr> 
<td></td> 
<td>11</td> 
<td>USysRibbons</td> 
<td>Microsoft Access Ribbon XML example</td> 
<td></td>
</tr> 
<tr> 
<td></td> 
<td></td> 
<td></td> 
<td></td> 
<td></td>
</tr> 
<tr> 
<td>Query (5)</td> 
<td></td> 
<td></td> 
<td></td> 
<td></td>
</tr> 
<tr> 
<td></td> 
<td>1</td> 
<td>101_Object_List_qry</td> 
<td>Creates a list of all objects in the MS Access database (Needs module modMain)</td> 
<td></td>
</tr> 
<tr> 
<td></td> 
<td>2</td> 
<td>200_Create_DATABASE_STRUCTURE_TB_qry</td> 
<td>Used to create the MS Access table using DDL</td> 
<td></td>
</tr> 
<tr> 
<td></td> 
<td>3</td> 
<td>201_Create_DATABASE_STRUCTURE_TB_qry</td> 
<td>Used to create the MS Access table using DDL</td> 
<td></td>
</tr> 
<tr> 
<td></td> 
<td>4</td> 
<td>202_Add_Constraint_pk001_Primary_Key_PROP_ID_qry</td> 
<td>Used to add a constraint to a MS Access table using DDL</td> 
<td></td>
</tr> 
<tr> 
<td></td> 
<td>5</td> 
<td>203_Add_Column_LST_MDFD_ID_qry</td> 
<td>Used to add a new column to a MS Access table using DDL</td> 
<td></td>
</tr> 
<tr> 
<td></td> 
<td>6</td> 
<td>204_Alter_Column_LST_MDFD_ID_qry</td> 
<td>Used to alter a column in a MS Access table using DDL</td> 
<td></td>
</tr> 
<tr> 
<td></td> 
<td>7</td> 
<td>205_Drop_DATABASE_STRUCTURE_TB_qry</td> 
<td>Used to drop a MS Access table using DDL</td> 
<td></td>
</tr> 
<tr> 
<td></td> 
<td>8</td> 
<td>720_Cumulative_Value_qry</td> 
<td>Used to test the function in modMathCumulative</td> 
<td></td>
</tr> 
<tr> 
<td></td> 
<td>9</td> 
<td>721_Decumulative_Value_qry</td> 
<td>Used to test the function in modMathCumulative</td> 
<td></td>
</tr> 
<tr> 
<td></td> 
<td>10</td> 
<td>803_Calendar_Style_Listing_qry</td> 
<td>Used in the report 803_Object_Calendar_List_rpt</td> 
<td></td>
</tr> 
<tr> 
<td></td> 
<td>11</td> 
<td>803a_Calendar_List_qry</td> 
<td>Used in the sub-report 803a_Calendar_List_srp</td> 
<td></td>
</tr> 
<tr> 
<td></td> 
<td>12</td> 
<td>804_Weekly_Report_qry</td> 
<td>Used in the report 804_Weekly_Report</td> 
<td></td>
</tr> 
<tr> 
<td></td> 
<td>13</td> 
<td>940_Date_Format_qry</td> 
<td>Examples of formating date & time</td> 
<td></td>
</tr> 
<tr> 
<td></td> 
<td>14</td> 
<td>950_Leszynski_Conventions_qry</td> 
<td>Leszynski Naming Conventions for Microsoft Solution Developers</td> 
<td></td>
</tr> 
<tr> 
<td></td> 
<td>15</td> 
<td>SQL_Server_Function_Login_Name_qry</td> 
<td>SQL Server pass-through query</td> 
<td></td>
</tr> 
<tr> 
<td></td> 
<td>16</td> 
<td>SQL_Server_View_Date_Time_qry</td> 
<td>SQL Server pass-through query</td> 
<td></td>
</tr> 
</table>
</kbd>
</html>
