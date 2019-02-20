# IDEA-SQL-Export-History
A SQL and Visual-Basic script to export the SQL data of a IDEA project



Option Explicit
'create a type to hold the entire overview table from the database
Type project
	TaskName As String
	DateTime As String
	UserName As String
	IDEAScript As String
	HistoryLog As String
	DataBaseGUID As String
	TaskGUID As String
	RecordGUID As String
	AllRecordsUsed As String
	TaskType As String
	TaskStream As String
	Filename As String
	SubFolder As String
	Unsupported As String
	ProjectName As String
	Deleted As String
	
		
	ParentGUID As String
	ChildGUID As String

	
End Type
	
Sub History
	Dim objConn As Object
	Dim connStr As String
	Dim rs As Object
	Dim i As Integer
	Dim ProjectInfo() As project
	Dim bFirstTime As Boolean
	
	
	Dim excel As Object
	Dim oBook As Object
	Dim oSheet As Object
	Dim t As Integer
	Dim x As Integer
	Dim count As Integer
	Dim Location  As String
	Dim strr As String
	Dim dstr As String
	Dim History_Data_Formatting As Object
	    Dim percentComplete As Object
Dim p As Integer


'-------------------------------------创建进度条---------------------------------------------

Set percentComplete = CreateObject ("CommonIdeaControls.StandaloneProgressCtl")
percentComplete.Start "Exporting, please wait......"
p = 20
    







'-------------------------------------得到当前项目的路径---------------------------------------------	
	Dim start As Double
	
	start = Timer
	
	Location = Client.WorkingDirectory()
	
'-------------------------------------复制Excel模板至Local Project Folder---------------------------------------------	
		
		strr = "C:\temp\History Template.xlsm"
		
		'strr = isplit( Location, "\", "",3,1) & "\Local Library\Macros.ILB\History Template.xlsm"
	
		'strr="R:\4.Fund Audit and KAAT\KAAT\Saint Xu\History_Automation\Excel Template\History Template.xlsm"
		
		'strr="C:\Users\saintxu\Desktop\KDC\近期任务\工具开发\History_Automation\HistoryDatabase\testing.xlsm"
		
		dstr= location  & isplit( Location, "", "\",2,1)& " - History.xlsm"

		MsgBox strr & dstr
				
		FileCopy strr, dstr

'-------------------------------------打开Excel（方案待完善）---------------------------------------------	

	Set excel = CreateObject("Excel.Application")
	excel.Visible = false
	
	Set oBook = excel.Workbooks.Open (dstr)
	t = 2

	bFirstTime = True
	
	ReDim ProjectInfo(0)
	i = 0

	
'-------------------------更新进度条add this code at the bottom of the loop to update the counter------------
		percentComplete.Start "Exporting, please wait......"
		percentComplete.Progress 40
	
	
'------------------------------------------循环读取Overview表内容并记录在Excel-------------------------------------------
'create the connection object to the database

Set objConn = CreateObject("ADODB.Connection")
		'create the connection string.  The Data source has to point to the sdf file of the project you want to extract the info

connStr = "Provider=Microsoft.SQLSERVER.CE.OLEDB.3.5;Data Source=" & Location  & "\ProjectOverview.sdf;"
		'connect to the database

objConn.open connStr
		'use SQL to access the information.  In this instance all the fields are accessed, I used the field names instead of using SELECT *.*

Set rs = objConn.execute("SELECT TaskName, DateTime, UserName, IDEAScript, HistoryLog, DatabaseGUID, TaskGUID, RecordGUID, AllRecordsUsed, TaskType, TaskStream, Filename, SubFolder, Unsupported, ProjectName, Deleted FROM Overview")'
			'loop through the table
			
Do While Not rs.EOF
				'increment the array to hold the informaiton
				
If Not bFirstTime Then
ReDim preserve ProjectInfo(UBound(ProjectInfo) + 1)
End If
				'populate the array with the information.
				
ProjectInfo(i).TaskName = rs.Fields("TaskName")
ProjectInfo(i).DateTime = rs.Fields("DateTime")
ProjectInfo(i).UserName = rs.Fields("UserName")
ProjectInfo(i).IDEAScript = rs.Fields("IDEAScript")
ProjectInfo(i).HistoryLog = rs.Fields("HistoryLog")
ProjectInfo(i).DataBaseGUID = rs.Fields("DatabaseGUID")
ProjectInfo(i).TaskGUID = rs.Fields("TaskGUID")
ProjectInfo(i).RecordGUID = rs.Fields("RecordGUID")
ProjectInfo(i).AllRecordsUsed = rs.Fields("AllRecordsUsed")
ProjectInfo(i).TaskType = rs.Fields("TaskType")
ProjectInfo(i).TaskStream = rs.Fields("TaskStream")
ProjectInfo(i).Filename = rs.Fields("Filename")
ProjectInfo(i).SubFolder = rs.Fields("SubFolder")
ProjectInfo(i).Unsupported = rs.Fields("Unsupported")
ProjectInfo(i).ProjectName = rs.Fields("ProjectName")
ProjectInfo(i).Deleted = rs.Fields("Deleted")			

'-------------------------更新进度条add this code at the bottom of the loop to update the counter------------

		'percentComplete.Start "Exporting, please wait......"
		percentComplete.Progress 50
		
		
											
'------------------------------------------导出数据库信息至Excel-------------------------------------------

Dim arr1(9999,16) As Variant


	arr1(i,0) = ProjectInfo(i).TaskName
	arr1(i,1)= ProjectInfo(i).DateTime
	arr1(i,2)= ProjectInfo(i).username
	arr1(i,3)= ProjectInfo(i).IDEAScript
	arr1(i,4)= ProjectInfo(i).HistoryLog
	arr1(i,5) = ProjectInfo(i).DataBaseGUID
	arr1(i,6) = ProjectInfo(i).TaskGUID
	arr1(i,7)= ProjectInfo(i).RecordGUID
	arr1(i,8) = ProjectInfo(i).AllRecordsUsed
	arr1(i,9)= ProjectInfo(i).Tasktype
	arr1(i,10) = ProjectInfo(i).Taskstream
	arr1(i,11)= ProjectInfo(i).Filename
	arr1(i,12)= ProjectInfo(i).subfolder
	arr1(i,13) = ProjectInfo(i).unsupported
	arr1(i,14) = ProjectInfo(i).projectname
	arr1(i,15) = ProjectInfo(i).deleted	

	
		t= t+1			
			
				i = i + 1
				'move to the next record.
				rs.MoveNext
			Loop
			
	oBook.Worksheets("ProjectOverview_overview_export").cells(2,1).resize(i,16).value = arr1

			
		Set rs = Nothing
	

	
	Set objConn = Nothing
	
'------------------------------------------Overview表内容读取完成，开始读取Parent表-------------------------------------------
	i=0
	t = 2

	'create the connection object to the database
	Set objConn = CreateObject("ADODB.Connection")
		'create the connection string.  The Data source has to point to the sdf file of the project you want to extract the info
		connStr = "Provider=Microsoft.SQLSERVER.CE.OLEDB.3.5;Data Source=" & Location  & "\ProjectOverview.sdf;"
		'connect to the database
		objConn.open connStr
		'use SQL to access the information.  In this instance all the fields are accessed, I used the field names instead of using SELECT *.*
		Set rs = objConn.execute("SELECT ChildGUID, ParentGUID FROM Parent")'
			'loop through the table
			Do While Not rs.EOF
				'increment the array to hold the informaiton
				If Not bFirstTime Then
					ReDim preserve ProjectInfo(UBound(ProjectInfo) + 1)
				End If
				'populate the array with the information.
	
				ProjectInfo(i).ChildGUID = rs.Fields("ChildGUID")			
				ProjectInfo(i).ParentGUID = rs.Fields("ParentGUID")
				
											
'------------------------------------------导出数据库信息至Excel-------------------------------------------

'-------------------------更新进度条add this code at the bottom of the loop to update the counter----------------------

		'percentComplete.Start "Exporting, please wait......"
		percentComplete.Progress 50



	Dim arr2(9999,1)  As Variant

	arr2(i,0) = ProjectInfo(i).ChildGUID
	arr2(i,1)  = ProjectInfo(i).parentGUID

		t= t+1			
				i = i + 1
				'move to the next record.
				rs.MoveNext
			Loop
	
	oBook.Worksheets("ProjectOverview_Parent_export").cells(2,1).resize(i,2).value = arr2
		
			Set rs = Nothing
	
	Set objConn = Nothing
	
	oBook.Worksheets("backup").range("A1").value = 0
	
	
	excel.application.DisplayAlerts = False
	oBook.Close (True)
	excel.quit

'-------------------------更新进度条add this code at the bottom of the loop to update the counter----------------------

		'percentComplete.Start "Exporting, please wait......"
		percentComplete.Progress 50

	
	Set excel = CreateObject("Excel.Application")

	
	
	'-------------------------更新进度条add this code at the bottom of the loop to update the counter----------------------

		'percentComplete.Start "Exporting, please wait......"
		percentComplete.Progress 140
		
	
	Set oBook = excel.Workbooks.Open (dstr)
	excel.application.DisplayAlerts = False
		excel.Visible = true
	oBook.Close (True)
	excel.quit
	
	Set oBook = Nothing
	Set excel = Nothing
	

	
	'MsgBox "程序执行时间为：" & Format(Timer - start, "0.00") & "秒"
	
	'MsgBox "History exported! " & " It took " & Format(Timer - start, "0.00") & " seconds. " & Chr(13) &"Please turn To:" & Chr(13) & dstr
	
End Sub



































