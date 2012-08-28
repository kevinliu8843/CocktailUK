<!--#include virtual="/includes/variables.asp" -->
<!--#include virtual="/includes/functions.asp" -->
<!--#include virtual="/includes/admin_functions.asp" -->
<!--#include virtual="/includes/clsDbBackup.asp" -->
<P>Starting database backup...</P>
<%
Server.ScriptTimeout = 600
Dim objDbBackup, timStart
timStart = Now()
Set objDbBackup = New clsDbBackup
Call objDbBackup.SetBackupPath("C:\Inetpub\wwwroot\backup\DB\cocktailuk")
objDbBackup.m_blnCreateStructure = False
objDbBackup.BackupDB()
%>
<P>Backup created, attemting to compress backup...</P>
<%
If objDbBackup.ZipBackup() Then
	objDbBackup.DeleteOrigBackup()
End If
Set objDbBackup = nothing
%>
<p>Database backed up in <%=DateDiff("s",timStart,Now())%> seconds</p>