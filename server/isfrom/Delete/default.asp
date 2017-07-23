<!-- #include virtual="/NetPower2.asp" --><!-- #include file="../Common.asp" -->
<!-- #include virtual="/Server/public.asp" -->
<%
'if not CheckPrivilege(ModName & ".Delete") then err.raise 10000, ModName & ".Delete", "NoPrivilege"

Function rsFieldExists(rs, sFiled)
	dim i, fCount, iField
	rsFieldExists = false
	
	iField = LCase(sFiled)
	fCount = rs.fields.count - 1
	for i = 0 to fCount
		if lcase(rs.fields(i).name) = iField then
			rsFieldExists = True
			Exit function
		end if
	next	
End Function

Function DeleteEx()
	dim sSQL, afectedrecords
	DeleteEx =0
	sSQL = "Select * from [" & TableName & "] Where ID=" & ID
	set rs = CreateObject("adodb.recordset")
	rs.CursorLocation =3
	rs.Open sSQL, ConnEx, 0, 1, 1
	if rs.eof then err.raise 10000, ModName & ".Delete", "NotExistsID"
	
	sSQL = "Update [" & TableName & "] set UpdateDate = " & DBDateString & ", Owner=" & Owner & ", IsDeleted=1 where "
	if rsFieldExists(rs, "AuditID") then
		if ClngEx(rs("AuditID").value) > 0 then err.raise 10000, ModName & ".Delete", "Audited"
		if rsFieldExists(rs, "BillTitle") then strTitle =rs("BillTitle").value else strTitle =ID
		sSQL = sSQL & "IsDeleted = 0 and AuditID =0 and ID=" & ID
		ConnEx.Execute "Delete from UnWorkflowTask where AccountID=" & AccountID & " and BillType =" & rs("BillType").value & " and BillID=" & ID, afectedrecords
	else
		sSQL = sSQL & "IsDeleted = 0 and ID=" & ID
		if rsFieldExists(rs, "Title") then strTitle =rs("Title").value else strTitle =ID
	end if
	ConnEx.Execute sSQL, afectedrecords
	if afectedrecords =0 then err.raise 10000, ModName & ".Delete", "Deleted"
	fnDelete
	DeleteEx =ID
End Function

Function fnDelete()
End Function

dim rs, ID, strTitle
Delete

Sub Delete()
	ID = ClngEx(stdin("ID"))
	ConnEx.BeginTrans
	on error resume next
	stdout("ID") = DeleteEx()
	if err.number <> 0 then
		ConnEx.RollbackTrans
		getLastError
		LogWrite EventName & ".É¾³ý", "Ê§°Ü, " & TableName & " ID=" & ID & ", Title=" & strTitle & ", " & err.Description
		on error goto 0
		if stdout("Err.Source") = "" then
		    err.raise stdout("Err.Number"), ModName & ".Delete", stdout("Err.Description")
		else
		    err.raise stdout("Err.Number"), stdout("Err.Source"), stdout("Err.Description")
		end if
	else
		ConnEx.CommitTrans
		LogWrite EventName & ".É¾³ý", "³É¹¦, " & TableName & ", ID=" & ID & ",  Title=" & strTitle
	end if
	set stdin = Nothing
	set ConnEx = Nothing
	set rs = Nothing
end Sub
%>

