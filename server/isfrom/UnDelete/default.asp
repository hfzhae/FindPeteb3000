<!-- #include virtual="/NetPower2.asp" --><!-- #include file="../Common.asp" -->
<!-- #include virtual="/Server/public.asp" -->
<%
'if not CheckPrivilege(ModName & ".UnDelete") then err.raise 10000, ModName & ".UnDelete", "NoPrivilege"

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

Function UnDeleteEx()
	dim sSQL, afectedrecords
	UnDeleteEx =0
	sSQL = "Select * from [" & TableName & "] Where ID=" & ID
	set rs = CreateObject("adodb.recordset")
	rs.CursorLocation =3
	rs.Open sSQL, ConnEx, 0, 1, 1
	if rs.eof then err.raise 10000, ModName & ".UnDelete", "NotExistsID"
	if rs("IsDeleted").value =0 then err.raise 10000, ModName & ".UnDelete", "UnDeleted"
	
	sSQL = "Update [" & TableName & "] set UpdateDate = " & DBDateString & ", Owner=" & Owner & ", IsDeleted=0 where "
	if rsFieldExists(rs, "AuditID") then
		if ClngEx(rs("AuditID").value) > 0 then err.raise 10000, ModName & ".UnDelete", "Audited"
		if rsFieldExists(rs, "BillTitle") then strTitle =rs("BillTitle").value else strTitle =ID
		sSQL = sSQL & "IsDeleted = 1 and AuditID =0 and ID=" & ID
	else
		sSQL = sSQL & "IsDeleted = 1 and ID=" & ID
		if rsFieldExists(rs, "Title") then strTitle =rs("Title").value else strTitle =ID
	end if
	ConnEx.Execute sSQL, afectedrecords
	if afectedrecords =0 then err.raise 10000, ModName & ".UnDelete", "UnDeleted"
	fnUnDelete
	UnDeleteEx =ID
End Function

Function fnUnDelete()
End Function

dim ID, rs, strTitle
UnDelete
Sub UnDelete
	ID = ClngEx(stdin("ID"))
	ConnEx.BeginTrans
	on error resume next
	stdout("ID") =UnDeleteEx()
	if err.number <> 0 then
		ConnEx.RollbackTrans
		getLastError
		LogWrite EventName & ".»Ö¸´", "Ê§°Ü, " & TableName & " ID=" & ID & ", Title=" & strTitle & ", " & err.Description
		on error goto 0
		if stdout("Err.Source") = "" then
		    err.raise stdout("Err.Number"), ModName & ".UnDelete", stdout("Err.Description")
		else
		    err.raise stdout("Err.Number"), stdout("Err.Source"), stdout("Err.Description")
		end if
	else
		ConnEx.CommitTrans
		LogWrite EventName & ".»Ö¸´", "³É¹¦, " & TableName & ", ID=" & ID & ",  Title=" & strTitle
	end if
	set stdin = Nothing
	set ConnEx = Nothing
	set rs = Nothing
end Sub
%>
