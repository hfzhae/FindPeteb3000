<!-- #include virtual="/NetPower2.asp" --><!-- #include file="../Common.asp" -->
<!-- #include virtual="/Server/cblib.asp" -->
<!-- #include virtual="/Server/public.asp" -->

<% @debug=on
dim ID, ParentID, rs

Sub Save()
	If IsEmpty(ID) then ID = ClngEx(stdin(ModName)("ID"))
	ParentID = ClngEx(stdin("ParentID"))
	'checkPrivilege_Save ModName, ID
	connEx.begintrans
	on error resume next
	stdout("ID") = SaveEx()
	if err.number <>0 then
		getLastError
		connEx.RollbackTrans
		LogWrite EventName & ".保存", "失败, ID=" & ID & ", Err.Description =" & Err.Description
		if stdout("Err.Source") = "" then
		    err.raise stdout("Err.Number"), ModName & ".Save", stdout("Err.Description")
		else
		    err.raise stdout("Err.Number"), stdout("Err.Source"), stdout("Err.Description")
		end if
	else
		connEx.commitTrans
		LogWrite EventName & ".保存", "成功, ID=" & ID & ", Title=" & rs("id").value
	end if
	
	set rs =Nothing
	set stdin = Nothing
End Sub

Function SaveEx()
	SaveEx = 0
	if ID =0 then 
		'Set rs = dbx(TableName).AddNew()
		'Set rs = connEx.execute("select * from [" & TableName & "] where 1=2").AddNew()
		Set rs=Server.CreateObject("ADODB.Recordset") 
        rs.open "select * from [" & TableName & "] where 1=2", connEx, 3, 3, 1
        
        rs.AddNew()
		'ID = CTIDGen(IGID)
		'rs("ID") = ID
		'rs("ParentID") = ParentID
		'rs("RootID") = getRootID(ParentID)
		rs("InfoType") = ModType
		rs("CreateDate") = Now
		rs("UpdateDate") = rs("CreateDate").value
		'rs("AccountID") = AccountID
		rs("UpdateCount") = 0
		rs("IsDeleted") = 0
	Else
		if not NetBox.TryLock(ModType & "." & ID) then err.raise 1000, ModName, "LockByOther"
		Set rs=Server.CreateObject("ADODB.Recordset") 
        rs.open "select * from [" & TableName & "] where id=" & ID,connEx,3,3,1
		if rs.eof then Err.Raise 10000, ModName & ".Save", "NotExistsID"
		ID = rs("ID").value
	    rs("UpdateDate") = Now
	End If
	rs("Owner") = Owner
	fnMapping
	fnSave
	rs.Update
	ID = connEx.execute("select max(id) as id from [" & TableName & "]")(0).value
	SaveEx = ID
End Function

Sub checkPrivilege_Save(ModName, id)
	dim strPrivilege
	id =ClngEx(id)
	if id = 0 then strPrivilege ="Create" else strPrivilege ="Modify"
	if not CheckPrivilege(ModName & "." & strPrivilege) then err.raise 10000, ModName & ".Save", "NoPrivilege"
End Sub

dim Int1
Int1 =0
Save

function fnMapping()
	map ModName,"version","version",MAPINT,"cbRSDirect"
end function

function fnSave()
	np2.mapping 1, stdin, ModName, rs, ID, TableName
end function

%>