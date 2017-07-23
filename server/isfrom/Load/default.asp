<!-- #include virtual="/NetPower2.asp" --><!-- #include file="../Common.asp" -->
<!-- #include virtual="/Server/cblib.asp" -->
<!-- #include virtual="/Server/public.asp" -->
<% 
dim ID, rs

Sub Load()
	If IsEmpty(ID) then ID = ClngEx(stdin("ID"))
	Set rs = connEx.execute("select * from [" & TableName & "] where id=" & ID)
	if rs.eof then Err.Raise 10000, ModName & ".Load", "NotExistsID"
	'if Clng(rs("IsDeleted").value) = 1 then Err.Raise 10000, ModName & ".Load", "Deleted"
	fnMapping
	np2.mapBind stdout
	fnLoad()
	LogWrite EventName & ".´ò¿ª", "ID=" & ID & ",  Title=" & rs("title").value
	set stdin = Nothing
End Sub

Function fnMapping()
end Function

Function fnLoad()
end Function

Load
Function fnMapping()
	map ModName, "AddNew","",MAPVOID,"cbRSAddNew"
	map ModName,"ID","ID",MAPINT,"cbRSDirect"
	map ModName,"title","title",MAPSTRING,"cbRSDirect"
	map ModName,"sort","sort",MAPINT,"cbRSDirect"
End Function

Function fnLoad()
	np2.mapping 0, stdout, ModName, rs
end Function

%>