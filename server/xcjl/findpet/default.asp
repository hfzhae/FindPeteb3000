<!-- #include virtual="/NetPower2.asp" --><!-- #include file="../Common.asp" -->
<!-- #include virtual="/Server/cblib.asp" -->
<!-- #include virtual="/Server/public.asp" -->
<!-- #include virtual="/findpet.asp" -->

<% 
dim fp
set fp = new FindpetObj
fp.init(Conn)
set stdout("findpetRs") = fp.findpetRs(stdin("timeInput"), stdin("Varieties"), "", "", stdin("placeText"), "", stdin("sendType"))
%>