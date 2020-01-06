<script language="jscript" runat="server">  
    function parseJSON(strJSON) { 
		//if(window.JSON) {
		//	var ob=JSON.parse(strJSON);
		//	return ob;
		//} else {
			return eval("(" + strJSON + ")");
		//}
	}
	function parseMac(jobj, num) {
		return eval("obj.COMTYPE.CONTENT.DATA.MAC" + num);
	}
	function parseJStr(jobj) {
		return JSON.stringify(jobj);
	}
</script>
<%
response.expires=-1
dim a(30)
'Fill up array with names
a(1)="Anna"
a(2)="Brittany"
a(3)="Cinderella"

'从 URL 获得参数 q
q=ucase(request.querystring("q"))

'如果长度 q>0，则从数组中查找所有提示
if len(q)>0 then
	hint=""
	for i=1 to 30
		if q=ucase(mid(a(i),1,len(q))) then
			if hint="" then
				hint=a(i)
			else
				hint=hint & " , " & a(i)
			end if
		end if
	next
end if

Set obj=parseJSON(q)
set fso=server.CreateObject("Scripting.FileSystemObject")
path=server.MapPath("./")
path=path&"/list.js"
set fd=fso.OpenTextFile(path,1,true,-2)
if obj.COMTYPE.CONTENT.TYPE="REPORTBORDCAST" then
	'接收的数量
	num=Cint(obj.COMTYPE.CONTENT.DATA.NUMBER)
	'读取文件内容
	bf=fd.readAll
	fd.close
	'文件内容生成json对象
	Set robj=parseJSON(bf)
	'文件储存的数量
	rnum=Cint(robj.COMTYPE.CONTENT.DATA.NUMBER)
	response.write(num)
	response.write(rnum)
	'遍历接收的数量
	for it=1 to num+1
		'获取mac
		gmac=parseMac(obj, it)
		response.write(gmac)
		for ji=1 to rnum+1
			if(parseMac(robj, ji)=gmac) then
				'robj.COMTYPE.CONTENT.DATA.PACK11="21"
			else
				response.write("566")
			end if
		next
	next
	'v=robj
	set ts=fso.OpenTextFile(path,2,true,-2)
	ts.writeline""&q&""
	ts.close
end if

'set ts=nothingset 
'fso=nothing

'如果未找到提示，则输出 "no suggestion"
'or output the correct values
if hint="" then
  response.write("no suggestion")
  response.write("no suggestion")
else
  response.write(hint)
end if

%>