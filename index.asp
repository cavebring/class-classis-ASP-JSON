<!--#include file="class_JSON.asp"-->
<%
'set instance'
set rs_json = new rsJSON
set j = new JSON

'query database
SQL = "Select top 100 User_ID from Users"
Set RecSet = Connect.Execute(SQL)

'load recset to jsonplaintext
jsonplaintext = rs_json.toJSON ("data", RecSet, false)

'Load JSON string
j.loadJSON(jsonplaintext)

if len(value) < 13 then '13 is the base data package, if less than 13 = NO DATA'
	response.write "no data"
else
	'return 1 post
	Set this = j.data("data").item(0)
	response.write this.item("user_id")
	'return several posts
	For Each recset In j.data("data")
	    Set this = j.data("data").item(recset)
	    Response.Write this.item("user_id") & "<br>"
	Next
end if

Set rs_json = Nothing
Set j = Nothing
%>
