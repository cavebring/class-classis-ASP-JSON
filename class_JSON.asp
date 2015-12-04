<%
'**************************************************************************************************************

'' @CLASSTITLE:		JSON
'' @CREATOR:		Michal Gabrukiewicz (gabru at grafix.at), Michael Rebec
'' @CONTRIBUTORS:	- Cliff Pruitt (opensource at crayoncowboy.com)
''					- Sylvain Lafontaine
''					- Jef Housein
''					- Jeremy Brown
'' @CREATEDON:		2007-04-26 12:46
'' @CDESCRIPTION:	Comes up with functionality for JSON (http://json.org) to use within ASP.
''					Correct escaping of characters, generating JSON Grammer out of ASP datatypes and structures
''					Some examples (all use the <em>toJSON()</em> method but as it is the class' default method it can be left out):
''					<code>
''					<%
''					'simple number
''					output = (new JSON)("myNum", 2, false)
''					'generates {"myNum": 2}
''										
''					'array with different datatypes
''					output = (new JSON)("anArray", array(2, "x", null), true)
''					'generates "anArray": [2, "x", null]
''					'(note: the last parameter was true, thus no surrounding brackets in the result)
''					% >
''					</code>
'' @REQUIRES:		-
'' @OPTIONEXPLICIT:	yes
'' @VERSION:		1.5.1

'**************************************************************************************************************
class rsJSON

	'private members
	private output, innerCall
	
	'public members
	public toResponse		''[bool] should the generated representation be written directly to the response (using <em>response.write</em>)? default = false
	public recordsetPaging	''[bool] indicates if only the current page should be processed on paged recordsets.
							''e.g. would return only 10 records if <em>RS.pagesize</em> is set to 10. default = false (means that always all records are processed)
	
	'**********************************************************************************************************
	'* constructor 
	'**********************************************************************************************************
	public sub class_initialize()
		newGeneration()
		toResponse = false
		recordsetPaging = false
	end sub
	
	'******************************************************************************************
	'' @SDESCRIPTION:	STATIC! takes a given string and makes it JSON valid
	'' @DESCRIPTION:	all characters which needs to be escaped are beeing replaced by their
	''					unicode representation according to the 
	''					RFC4627#2.5 - http://www.ietf.org/rfc/rfc4627.txt?number=4627
	'' @PARAM:			val [string]: value which should be escaped
	'' @RETURN:			[string] JSON valid string
	'******************************************************************************************
	public function escape(val)
		dim cDoubleQuote, cRevSolidus, cSolidus
		cDoubleQuote = &h22
		cRevSolidus = &h5C
		cSolidus = &h2F
		dim i, currentDigit
		for i = 1 to (len(val))
			currentDigit = mid(val, i, 1)
			if ascw(currentDigit) > &h00 and ascw(currentDigit) < &h1F then
				currentDigit = escapequence(currentDigit)
			elseif ascw(currentDigit) >= &hC280 and ascw(currentDigit) <= &hC2BF then
				currentDigit = "\u00" + right(padLeft(hex(ascw(currentDigit) - &hC200), 2, 0), 2)
			elseif ascw(currentDigit) >= &hC380 and ascw(currentDigit) <= &hC3BF then
				currentDigit = "\u00" + right(padLeft(hex(ascw(currentDigit) - &hC2C0), 2, 0), 2)
			else
				select case ascw(currentDigit)
					case cDoubleQuote: currentDigit = escapequence(currentDigit)
					case cRevSolidus: currentDigit = escapequence(currentDigit)
					case cSolidus: currentDigit = escapequence(currentDigit)
				end select
			end if
			escape = escape & currentDigit
		next
	end function
	
	'******************************************************************************************************************
	'' @SDESCRIPTION:	generates a representation of a name value pair in JSON grammer
	'' @DESCRIPTION:	It generates a name value pair which is represented as <em>{"name": value}</em> in JSON.
	''					the generation is fully recursive. Thus the value can also be a complex datatype (array in dictionary, etc.) e.g.
	''					<code>
	''					<%
	''					set j = new JSON
	''					j.toJSON "n", array(RS, dict, false), false
	''					j.toJSON "n", array(array(), 2, true), false
	''					% >
	''					</code>
	'' @PARAM:			name [string]: name of the value (accessible with javascript afterwards). leave empty to get just the value
	'' @PARAM:			val [variant], [int], [float], [array], [object], [dictionary], [recordset]: value which needs
	''					to be generated. Conversation of the data types is as follows:<br>
	''					- <strong>ASP datatype -> JavaScript datatype</strong>
	''					- NOTHING, NULL -> null
	''					- INT, DOUBLE -> number
	''					- STRING -> string
	''					- BOOLEAN -> bool
	''					- ARRAY -> array
	''					- DICTIONARY -> Represents it as name value pairs. Each key is accessible as property afterwards. json will look like <code>"name": {"key1": "some value", "key2": "other value"}</code>
	''					- <em>multidimensional array</em> -> Generates a 1-dimensional array (flat) with all values of the multidimensional array
	''					- RECORDSET -> array where each row of the recordset represents a field of the array. Each array field has properties according to the column names of the recordset (<strong>lowercase!</strong>) e.g <em>toJSON("r", RS, false)</em> can be accessed afterwards with <em>r[0].id</em>
	''					- <em>request</em> object -> every property and collection (cookies, form, querystring, etc) of the asp request object is exposed as an item of a dictionary. Property names are <strong>lowercase</strong>. e.g. <em>servervariables</em>.
	''					- OBJECT -> name of the type (if unknown type) or all its properties (if class implements <em>reflect()</em> method)
	''					Implement a <strong>reflect()</strong> function if you want your custom classes to be recognized. The function must return
	''					a dictionary where the key holds the property name and the value its value. Example of a reflect function within a User class which has firstname and lastname properties
	''					<code>
	''					<%
	''					function reflect()
	''					.	set reflect = server.createObject("scripting.dictionary")
	''					.	reflect.add "firstname", firstname
	''					.	reflect.add "lastname", lastname
	''					end function
	''					% >
	''					</code>
	''					Example of how to generate a JSON representation of the asp request object and access the <em>HTTP_HOST</em> server variable in JavaScript:
	''					<code>
	''					<script>alert(<%= (new JSON)(empty, request, false) % >.servervariables.HTTP_HOST);</script>
	''					</code>
	'' @PARAM:			nested [bool]: indicates if the name value pair is already nested within another? if yes then the <em>{}</em> are left out.
	'' @RETURN:			[string] returns a JSON representation of the given name value pair
	''					(if toResponse is on then the return is written directly to the response and nothing is returned)
	'******************************************************************************************************************
	public default function toJSON(name, val, nested)
		if not nested and not isEmpty(name) then write("{")
		if not isEmpty(name) then write("""" & escape(name) & """: ")
		generateValue(val)
		if not nested and not isEmpty(name) then write("}")
		toJSON = output
		
		if innerCall = 0 then newGeneration()
	end function
	
	'******************************************************************************************************************
	'* generate 
	'******************************************************************************************************************
	private function generateValue(val)
		if isNull(val) then
			write("null")
		elseif isArray(val) then
			generateArray(val)
		elseif isObject(val) then
			dim tName : tName = typename(val)
			if val is nothing then
				write("null")
			elseif tName = "Dictionary" or tName = "IRequestDictionary" then
				generateDictionary(val)
			elseif tName = "Recordset" then
				generateRecordset(val)
			elseif tName = "IRequest" then
				set req = server.createObject("scripting.dictionary")
				req.add "clientcertificate", val.ClientCertificate
				req.add "cookies", val.cookies
				req.add "form", val.form
				req.add "querystring", val.queryString
				req.add "servervariables", val.serverVariables
				req.add "totalbytes", val.totalBytes
				generateDictionary(req)
			elseif tName = "IStringList" then
				if val.count = 1 then
					toJSON empty, val(1), true
				else
					generateArray(val)
				end if
			else
				generateObject(val)
			end if
		else
			'bool
			dim varTyp
			varTyp = varType(val)
			if varTyp = 11 then
				if val then write("true") else write("false")
			'int, long, byte
			elseif varTyp = 2 or varTyp = 3 or varTyp = 17 or varTyp = 19 then
				write(cLng(val))
			'single, double, currency
			elseif varTyp = 4 or varTyp = 5 or varTyp = 6 or varTyp = 14 then
				write(replace(cDbl(val), ",", "."))
			else
				write("""" & escape(val & "") & """")
			end if
		end if
		generateValue = output
	end function
	
	'******************************************************************************************************************
	'* generateArray 
	'******************************************************************************************************************
	private sub generateArray(val)
		dim item, i
		write("[")
		i = 0
		'the for each allows us to support also multi dimensional arrays
		for each item in val
			if i > 0 then write(",")
			generateValue(item)
			i = i + 1
		next
		write("]")
	end sub
	
	'******************************************************************************************************************
	'* generateDictionary 
	'******************************************************************************************************************
	private sub generateDictionary(val)
		innerCall = innerCall + 1
		if val.count = 0 then
			toJSON empty, null, true
			exit sub
		end if
		dim key, i
		write("{")
		i = 0
		for each key in val
			if i > 0 then write(",")
			toJSON key, val(key), true
			i = i + 1
		next
		write("}")
		innerCall = innerCall - 1
	end sub
	
	'******************************************************************************************************************
	'* generateRecordset 
	'******************************************************************************************************************
	private sub generateRecordset(val)
		dim i, curRow
		write("[")
		curRow = 0
		'recordset.pagesize = -1 means it is not paged.
		while not val.eof and ((recordsetPaging and curRow < val.pageSize) or val.recordCount = -1 or not recordsetPaging)
			innerCall = innerCall + 1
			write("{")
			for i = 0 to val.fields.count - 1
				if i > 0 then write(",")
				toJSON lCase(val.fields(i).name), val.fields(i).value, true
			next
			write("}")
			val.movenext()
			curRow = curRow + 1
			if not val.eof and ((recordsetPaging and curRow < val.pageSize) or val.recordCount = -1 or not recordsetPaging) then write(",")
			innerCall = innerCall - 1
		wend
		write("]")
	end sub
	
	'******************************************************************************************************************
	'* generateObject 
	'******************************************************************************************************************
	private sub generateObject(val)
		dim props
		on error resume next
		set props = val.reflect()
		if err = 0 then
			on error goto 0
			innerCall = innerCall + 1
			toJSON empty, props, true
			innerCall = innerCall - 1
		else
			on error goto 0
			write("""" & escape(typename(val)) & """")
		end if
	end sub
	
	'******************************************************************************************************************
	'* newGeneration 
	'******************************************************************************************************************
	private sub newGeneration()
		output = empty
		innerCall = 0
	end sub
	
	'******************************************************************************************
	'* JsonEscapeSquence 
	'******************************************************************************************
	private function escapequence(digit)
		escapequence = "\u00" + right(padLeft(hex(ascw(digit)), 2, 0), 2)
	end function
	
	'******************************************************************************************
	'* padLeft 
	'******************************************************************************************
	private function padLeft(value, totalLength, paddingChar)
		padLeft = right(clone(paddingChar, totalLength) & value, totalLength)
	end function
	
	'******************************************************************************************
	'* clone 
	'******************************************************************************************
	private function clone(byVal str, n)
		dim i
		for i = 1 to n : clone = clone & str : next
	end function
	
	'******************************************************************************************
	'* write 
	'******************************************************************************************
	private sub write(val)
		if toResponse then
			response.write(val)
		else
			output = output & val
		end if
	end sub

end class


'Februari 2014 - Version 1.17 by Gerrit van Kuipers
Class JSON
	Public data
	Private p_JSONstring
	private aj_in_string, aj_in_escape, aj_i_tmp, aj_char_tmp, aj_s_tmp, aj_line_tmp, aj_line, aj_lines, aj_currentlevel, aj_currentkey, aj_currentvalue, aj_newlabel, aj_XmlHttp, aj_RegExp, aj_colonfound

	Private Sub Class_Initialize()
		Set data = Collection()

	    Set aj_RegExp = new regexp
	    aj_RegExp.Pattern = "\s{0,}(\S{1}[\s,\S]*\S{1})\s{0,}"
	    aj_RegExp.Global = False
	    aj_RegExp.IgnoreCase = True
	    aj_RegExp.Multiline = True
	End Sub

	Private Sub Class_Terminate()
		Set data = Nothing
	    Set aj_RegExp = Nothing
	End Sub

	Public Sub loadJSON(inputsource)
		inputsource = aj_MultilineTrim(inputsource)
		If Len(inputsource) = 0 Then Err.Raise 1, "loadJSON Error", "No data to load."
		
		select case Left(inputsource, 1)
			case "{", "["
			case else
				Set aj_XmlHttp = Server.CreateObject("Msxml2.ServerXMLHTTP")
				aj_XmlHttp.open "GET", inputsource, False
				aj_XmlHttp.setRequestHeader "Content-Type", "text/json"
				aj_XmlHttp.setRequestHeader "CharSet", "UTF-8"
				aj_XmlHttp.Send
				inputsource = aj_XmlHttp.responseText
				set aj_XmlHttp = Nothing
		end select

		p_JSONstring = CleanUpJSONstring(inputsource)
		aj_lines = Split(p_JSONstring, Chr(13) & Chr(10))

		Dim level(99)
		aj_currentlevel = 1
		Set level(aj_currentlevel) = data
		For Each aj_line In aj_lines
			aj_currentkey = ""
			aj_currentvalue = ""
			If Instr(aj_line, ":") > 0 Then
				aj_in_string = False
				aj_in_escape = False
				aj_colonfound = False
				For aj_i_tmp = 1 To Len(aj_line)
					If aj_in_escape Then
						aj_in_escape = False
					Else
						Select Case Mid(aj_line, aj_i_tmp, 1)
							Case """"
								aj_in_string = Not aj_in_string
							Case ":"
								If Not aj_in_escape And Not aj_in_string Then
									aj_currentkey = Left(aj_line, aj_i_tmp - 1)
									aj_currentvalue = Mid(aj_line, aj_i_tmp + 1)
									aj_colonfound = True
									Exit For
								End If
							Case "\"
								aj_in_escape = True
						End Select
					End If
				Next
				if aj_colonfound then
					aj_currentkey = aj_Strip(aj_JSONDecode(aj_currentkey), """")
					If Not level(aj_currentlevel).exists(aj_currentkey) Then level(aj_currentlevel).Add aj_currentkey, ""
				end if
			End If
			If right(aj_line,1) = "{" Or right(aj_line,1) = "[" Then
				If Len(aj_currentkey) = 0 Then aj_currentkey = level(aj_currentlevel).Count
				Set level(aj_currentlevel).Item(aj_currentkey) = Collection()
				Set level(aj_currentlevel + 1) = level(aj_currentlevel).Item(aj_currentkey)
				aj_currentlevel = aj_currentlevel + 1
				aj_currentkey = ""
			ElseIf right(aj_line,1) = "}" Or right(aj_line,1) = "]" or right(aj_line,2) = "}," Or right(aj_line,2) = "]," Then
				aj_currentlevel = aj_currentlevel - 1
			ElseIf Len(Trim(aj_line)) > 0 Then
				if Len(aj_currentvalue) = 0 Then aj_currentvalue = aj_line
				aj_currentvalue = getJSONValue(aj_currentvalue)

				If Len(aj_currentkey) = 0 Then aj_currentkey = level(aj_currentlevel).Count
				level(aj_currentlevel).Item(aj_currentkey) = aj_currentvalue
			End If
		Next
	End Sub

	Public Function Collection()
		set Collection = Server.CreateObject("Scripting.Dictionary")
	End Function

	Public Function AddToCollection(dictobj)
		if TypeName(dictobj) <> "Dictionary" then Err.Raise 1, "AddToCollection Error", "Not a collection."
		aj_newlabel = dictobj.Count
		dictobj.Add aj_newlabel, Collection()
		set AddToCollection = dictobj.item(aj_newlabel)
	end function

	Private Function CleanUpJSONstring(aj_originalstring)
		aj_originalstring = Replace(aj_originalstring, Chr(13) & Chr(10), "")
		aj_originalstring = Mid(aj_originalstring, 2, Len(aj_originalstring) - 2)
		aj_in_string = False : aj_in_escape = False : aj_s_tmp = ""
		For aj_i_tmp = 1 To Len(aj_originalstring)
			aj_char_tmp = Mid(aj_originalstring, aj_i_tmp, 1)
			If aj_in_escape Then
				aj_in_escape = False
				aj_s_tmp = aj_s_tmp & aj_char_tmp
			Else
				Select Case aj_char_tmp
					Case "\" : aj_s_tmp = aj_s_tmp & aj_char_tmp : aj_in_escape = True
					Case """" : aj_s_tmp = aj_s_tmp & aj_char_tmp : aj_in_string = Not aj_in_string
					Case "{", "["
						aj_s_tmp = aj_s_tmp & aj_char_tmp & aj_InlineIf(aj_in_string, "", Chr(13) & Chr(10))
					Case "}", "]"
						aj_s_tmp = aj_s_tmp & aj_InlineIf(aj_in_string, "", Chr(13) & Chr(10)) & aj_char_tmp
					Case "," : aj_s_tmp = aj_s_tmp & aj_char_tmp & aj_InlineIf(aj_in_string, "", Chr(13) & Chr(10))
					Case Else : aj_s_tmp = aj_s_tmp & aj_char_tmp
				End Select
			End If
		Next

		CleanUpJSONstring = ""
		aj_s_tmp = split(aj_s_tmp, Chr(13) & Chr(10))
		For Each aj_line_tmp In aj_s_tmp
			aj_line_tmp = replace(replace(aj_line_tmp, chr(10), ""), chr(13), "")
			CleanUpJSONstring = CleanUpJSONstring & aj_Trim(aj_line_tmp) & Chr(13) & Chr(10)
		Next
	End Function

	Private Function getJSONValue(ByVal val)
		val = Trim(val)
		If Left(val,1) = ":"  Then val = Mid(val, 2)
		If Right(val,1) = "," Then val = Left(val, Len(val) - 1)
		val = Trim(val)

		Select Case val
			Case "true"  : getJSONValue = True
			Case "false" : getJSONValue = False
			Case "null" : getJSONValue = Null
			Case Else
				If (Instr(val, """") = 0) Then
					If IsNumeric(val) Then
						getJSONValue = CDbl(val)
					Else
						getJSONValue = val
					End If
				Else
					If Left(val,1) = """" Then val = Mid(val, 2)
					If Right(val,1) = """" Then val = Left(val, Len(val) - 1)
					getJSONValue = aj_JSONDecode(Trim(val))
				End If
		End Select
	End Function

	Private JSONoutput_level
	Public Function JSONoutput()
		dim wrap_dicttype, aj_label
		JSONoutput_level = 1
		wrap_dicttype = "[]"
		For Each aj_label In data
			 If Not aj_IsInt(aj_label) Then wrap_dicttype = "{}"
		Next
		JSONoutput = Left(wrap_dicttype, 1) & Chr(13) & Chr(10) & GetDict(data) & Right(wrap_dicttype, 1)
	End Function

	Private Function GetDict(objDict)
		dim aj_item, aj_keyvals, aj_label, aj_dicttype
		For Each aj_item In objDict
			Select Case TypeName(objDict.Item(aj_item))
				Case "Dictionary"
					GetDict = GetDict & Space(JSONoutput_level * 4)
					
					aj_dicttype = "[]"
					For Each aj_label In objDict.Item(aj_item).Keys
						 If Not aj_IsInt(aj_label) Then aj_dicttype = "{}"
					Next
					If aj_IsInt(aj_item) Then
						GetDict = GetDict & (Left(aj_dicttype,1) & Chr(13) & Chr(10))
					Else
						GetDict = GetDict & ("""" & aj_JSONEncode(aj_item) & """" & ": " & Left(aj_dicttype,1) & Chr(13) & Chr(10))
					End If
					JSONoutput_level = JSONoutput_level + 1
					
					aj_keyvals = objDict.Keys
					GetDict = GetDict & (GetSubDict(objDict.Item(aj_item)) & Space(JSONoutput_level * 4) & Right(aj_dicttype,1) & aj_InlineIf(aj_item = aj_keyvals(objDict.Count - 1),"" , ",") & Chr(13) & Chr(10))
				Case Else
					aj_keyvals =  objDict.Keys
					GetDict = GetDict & (Space(JSONoutput_level * 4) & aj_InlineIf(aj_IsInt(aj_item), "", """" & aj_JSONEncode(aj_item) & """: ") & WriteValue(objDict.Item(aj_item)) & aj_InlineIf(aj_item = aj_keyvals(objDict.Count - 1),"" , ",") & Chr(13) & Chr(10))
			End Select
		Next
	End Function

	Private Function aj_IsInt(val)
		aj_IsInt = (TypeName(val) = "Integer" Or TypeName(val) = "Long")
	End Function

	Private Function GetSubDict(objSubDict)
		GetSubDict = GetDict(objSubDict)
		JSONoutput_level= JSONoutput_level -1
	End Function

	Private Function WriteValue(ByVal val)
		Select Case TypeName(val)
			Case "Double", "Integer", "Long": WriteValue = val
			Case "Null"						: WriteValue = "null"
			Case "Boolean"					: WriteValue = aj_InlineIf(val, "true", "false")
			Case Else						: WriteValue = """" & aj_JSONEncode(val) & """"
		End Select
	End Function

	Private Function aj_JSONEncode(ByVal val)
		val = Replace(val, "\", "\\")
		val = Replace(val, """", "\""")
		'val = Replace(val, "/", "\/")
		val = Replace(val, Chr(8), "\b")
		val = Replace(val, Chr(12), "\f")
		val = Replace(val, Chr(10), "\n")
		val = Replace(val, Chr(13), "\r")
		val = Replace(val, Chr(9), "\t")
		aj_JSONEncode = Trim(val)
	End Function

	Private Function aj_JSONDecode(ByVal val)
		val = Replace(val, "\""", """")
		val = Replace(val, "\\", "\")
		val = Replace(val, "\/", "/")
		val = Replace(val, "\b", Chr(8))
		val = Replace(val, "\f", Chr(12))
		val = Replace(val, "\n", Chr(10))
		val = Replace(val, "\r", Chr(13))
		val = Replace(val, "\t", Chr(9))
		aj_JSONDecode = Trim(val)
	End Function

	Private Function aj_InlineIf(condition, returntrue, returnfalse)
		If condition Then aj_InlineIf = returntrue Else aj_InlineIf = returnfalse
	End Function

	Private Function aj_Strip(ByVal val, stripper)
		If Left(val, 1) = stripper Then val = Mid(val, 2)
		If Right(val, 1) = stripper Then val = Left(val, Len(val) - 1)
		aj_Strip = val
	End Function

	Private Function aj_MultilineTrim(TextData)
		aj_MultilineTrim = aj_RegExp.Replace(TextData, "$1")
	End Function

	private function aj_Trim(val)
		aj_Trim = Trim(val)
		Do While Left(aj_Trim, 1) = Chr(9) : aj_Trim = Mid(aj_Trim, 2) : Loop
		Do While Right(aj_Trim, 1) = Chr(9) : aj_Trim = Left(aj_Trim, Len(aj_Trim) - 1) : Loop
		aj_Trim = Trim(aj_Trim)
	end function
End Class


%>