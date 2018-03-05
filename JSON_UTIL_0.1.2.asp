<%
'I don't remember where comes from this code... sorry....'
'This Code using with json 2.0.4 --> https://code.google.com/archive/p/aspjson/'

'This is already existing code..'
Function QueryToJSON(Db, sql)
        Dim rs, jsa
		'Set rs = Db.Execute(sql)'
		Set rs = Db.execRs(sql, DB_CMDTYPE_TEXT, Null, Nothing)
        Set jsa = jsArray()
        While Not (rs.EOF Or rs.BOF)
                Set jsa(Null) = jsObject()
                For Each col In rs.Fields
                        jsa(Null)(col.Name) = col.Value
                Next
        rs.MoveNext
        Wend
        Set QueryToJSON = jsa
		rs.close : Set rs = Nothing
End Function

'Added by jaymz9634 at 2018-03-05 : upper code is'nt fit my dev env... so I Added this code'
Function RsToJSON(rs)
        Set jsa = jsArray()
        While Not (rs.EOF Or rs.BOF)
                Set jsa(Null) = jsObject()
                For Each col In rs.Fields
                        jsa(Null)(col.Name) = col.Value
                Next
        rs.MoveNext
        Wend
        Set RsToJSON = jsa
		rs.close : Set rs = Nothing
End Function
%>
