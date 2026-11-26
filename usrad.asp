<%



		Dim con,rst,cmd
		Set con= server.CreateObject("ADODB.Connection")
		con.provider = "ADSDSOObject"
				con.Properties("User ID") = "administrador"
				con.Properties("Password") = ""
		con.Properties("Encrypt Password") = true

		 con.Open
		Set cmd = server.CreateObject("ADODB.Command")
		Set cmd.ActiveConnection= con

		strFilter = "(&(objectCategory=person)(objectClass=user)" & "(samAccountName=adm))"
		cmd.CommandText = "<LDAP://192.168.1.10>;" &strFilter &";memberof;subTree"

		Set rst = cmd.Execute()
		Response.Write "<TABLE BORDER=1>"
		Do While NOT rst.EOF

		   For I = 0 to rst.Fields.Count - 1
		        Response.Write "<TR><TD>" & rst.Fields(I).name  & "</TD><TD>"
		        If IsArray(rst.Fields(I).value) Then
		             strTemp = Join(rst.Fields(I).value )
		             Response.Write strTemp
		        Else
		            Response.Write rst.Fields(I).value
		        End If
		        Response.Write "</TD></TR>"
		   Next
		   rst.MoveNext
		Loop
Response.Write "</TABLE>"



%>