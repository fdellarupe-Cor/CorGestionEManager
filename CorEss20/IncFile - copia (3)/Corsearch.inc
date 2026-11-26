<%
Class Corsearch
    ' Propiedades

    public sql
    public cnb
    public ofv
    public orden_doc
    public canttx
	public cantreg
	public listacar
	public ADDINNER
	public search_phrase


' metodos

    Private Sub Class_Initialize()

    End Sub

    Public Function IsNoiseWordex(dato)
	       IsNoiseWordex =false
    End Function

    Public Function findDocex()
       set  findDocex = ofv.crearconsultaEx(SQL,cnbsq,1,parametros)
    End Function

	Function GetResultsFullTextex(search_phrase)
     Dim count, noise_words, query, search_string, search_strings, success, queryor, queryf
     Dim ppal
     search_strings = Split(search_phrase, " ", -1, 1)



     count = 0
     queryz = ""
     queryzor = ""
     ppal = ""
     For Each search_string In search_strings

        If Not IsNoiseWord(search_string) Then

            If CarEspeciales(search_string, listacar) Then
             wildcar = ""
            Else
             wildcar = "*"
            End If
            If count <> 0 Then

                queryz = queryz & " near " & Chr(34) & search_string & wildcar & Chr(34)
               ' queryzor = queryzor & " near " & Chr(34) & search_string & Chr(34)

            Else
             If ppal = "" Then ppal = search_string
                queryz = queryz & Chr(34) & search_string & wildcar & Chr(34)
               ' queryzor = queryzor & Chr(34) & search_string & Chr(34)
            End If

            count = count + 1



        End If

    Next

 GetResultsFullTextex = queryz

end function
End Class


%>