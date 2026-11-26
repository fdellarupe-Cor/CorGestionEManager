
<%

 FUNCTION GENFMTARCHIVOS(FNCODOBJ,FNTIPOBJ)


  FNSHTML = ""


  strsqlx = "Select Version,estado,tipobjeto From Objprty Where Objprty.codobjeto = " & nCodigo
  set rs = ofv.crearconsultaEx(StrSqlx,cn,1,parametros)
  bVersion = "NO"
  if not rs.eof then
    if rs(0) = "V" then bVersion = "SI"
      Nestado=rs(1)

  end if
  call ofv.cerrarconsulta(rs)


    sarch = "NO"

    Strsql = "SELECT descri, archivo, objid, directorio "
    Strsql = Strsql & " from ObjHtm Where codobjeto = " & nCodigo & " AND TIPO ='" & SKIND &"' 	"
    Strsql = Strsql & " order by secuen "

  Formulario = ""
  'response.write strsql
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then
    do until rs.eof

    sdirec = rs(3) & ""

    if sdirec = "" then sdirec = "Archivos"

       sarchX = sdirec & "/" & rs(1)

    if sdocum = cstr(rs(2)) then
       sarch = sdirec & "/" & rs(1)
       clase = "<font color=#993333><u>"
       stexto = sfida & sfta & replace(rs(0)," ","&nbsp;") & sftc
    else
       clase = ""
       stexto = sfidc & sft & replace(rs(0)," ","&nbsp;")
    end if

        xinstr = InStrRev(sarchX,".")
        if xinstr > 0 then
           xext = mid(sarchX,xinstr + 1,len(sarchX) - xinstr)
        end if
     if session("EMPMUN") = "S" then
      sarchX="Emp" & SESSION("D" & FNCODOBJ & "-0" )   & "/" & sarchX
    ' else
    ' sarchX="Emp" & SESSION("D" & FNCODOBJ & "-0" )   & "/" & sarchX
     end if
'response.write "Emp" & SESSION("D" & FNCODOBJ & "-0" )

      select case  Ucase(xext)
             case "TXT","MP4","WAV","MHT","JPG","JPEG","HTM","HTML","BMP","GIF","TIFF","XLS","ASP","PDF","PPT","DOC","MPG","MPEG","XLS","DOC","DOCX","XLSX"

                  sPath = Request.ServerVariables("PATH_INFO")

                  if  SESSION("URLPORT")  = "N" then
                      Xruta = Request.ServerVariables("SERVER_NAME") & left(sPath,instrRev(sPath,"/",-1))
                  else
                      Xruta = Request.ServerVariables("SERVER_NAME")& ":" & Request.ServerVariables("SERVER_PORT") & left(sPath,instrRev(sPath,"/",-1))
                  end if

                  Xruta = Xruta &  sarch

              '''   sAccion = "http://" & Xruta
                 sAccion =  sarch

                 stexto1 = "<A class=docs  href=" & chr(34) & sarchX & chr(34) & " target=""_blank"" "


             case else
                 stexto1 = "<A class=ttt  href=CoArchCons.asp?Codobjeto=" & ncodigo & "&Tipobjeto=" & stipo
                 stexto1 = stexto1 & "&sdocum=" & rs(2)

      end select



    stextoF=restaltarbsq(nCodigo,rs(0),stexto)

    if stextoF <> stexto then
    stitle=" title='Aqui se encuentra la palabra/s [" & session("yider") &"]'"
    else
    stitle=""
    end if

    stexto1 = stexto1  &  stitle & " >" & clase & stextof & "</u></font></a>"

    if (rs(1) & "") = "" then
       stexto1 = "<A class=ttt  href='javascript:void(0);' "
       stexto1 = stexto1 & "   >" & clase & stexto & "</u></font></a>"
    end if

       Celda = ""


      ' estado observacion

      Celdaobs  = ofv.GenCelda("","ut","","1%","","","","&nbsp;","si")




      if Nestado=5  then


               sAccion = "genObsdoc.asp?wasotxt=N&vuelve=" & vuelta  & "&Codobjeto=" & nCodigo
      	       sAccion = sAccion & "&Tipobjeto=" & sTipo & "&ID=" & rs(2)  & "&ID2=" & rs(2)

      	       xsffun  = " onclick=""document.getElementById('DX4').style.display='block';"" "
      	       sHref   = sAccion
      	       sInput  = "<div style='border-style:solid;border-color:#993333;border-width:1;' >" & "OBS" & "</div>"
      	    '   Celdaobs  = ofv.GenCelda("","ut" & xsffun & "  title='Agregar/consultar observaciones para el elemento.' ","","","1","","",sInput,"si")'
            '   Celdaobs  = ofv.GenLink("","ut" & "  ",sHref,"WK4","","",Celdaobs,"si")'

                      Celdaobs= ofv.GenLink("","ut" & "  ",sHref,"WK4","","",sInput,"si")
     Celdaobs =  ofv.GenCelda("","ut" & xsffun & "  title='Agregar/consultar observaciones para el elemento.' ","","","1","","",Celdaobs,"si")



      end if


      Celda = Celda & Celdaobs


       Celda = Celda & ofv.GenCelda("","ut  ","","95%","","","",stexto1,"si")
       Formulario = Celda
       Fila = fila & ofv.GenRow("","","","","","","",Formulario,"si")' & GetResultsFullText2(session("yider"), nCodigo)

    rs.movenext
    loop

  FNSHTML = FNSHTML  & ofv.GenTabla("","","","100%","","0","0","4","0","","",Fila,"si")
 '' if not sdocum = "" then
  if sarch <> "NO" then
     FNSHTML = FNSHTML & ofv.GenForm("cons","",sAccion,"_blank","post","",sInput,"si")
  end if

  else

    sdocus = "NO"


  end if

'''SELECT  codcomp, nombre, tipocomp FROM CSACOMPLE where codcomp

    fila = ""

    sarch = 0
  '  response.write ssclick
    if (skind=1 and session("APLMDCLI") = "BLP") OR SKIND=2 then
    'strsql = "SELECT DISTINCT COMPLEMENTOS.codcomp, COMPLEMENTOS.codiden, "
    'strsql = strsql & "COMPLEMENTOS.descri, CSACOMPLE.nombre FROM COMPLEMENTOS, "
    'strsql = strsql & " OBJACCION, CSACOMPLE where COMPLEMENTOS.codcomp = OBJACCION.codcomp "
   ' strsql = strsql & " and COMPLEMENTOS.codcomp = CSACOMPLE.codcomp "
    'strsql = strsql & " and OBJACCION.codobjeto =" & ncodigo
    'strsql = strsql & " ORDER BY COMPLEMENTOS.codcomp "

   'revisar!!!!!!
      Strsql = "Select distinct a.codobjetol, b.codiden, b.nombre , b.nombre From objlnk a, csacomple b  Where a.ltipo = 'KL' and a.codobjeto =" & ncodigo
	  strsql = strsql & " and a.codobjetol = b.codcomp "
      strsql = strsql & " ORDER BY a.codobjetol"
     'response.write strsql
    Set rs = OFV.crearconsultaEx(strsql, cn, 1, "")
  Formulario = ""
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then
    do until rs.eof


   if sdocum = cstr(rs(0)) then
       sarch = rs(0)
       clase = "<font color=#993333><u>"
       stexto = sfida & sfta & replace(rs(3) & ":&nbsp;" & rs(1) & "&nbsp;-&nbsp;" & rs(2)," ","&nbsp;") & sftc
    else
       clase = ""
       stexto = sfidc & sft & replace(rs(3) & ":&nbsp;" & rs(1) & "&nbsp;-&nbsp;" & rs(2)," ","&nbsp;")
    end if
    stexto1 = "<A class=ttt  href=CoArchCons.asp?Codobjeto=" & ncodigo & "&Tipobjeto=" & stipo
    stexto1 = stexto1 & "&sdocum=" & rs(0)
    stexto1 = stexto1 & "   >" & clase & stexto & "</u></font></a>"


       Celda = ofv.GenCelda("","ut  ","","1%","","","",stexto1,"si")
       Formulario = Celda

       rs.movenext

       Fila = fila & ofv.GenRow("","","","","","","",Formulario,"si")
    loop

    FNSHTML = FNSHTML  & ofv.GenTabla("","","","100%","","0","0","4","0","","",Fila,"si")
    END IF
    if  sarch > 0 then

        sAccion = "EditarImagen2.asp?ID=" & sarch
        FNSHTML   = FNSHTML & ofv.GenForm("cons","",sAccion,"_blank","post","",sInput,"si")
    end if

  else
    scomple = "NO"
  end if
    Call OFV.cerrarconsulta(rs)


  if sdocum <> "" then
     FNSHTML   = FNSHTML & "<script language=javascript >"
    FNSHTML   = FNSHTML & " document.cons.submit();"
     FNSHTML   = FNSHTML & "</script>"
  end if




  if scomple = "NO" and sdocus = "NO" then
    sestr = "<font>Sin&nbsp;datos&nbsp;disponibles"
    sestr = sestr & "</font>"
    Formulario = ofv.GenCelda("","ttt","center","80%","","","",sestr,"si")
    Fila = ofv.GenRow("","","","","","","",Formulario,"si")
    FNSHTML = FNSHTML  & ofv.GenTabla("","","","100%","","0","0","4","0","","",Fila,"si")
  end if


  GENFMTARCHIVOS = "<div id=zglos class='norma-card'>" & FNSHTML  & "</div>" '&  session("yider")





 END FUNCTION

function restaltarbsq(id,texto,stexto)

If SESSION("APLMDCLI") = "BV" Or SESSION("APLMDCLI") = "BGAL"  Or SESSION("APLMDCLI") = "BNX" Or SESSION("APLMDCLI") = "CORF" Or SESSION("APLMDCLI") = "COR"  Or SESSION("APLMDCLI") = "BC"  Or SESSION("APLMDCLI") = "BP" Or SESSION("APLMDCLI") = "BPN" Then

if session("yider") <> "" and session("usuario")<> "GOOGLE"  then
strsql= GetResultsFullText2(session("yider"),id,texto)
'response.write strsql & "<br>"

 Set rsbx = OFV.crearconsultaEx(strsql, cnbsq, 1, "")

 if not rsbx.eof then
 restaltarbsq= "<font color=red><u>" & texto & "</u></font>"
 else
 restaltarbsq=stexto
 end if

 Call OFV.cerrarconsulta(rsbx)
else

restaltarbsq=stexto

end if
else
restaltarbsq=stexto
end if
end function

Function GetResultsFullText2(search_phrase, id,texto)
    Dim count, noise_words, query, search_string, search_strings, success
    texto=replace(texto&"","&nbsp;"," ")
    search_strings = Split(search_phrase, " ", -1, 1)

    'query = "select [title], [text], [url],codigo,tipo,solapa,index1 from containstable(Yider, *, '"
    query = "select title, text, url,codigo,tipo,solapa,index1 from containstable(Yider, text, '"
    '  response.write query
    '"brown dog" and "lazy"') as ct join Yider as y on ct.[key]=y.[key] order by Rank desc

    count = 0
    For Each search_string In search_strings

        If Not IsNoiseWord(search_string) Then

            If count <> 0 Then
                query = query & " and """ & search_string & "*" & """"
            Else
                query = query & """" & search_string & "*" & """"
            End If

            count = count + 1


            AddToArray search_string, search_array
        End If

    Next

    query = query & "') as ct join Yider as y on ct.[key]=y.[OBJID] and y.codigo=" & id & " and y.url='indexdoc' and y.title like'%" & texto & "'   order by Rank desc,codigo,solapa"
    'Session("vvv") = Session("vvv") & query & "<br>"

   GetResultsFullText2= query
End Function

Function IsNoiseWord(noise_word)
    Dim is_noise_word, query

    On Error Resume Next
    Err.Clear

    query = "select top 1 [y].[OBJID] from containstable(Yider, *, '""" & noise_word & """') as ct join Yider as y on ct.[key]=y.[OBJID] order by Rank desc"

    '
    'response.write "<br>query is " & query
    m_cnb.Execute (query)
    '  response.write query & "----" & Err.Number & "<br>"
    ' MsgBox Err.Number
    If Err.Number = -2147217900 Then
        Err.Clear
        is_noise_word = True
    Else
        is_noise_word = False
    End If

    IsNoiseWord = is_noise_word

    If (Len(noise_word) < 3) Then IsNoiseWord = True
    'Call OFV.cerrarconsulta(Recordsetz)
End Function

Sub AddToArray(Item, ByRef the_array)

    Dim array_size

    If Not IsArray(the_array) Then
        the_array = Array()
    End If

    array_size = UBound(the_array) + 1
    ReDim Preserve the_array(array_size)
    the_array(array_size) = Item

End Sub

%>