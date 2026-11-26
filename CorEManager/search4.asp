
<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->



<%

ok = OFV.CheckUsuario()
If ok <> "Y" Then
  uv = OFV.uv
  If uv = "NO" Then ok = OFV.ValidarUsuario
  shtml = ok
Else
  response.expires = 0
 CantRegMostrarx = 8
 CantRegMoverx = CDbl(Request.querystring("CantRegMover"))
 If CantRegMoverx = "" Or CStr(CantRegMoverx) = "0" Then CantRegMoverx = 1
 
  vuelta = OFV.vueltaasp("../", CantRegMostrarx, CantRegMoverx)

  

  
 ' response.write Request.form("buscar") & "---------"
  search_phrase = Request.querystring("YIDER")
  if search_phrase="" then
  search_phrase = Request.form("buscar")
  end if
  
  xxx=armaglos()
 ' response.write  search_phrase
  If search_phrase <> "" Then
    Session("search_phrase") = search_phrase
  Else
    search_phrase = Session("search_phrase")
  End If
  

    

  
  shtml = OFV.FormHeader(Session("FormN"))
  shtml = shtml & CorPnlFnc("N", "")
  shtml = shtml & "<script languaje=javascript src='incfile/tooltip.js'></script>"
  Formulario = "" 'ofv.Encabezados("info")
  
  
  
    m_search_style = "font-family:Arial; font-size:12pt; color:#000000;font-weight: bold"
    m_style_you_searched_for = "font-family:Arial; font-size:10pt; color:#000000"
    m_style_title = "font-size:10pt; font-family:Arial;color:#800000;font-weight: bold;text-decoration: underline;font-style: italic"
    m_style_text = "font-size:9pt; font-family:Arial;font-style: italic"
    m_style_noise_word = "font-size:10pt; font-family:Arial; color:#6f6f6f"
    m_style_url = "font-size:10pt; font-family:Arial; color:green"
    m_style_br = "font-size:10pt; font-family:Arial"
    m_style_more_results_text = "font-size:10pt; font-family:Arial;color:#000000;"
    m_style_more_results_link = "font-size:10pt; font-family:Arial;color:#000000;text-decoration: underline"
    m_style_more_results_link_Next = "font-size:12pt; font-family:Arial; font-weight:bold;color:#000000;text-decoration: underline"
  
  
   Set findd = Server.CreateObject("corsearch.search")
   'Set cnbsq = OFV.conectar(ofv.strconn7)
   Set cnbsq = OFV.conectar(strconn8)
    
   Set findd.OFV = OFV
   Set findd.Session = Session
   Set findd.response = response
   Set findd.Server = Server
   set findd.Cnb = cnbsq
   findd.m_trailing_words = 10
   findd.search_phrase = search_phrase
   
'ddd= findd.GetResultsFullText(cstr(search_phrase))

   if findd.noisewords=0 then
   tablab =  "<DIV id='DTBUSCAR' STYLE='float: left; z-index:1;overflow:auto;width:98%;height:50px;border-style:solid;border-width:2;border-color:#ffffff;' >"
      
             sInputT = "<input type='text' class='dt' name='BUSCAR' id='textot' value=""" & search_phrase & """ size='140' style='border-style:solid;border-width:1;border-color:#666666;' >"
             
      
              celda = ""
              celda = celda & ofv.GenCelda("", "UT  ", "left", "30%", "", "", "", sInputT, "si")
      
              sInput = ofv.GenerarInput("", "Buscar", "submit", "", "5", "at STYLE='color:#ffffff;font-weight:bold;background-color:#990000;border-style:solid;border-width:1;border-color:#dddddd;height:20px;cursor:pointer;' ", "")
              celda = celda & ofv.GenCelda("", "vt", "", "70%", "", "", "", sInput, "si")
      
              formx = ofv.GenRow("", "", "", "", "", "", "", celda, "si")
      
              sTarget = ""
      
      
             
                  sAccion = "search4.asp"
         
      
              Fila = ofv.GenForm("Subconsulta", "", sAccion, sTarget, "post", "", formx, "si")
      
              tablab = tablab & ofv.GenTabla("", "aaa   STYLE='border-style:solid;border-width:1;border-color:#dddddd;' ", "", "98%", "", "0", "", "-2", "0", "", "", Fila, "si")
      
        tablab = tablab & "</DIV>" & Chr(13)
   shtml=shtml &  tablab & "<b><p style='" & m_search_style &"'>Palabras excludias de la busqueda : " & search_phrase & "</p></b><br>"
   else
  
   Call findd.findDoc(CInt(CantRegMostrarx), CInt(CantRegMoverx))
    
   
     findd.xrs.MoveFirst

     
 
 Do While Not findd.xrs.EOF
 
             
             xincodobj = findd.xrs.Item(3)
             sActcon = findd.xrs.Item(2) 
             extracted =findd.xrs.Item(1)
             extracted2 =findd.xrs.Item(4)
            
   '
              title_highlighted = findd.xrs.Item(0)
              url = "#Foo"
  '
               formularioz="" '"<b>&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;</b>"
              If Len(Trim(extracted2)) > 0 Then
	                                 formularioz="<b>&nbsp;&nbsp;&nbsp;&nbsp;  &nbsp;&nbsp;</b>"     
	      	                       xrefs=split(extracted2,"#XXXX#")
	      	                       if isarray(xrefs) then
	      	                       for i=o to ubound(xrefs)
	      	                         XREFZ=SPLIT(XREFS(I),"XXSOLXX")
	      	                         clickn=chr(34) &" AbrirCons('"&XREFZ(0)&"','pcd');return false;" & chr(34)
	      	                        Formularioz = Formularioz  & "<a  onMouseout=""hideddrivetip()"" onMouseover=""ddrivetip('"&GETSOLAPA(XREFZ(1))&"',150,'#dddddd')"" class=xxx onclick="& clickn& " href=""" & url & search_maintain_search_term & """ style=""" & m_style_title & """>&nbsp;" & i+1 &"&nbsp;</a></span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" 
	      	                       next
	      	                       
	                      End If
                end if
              clickn=chr(34) &" AbrirCons('"&sActcon&"','pcd');return false;" & chr(34)
              Formulario= Formulario & NewLine(0) & NewLine(0) & "<a  class=xxx onMouseout=""hideddrivetip()"" onMouseover=""ddrivetip('"&armatip(xincodobj)&"',250,'#dddddd')"" onclick="&clickn&" href=""" & url & search_maintain_search_term & """ style=""" & m_style_title & """>" & title_highlighted &"</a>" & formularioz & "<br>"
             if Len(Trim(extracted)) > 0 then
	                 Formulario= Formulario & NewLine(0) & "<br><span style=""" & m_style_text & """>" & extracted & "</span>"
             end if
               
              url = "<hr style='border-bottom: dashed 2px #909090;'>" '& Session("CMFPNIPFT")
             Formulario = Formulario & NewLine(0) & "<span style=""" & m_style_url & """>" & url & "</span>"
           ' Formulario = Formulario & NewLine(0) & "<span style=""" & m_style_br & """></span><br>"
           findd.xrs.MoveNext
 Loop
   If findd.xrs.AbsolutePage < findd.xrs.PageCount Then
    pagsig = findd.xrs.AbsolutePage + 1
    xclase="bietx"
   Else
   pagsig = findd.xrs.PageCount
   xclase="bietxG"
   End If
   tablab =  "<DIV id='DTBUSCAR' STYLE='float: left; z-index:1;overflow:auto;width:98%;height:50px;border-style:solid;border-width:2;border-color:#ffffff;' >"
   
          sInputT = "<input type='text' class='dt' name='BUSCAR' id='textot' value=""" & search_phrase & """ size='140' style='border-style:solid;border-width:1;border-color:#666666;' >"
          
   
           celda = ""
           celda = celda & ofv.GenCelda("", "UT  ", "left", "30%", "", "", "", sInputT, "si")
   
           sInput = ofv.GenerarInput("", "Buscar", "submit", "", "5", "at STYLE='color:#ffffff;font-weight:bold;background-color:#990000;border-style:solid;border-width:1;border-color:#dddddd;height:20px;cursor:pointer;' ", "")
           celda = celda & ofv.GenCelda("", "vt", "", "70%", "", "", "", sInput, "si")
   
           formx = ofv.GenRow("", "", "", "", "", "", "", celda, "si")
   
           sTarget = ""
   
   
          
               sAccion = "search4.asp"
      
   
           Fila = ofv.GenForm("Subconsulta", "", sAccion, sTarget, "post", "", formx, "si")
   
           tablab = tablab & ofv.GenTabla("", "aaa   STYLE='border-style:solid;border-width:1;border-color:#dddddd;' ", "", "98%", "", "0", "", "-2", "0", "", "", Fila, "si")
   
        tablab = tablab & "</DIV>" & Chr(13)
   
   if findd.xrs.recordcount=0 then
   pospag=0
   else
   pospag=findd.xrs.AbsolutePage
   end if
        
   xinfo=tablab & "<b><p style='" & m_search_style &"'>Pagina " & pospag & " de "& findd.xrs.PageCount & " - Total Documentos encontrados :"& findd.xrs.recordcount&"</p></b><br>"
   sAccion = "search4.asp?cantregmover=" & pagsig 
   SIGUIENTE = "<td><input  type=BUTTON  size=10%  value='SIGUIENTE' class="&xclase&"  onclick='document.all.TestFrm.submit();'   ></td>"
   SIGUIENTE = SIGUIENTE & OFV.genform("TestFrm", "AAA STYLE='display:none;' ", sAccion, "", "post target=_self", "", "", "si")
   
   
    
      If findd.xrs.AbsolutePage = 1 Then
      pagant = 1
      xclase="bietxG"
      Else
      xclase="bietx"
      pagant = findd.xrs.AbsolutePage - 1
   End If
   
   sAccion = "search4.asp?cantregmover=" & pagant
   ANTERIOR = "<td><input  type=BUTTON  size=10%  value='ANTERIOR' class="&xclase &"   onclick='document.all.TestFrm2.submit();'   ></td>"
   ANTERIOR = ANTERIOR & OFV.genform("TestFrm2", "AAA STYLE='display:none;' ", sAccion, "", "post target=_self", "", "", "si")
   
   
  ' sHTML = sHTML & ofv.botoneraESSX("0","",CantRegAMover,CantRegAMostrar,"no","no",posi,"","","no",Agregar,Listar,"")
  'else
 '
   shtml = shtml & xinfo & Formulario & OFV.botoneraESSX("0", "", "", "", "no", "no", posi, "", "", "no", ANTERIOR, SIGUIENTE, "")
   shtml = shtml & "</body></html>"
   
end if   
End If

set  findd.xrs=nothing
set findd=NOTHING
call ofv.cerrarconn(cnbsq)
response.write shtml



Function armatip(idC)
    TEXTO = ""
    on error resume next
    Set CN = OFV.conectar(OFV.strconn2)
    strSQL = "select a.temadescri,a.stemadescri,b.*,c.estado AS ESTADOD,b1.codiden from CSRELPATHDESCRIex a "
    strSQL = strSQL & " inner join objetos b1 on(b1.codobjeto= a.cobj)"
    strSQL = strSQL & " inner join objprty b on(b.codobjeto= a.cobj)"
    strSQL = strSQL & " inner join estados c on(c.objid=b.estado) where a.cobj=0" & idC
    'RESPONSE.WRITE STRSQL
    Set recordsetip = CN.Execute(strSQL)
    'response.write recordsetip("autor")
    'xxx=getinfouser(recordsetip("autor"))

    TEXTO = "<B><U>TEMA</U> :</b>&nbsp;&nbsp;" & recordsetip("TEMADESCRI") & "<BR><B><U>SUBTEMA</U> :</b>&nbsp;&nbsp;" & recordsetip("STEMADESCRI") '& "<BR>"
    'TEXTO = TEXTO & "<B><U>AUTOR</U>:&nbsp;&nbsp;</B>" & Replace(Session("SGU_NOMBRE"), " ", "&nbsp;") '& "&nbsp;<B>SUBTEMA : </b>"& recordsetip("STEMADESCRI")& "<BR>"
   ' TEXTO = TEXTO &  "&nbsp;<B>SUBTEMA : </b>"& recordsetip("STEMADESCRI")& "<BR>"
    TEXTO = TEXTO & "<br><B><U>ESTADO</U>&nbsp;&nbsp;:</B>" & recordsetip("ESTADOD") & "<br><U><b>VIGENCIA</U> : </b>" & fechadorcor(CDate(CLng("0" & recordsetip("fecvigen")) + CLng(CDate("01-01-1900"))), "") & "<BR>"
    If recordsetip("VERSION") = "O" Then
        Version = "(O)"
    Else
        Version = "V (" & recordsetip("NROVER") & ")"
    End If
    TEXTO = TEXTO & "<B><U>COD. IDENT.</U> :&nbsp;&nbsp;</B>" & recordsetip("CODIDEN")
    TEXTO = TEXTO & "<BR><B><U>VERSIÓN</U>  :&nbsp;&nbsp;</B>" & Version & "  " '& idC
    armatip = "<b>" & TEXTO & "</b>"
    Call OFV.cerrarconsulta(recordsetip)
End Function


Private Function NewLine(spaces)
    NewLine = vbCrLf & Space(spaces)
End Function


Public Function ReplaceRegExp(ByVal strString, strPattern, strReplace)
        Dim RE: Set RE = New RegExp
    
        With RE
            .Pattern = strPattern
            .Global = True
            .IgnoreCase = True
            ReplaceRegExp = .Replace(strString, strReplace)
        End With
    End Function
    
    
    FUNCTION GETSOLAPA(ID)
    ID=CSTR(ID&"")
    
       Select Case ID
                    
                    Case "1"
                    tit = "Solapa propiedades"
                    Case "2"
                    tit = "Solapa consideraciones"
                    Case "3"
                    tit = "Solapa sectores"
                    Case "4"
                    tit = "Solapa eventos"
                    
                    Case "5"
                    tit = "Solapa procedimientos"
                    Case "6"
                    tit = "Solapa descripcion"
                    Case "7"
                    tit = "Solapa controles"
                    Case "8"
                    tit = "Solapa esquemas"
                    Case "9"
                    tit = "Solapa tablas"
                    Case "10"
                    tit = "Solapa documentos"
                    Case "11"
                    tit = "Solapa contenido"
                    Case "12"
                    tit = "Solapa cursograma"
                    Case "13"
                    tit = "Solapa contratos"
                    Case "14"
                    tit = "Solapa faq"
                    Case "15"
                    tit = "Solapa general"
                    Case Else
                    tit = "Sin descripción"
                    
                    
                    
                End Select
       GETSOLAPA=TIT         
    END FUNcTION
    
%>
