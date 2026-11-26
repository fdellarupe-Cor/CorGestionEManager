
<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->

<!-- #INCLUDE FILE="XoiseRecordset.inc" -->



<%

ok = OFV.CheckUsuario()
If ok <> "Y" Then
  uv = OFV.uv
  If uv = "NO" Then ok = OFV.ValidarUsuario
  shtml = ok
Else
  response.expires = 0
 CantRegMostrarx = 10
 CantRegMoverx = CDbl(Request.querystring("CantRegMover"))
 If CantRegMoverx = "" Or CStr(CantRegMoverx) = "0" Then CantRegMoverx = 1
 
  vuelta = OFV.vueltaasp("../", CantRegMostrarx, CantRegMoverx)

  

  set ss=new XoiseRecordset
  
  search_phrase = Request.querystring("YIDER")
  
  If search_phrase <> "" Then
    Session("search_phrase") = search_phrase
  Else
    search_phrase = Session("search_phrase")
  End If
  

    
    
    
    
  
  shtml = OFV.FormHeader(Session("FormN"))
  shtml = shtml & CorPnlFnc("N", "")
  shtml = shtml & "<script languaje=javascript src='incfile/tooltip.js'></script>"
  Formulario = "" 'ofv.Encabezados("info")
  
  
  
    m_search_style = "font-family:Arial; font-size:12pt; color:#000000"
    m_style_you_searched_for = "font-family:Arial; font-size:10pt; color:#000000"
    m_style_title = "font-size:10pt; font-family:Arial;color:#ff0000;font-weight: bold;text-decoration: underline"
    m_style_text = "font-size:9pt; font-family:Arial"
    m_style_noise_word = "font-size:10pt; font-family:Arial; color:#6f6f6f"
    m_style_url = "font-size:10pt; font-family:Arial; color:green"
    m_style_br = "font-size:10pt; font-family:Arial"
    m_style_more_results_text = "font-size:10pt; font-family:Arial;color:#000000;"
    m_style_more_results_link = "font-size:10pt; font-family:Arial;color:#000000;text-decoration: underline"
    m_style_more_results_link_Next = "font-size:12pt; font-family:Arial; font-weight:bold;color:#000000;text-decoration: underline"
  
  
   Set findd = Server.CreateObject("corsearch.search")
    
   Set findd.OFV = OFV
   Set findd.Session = Session
   Set findd.response = response
    Set findd.Server = Server
   findd.strconn = strconn8
   findd.search_phrase = search_phrase
   
    set cn = ofv.conectar(strconn8)
    ss.PageSize = 10
    ss.AbsolutePage = 1
    ss.conn=cn
    ss.sql="select * from yider where [key]='111139'"
    
    response.write  ss.eof
    
    'response.write CInt(CantRegMostrarx) &"  " & CInt(CantRegMoverx)
   Call findd.findDoc(CInt(CantRegMostrarx), CInt(CantRegMoverx))
   'if not findd.xrs.eof then
     findd.xrs.MoveFirst

     
 
 Do While Not findd.xrs.EOF
 
             
             xincodobj = findd.xrs.Item(3)
             sActcon = findd.xrs.Item(2)
             extracted = findd.xrs.Item(1)
   '
              title_highlighted = findd.xrs.Item(0)
              url = "#Foo"
  '
              clickn=chr(34) &" AbrirCons('"&sActcon&"','pcd');return false;" & chr(34)
              Formulario= Formulario & NewLine(0) & NewLine(0) & "<a  class=xxx onMouseout=""hideddrivetip()"" onMouseover=""ddrivetip('"&armatip(xincodobj)&"')"" onclick="&clickn&" href=""" & url & search_maintain_search_term & """ style=""" & m_style_title & """>" & title_highlighted &"</a>"
             
             If Len(Trim(extracted)) > 0 Then
                 Formulario = Formulario & NewLine(0) & "<br><span style=""" & m_style_text & """>" & extracted & "</span>"
              End If
              url = "<hr>"
             Formulario = Formulario & NewLine(0) & "<br><span style=""" & m_style_url & """>" & url & "</span>"
            Formulario = Formulario & NewLine(0) & "<span style=""" & m_style_br & """></span>"
 findd.xrs.MoveNext
 Loop
   If findd.xrs.AbsolutePage < findd.xrs.PageCount Then
    pagsig = findd.xrs.AbsolutePage + 1
    xclase="bietx"
   Else
   pagsig = findd.xrs.PageCount
   xclase="bietxG"
   End If
   
   xinfo="<p style='" & m_search_style &"'>Pagina " & findd.xrs.AbsolutePage & " de "& findd.xrs.PageCount & " - Total Docuentos encontrados :"& findd.xrs.recordcount&"</p><br>"
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
   
   
End If
response.write shtml



Function armatip(idC)
    TEXTO = ""
    Set CN = OFV.conectar(OFV.strconn2)
    strSQL = "select a.temadescri,a.stemadescri,b.*,c.estado AS ESTADOD,b1.codiden from CSRELPATHDESCRIex a "
    strSQL = strSQL & " inner join objetos b1 on(b1.codobjeto= a.cobj)"
    strSQL = strSQL & " inner join objprty b on(b.codobjeto= a.cobj)"
    strSQL = strSQL & " inner join estados c on(c.objid=b.estado) where a.cobj=0" & idC
    'RESPONSE.WRITE STRSQL
    Set recordsetip = CN.Execute(strSQL)
    'response.write recordsetip("autor")
    'xxx=getinfouser(recordsetip("autor"))

    TEXTO = "<B><U>TEMA</U> :</b>&nbsp;&nbsp;" & recordsetip("TEMADESCRI") & "<BR><B><U>SUBTEMA</U> :</b>&nbsp;&nbsp;" & recordsetip("STEMADESCRI") & "<BR>"
    TEXTO = TEXTO & "<B><U>AUTOR</U>:&nbsp;&nbsp;</B>" & Replace(Session("SGU_NOMBRE"), " ", "&nbsp;") '& "&nbsp;<B>SUBTEMA : </b>"& recordsetip("STEMADESCRI")& "<BR>"
    TEXTO = TEXTO & "<br><B><U>ESTADO</U>&nbsp;&nbsp;:</B>" & recordsetip("ESTADOD") & "<br><U><b>VIGENCIA</U> : </b>" & fechadorcor(CDate(CLng("0" & recordsetip("fecvigen")) + CLng(CDate("01-01-1900"))), "") & "<BR>"
    If recordsetip("VERSION") = "O" Then
        Version = "(O)"
    Else
        Version = "V (" & recordsetip("NROVER") & ")"
    End If
    TEXTO = TEXTO & "<B><U>COD. IDENT.</U> :&nbsp;&nbsp;</B>" & recordsetip("CODIDEN")
    TEXTO = TEXTO & "<BR><B><U>VERSIÓN</U>  :&nbsp;&nbsp;</B>" & Version
    armatip = "<b>" & TEXTO & "</b>"
    Call OFV.cerrarconsulta(recordsetip)
End Function


Private Function NewLine(spaces)
    NewLine = vbCrLf & Space(spaces)
End Function

%>
