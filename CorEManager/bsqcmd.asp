<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires = 0
  response.expires=-1441
  Response.AddHeader "PRAGMA", "NO-CACHE"

   moverreg = clng("0" & request.querystring("canTregmover"))
   tope=clng("0" & request.querystring("tope"))




     WSPDIVN = "DTNOVINT8"       ' ID DEL DIV DE VENTANA PARA MOSTRAR
     sHTML = ""
     sHTML = sHTML & "<HTML>"
     sHTML = sHTML & "<BODY onload=""window.top.LDVENTANA('" & WSPDIVN & "');"" >"


''''''''''''''




    Set cn = OFV.conectar(OFV.strconn2)

  Fila    = ""



   xTrColor = "#FFFFFF"

  stitem = "Tema"
  sIcono = "ICorTema.jpg"
  sNuevoTipo = "SUBTEMAS"
  nNivel = 0

  sID = 0
  nv = 0
  cborderB = Session("CBRDBOT")
  cborderl = Session("CBRDLFT")
  borderB = ""
  borderl = ""


  StrSql = "SELECT distinct a.codobjeto,a.tipobjeto,e.tobjdescri + ':' +e.expr7 as DESCRI FROM OBJPRTY a "
StrSql = StrSql & " INNER JOIN  OBJOBSCMT B ON (B.tipo='CMD' AND B.ACTPEND is null  AND B.CODOBJETO=a.CODOBJETO AND B.TIPOBJETO=a.TIPOBJETO ) "
StrSql = StrSql & " INNER JOIN  TIPOBJETO C ON (C.COMENTARIOS='S'  AND A.TIPOBJETO=C.TIPOBJETO ) "
StrSql = StrSql & " INNER JOIN  objsdkdet D ON (D.tipoSEG='9' AND D.CODOBJETO=0 AND A.TIPOBJETO=D.TIPOBJETO ) "
StrSql = StrSql & " INNER JOIN  csrelpathdescriex e ON (a.codobjeto=e.cobj ) "
StrSql = StrSql & " WHERE  (a.ESTADO = 3 )"
StrSql = StrSql & " And (( b.codusr='" &  SESSION("usuario") & "' ) or (a.autor='"& SESSION("usuario") &"'))"
StrSql = StrSql & " And ("
StrSql = StrSql & " (D.GRPUSR='USR-" & SESSION("usuario") & "')"
StrSql = StrSql & " or "
StrSql = StrSql & " (D.GRPUSR IN ('GRP-" & Session("SegGrp") &"'"
  wktgrupos = SPLIT(session("SegGrpLst"), CHR(9))

		   FOR wkti = 0 TO UBOUND(wktgrupos)


		 StrSql = StrSql & ",'GRP-" &  wktgrupos(wkti)  &"'"


		NEXT



	StrSql = StrSql & ")))"

set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

 if Not rs.EOF then rs.move(moverreg)

   Celda   = ""



    Fila=""
    POS=0
      xTrColor="#FFFFFF"
    do WHILE (NOT  rs.eof) AND (POS < tope)
       sxlink  = "ObjTreeSubTemaCons.asp?"
       sHref=""



         sActcon = "ConsultarcMD.asp?Codigo=" & rs(0) & "&Tipo=" & rs(1)
         sHref = "&nbsp;"


		sFuncion = " onclick=" & Chr(34) & " AbrirVentana('" & sActcon & "'); return false;" & Chr(34)

         xarc = 	"<B><font color='#000000' size=2 >&#0149;</font></B>"

         sLink2 = ofv.GenLink("", "nt  title='Consultar Documento...' ", sHref, "", "", sFuncion,   "&nbsp;" & replace(rs("DESCRI")," ","&nbsp;"), "si")

	    ' sLink2   = ofv.GenLink("","cmb" ,sHref,starget,"","","<u>" & replace(rs("DESCRI")," ","&nbsp;") & "&nbsp;</u>","si")


          celda= ofv.GenCelda("","ut bgcolor='" & xTrColor & "' xstyle='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#cccccc;' ","","10%","","","",xarc  & slink2,"si")

         Fila    = Fila & ofv.GenRow("","aaa ","","","20","","",Celda,"si")

         If xTrColor = "#FFFFFF" Then
		        xTrColor = "#eeeeff"
		       Else
		        xTrColor = "#FFFFFF"
       End If
      RS.MOVENEXT
      POS=POS+1
    loop













   sHTML = sHTML & ofv.GenTabla("","aaa   ","","95% ","","0","","0","0","","",Fila,"si")

   if rs.eof then
   sxlinksig =""
   else
   sxlinksig =  "<A id=lnksig class=nt style='display:Xnone;'  HREF='BSQcMD.asp?CANTREGMOVER=" & moverreg + tope  & "&tope="& tope&"'"
   sxlinksig = sxlinksig & " TARGET=fesquemas><font size=2 color=red face='Corflaticon'>&#0232;</font></a>"
   end if

   if moverreg - tope < 0 then
   sxlinkant=""
   else
    sxlinkant =  "<A id=lnkant class=nt style='display:Xnone;' HREF='BSQcMD.asp?CANTREGMOVER=" &  moverreg - tope  & "&tope="& tope&"'"

	sxlinkant = sxlinkant & " TARGET=fesquemas><font size=2 color=red face='Corflaticon'>&#57463;</font></a>"
   end if

	botones=  "&nbsp;&nbsp;&nbsp;&nbsp; " & sxlinkant &  "&nbsp;&nbsp;&nbsp;&nbsp;" & sxlinksig

  botones  ="<font size=2 color=red face=tahoma>Ver Más</font>" & botones
    sHTML = sHTML & botones



'''''''''''''

    Call OFV.cerrarconsulta(rs)
    Call OFV.cerrarconn(cn)
end if

response.write sHTML & "</body></html>"





%>
