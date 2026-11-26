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


  LNombre = request.querystring("Nombre")
  sNombre = request.querystring("Nombre")
  sNombre = ofv.convertircar(sNombre,"_"," ")

  sID     = request.querystring("ID")

  sitem   = ofv.convertircar(request.querystring("sitem"),"_"," ")
  stitem  = ofv.convertircar(request.querystring("stitem"),"_"," ")
  sticono = request.querystring("sticono")
  stnivel = request.querystring("stnivel")
  sformato = request.querystring("formato")

  if request.querystring("Menu") <> "" then Session("Menu") = request.querystring("Menu")
  if sNombre = "" Then sNombre = UCase(sTipo)

  ssTipo                = sTipo
  session("sNivelCero") = sTipo
  session("sEstado")    = "APROBADA"
  session("sEstadon")   = "3"


'RESPONSE.WRITE "222"

  if sformato="ABR" THEN
     WSPDIVN = "DTTEMAINT"       ' ID DEL DIV DE VENTANA PARA MOSTRAR
     sHTML = ""
     sHTML = sHTML & "<HTML>"
     sHTML = sHTML & "<style><link rel=""stylesheet"" type=""text/css"" href=""../fonts/Corfont/Corflaticon.css""></STYLE>"
     sHTML = sHTML & "<BODY onload=""window.top.LDVENTANA('" & WSPDIVN & "');"" >"
  ELSE
  sHTML      = ofv.FormHeader(session("FormN"))
  END IF


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



               sVerEstado2 = "  and Objetos.objnivel >= " & Session("SegNivel") & " "

              if SESSION("EMPMUN") = "S" THEN

                if SESSION("GLBMUN") = 1 THEN
                   sverestado2 = sverestado2 & " and Objetos.OBJUN = " & Session("UN") & " "
                ELSE
                   sverestado2 = sverestado2 & " and "
                   sverestado2 = sverestado2 & " ( (OBJETOS.OBJUN = " & Session("UN") & ") OR "
                   sverestado2 = sverestado2 & "   ( OBJETOS.OBJUN = 1 AND "
                   sverestado2 = sverestado2 & "     OBJETOS.CODOBJETO = OBJMEMPALC.CODOBJETO AND "
                   sverestado2 = sverestado2 & "     OBJMEMPALC.CODEMPRE = " & Session("UN") & ") "
                   sverestado2 = sverestado2 & " ) "
                END IF

              end if


                sverestado2 = sverestado2 & " and Objprty.Estado = 3 "

                        xStrSqla = ""
                        xStrSqla  = xStrSqla & " and (grpusr = 'USR-" & UCase(Session("usuario")) & "' "

'''''''''''''''''''''''''''''''''''''''''''
                      if  Session("DELLOG") = "S" then

                          XMFTGRP = split(session("SegGrpLst"),chr(9))
                          xcount  = ubound(XMFTGRP) + 1
                          XMFSEP  = ""
                          xStrSqla = xStrSqla & " or grpusr in ("

                          for XMFI = 0 to ubound(XMFTGRP)
                              xStrSqla = xStrSqla  & XMFSEP & "'GRP-" & XMFTGRP(XMFI) & "'"
                              XMFSEP  = ","
                          next

                          xStrSqla = xStrSqla & ") "

                      else
                          xStrSqla = xStrSqla & " or grpusr = 'GRP-" & UCase(Session("seggrp")) & "' "
                      end if

''''''''''''''''''''''''''''''''''''''''''
                        xStrSqla  = xStrSqla & ")"





 ' Strsql = "Select * From Tema where aplic = '" & wkaplic & "' or aplic = 'GLB' order by inda, descri"

'  Strsql = "SELECT X.DESCRI as tema , XT.DESCRI as subtema ,isnull(Z.expr7,'----------'), isnull(Z.NOMBRE,'------------------------------------'),len(Z.TSEC) - len( replace(Z.TSEC, '.' , '' ) ) -1  NIVELC, Z.TSEC, X.CODTEMA,"
'  strsql =strsql & " XT.CODPRODSERV, Z.COBJ, Z.TOBJ ,x.aplic "

'  strsql =strsql & " FROM TEMA X "
 ' strsql =strsql &  "LEFT JOIN PRODSERV XT ON (X.CODTEMA = XT.CODTEMA)"
'  strsql =strsql & " LEFT JOIN ("


' strsql =strsql & "  select D.STEMA, D.COBJ, D.TOBJ, D.expr7,isnull( isnull(d.expr2 + '/','') + isnull(d.expr3 + '/','') + ( isnull(d.expr4 + '/','') + isnull(d.expr5 + '/','')  + isnull(d.expr6 + '/','') "
' strsql =strsql & "              + isnull(d.expr7 + '/',''))  ,'------------------------------------') as nombre ,D.TSEC   from csrelpathdescriex D "
' strsql =strsql & " where D.tipocpt='C' "
' strsql =strsql & "             ) Z ON (XT.CODPRODSERV = Z.STEMA)"
' strsql =strsql & " where (x.aplic = '" & wkaplic & "' or x.aplic = 'GLB') and X.PSTIPO IN ('C','R') "

' strsql =strsql & "   ORDER BY X.DESCRI, XT.DESCRI, Z.NOMBRE "

 ' cuneo
 Strsql = "Select codtema,descri,pstipo,psurl,aplic From Tema where (aplic = '" & wkaplic & "' or aplic = 'GLB') and PSTIPO IN ('C','R') order by inda, descri"

 'mendez no trae link a internet ni pedidos ni solictud ni formularios
 Strsql = "Select distinct a.codtema,a.descri,a.pstipo,a.psurl,a.aplic,a.inda From Tema a "
 Strsql = strsql & " inner join prodserv b on (a.codtema = b.codtema and b.marcaps <> 'PEDIDO' and b.marcaps <> 'SNN') "
 Strsql = strsql & " where a.PSTIPO <> 'K' AND (a.aplic = '" & wkaplic & "' or a.aplic = 'GLB') order by a.inda, a.descri"

'response.write strsql
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if  rs.eof or session("un") = 0 then

      Celda   = ""
      Celda   = Celda & ofv.GenCelda("","ttt","","2%","","","","&nbsp;","si")
      Celda   = Celda & ofv.GenCelda("","ttt","","90%","","","","No hay elementos asociados","si")

      Fila    = fila & ofv.GenRow("","","","","25","","",Celda,"si")

  else

      Celda   = ""



    Fila=""
    codtema=0
    isubtema=0
    ISUBCARPETA=""

	    If (SESSION("UN") & "") = "" Then


	        xxxx = NEWEMPRE("NO")


	        SEMPRE = ATREMP(0, 0)
	        SESSION("UN") = SEMPRE & ""

	        If (SESSION("UN") & "") = "" Then SESSION("UN") = 1

      else
       '  xxxx = NEWEMPRE(SESSION("UN"))


       End If

    do until rs.eof

       xxnova = "S"
       Strsql = "Select * From ATRMTEMAST WHERE ATRMCODST = " & rs("codtema") & " AND ATRMTIPOST = 'T' AND ATRMCODES = " & session("UN")
'      response.write strsql
       set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)

       if not  rs1.eof then
          if rs1(4) = 0 then xxnova = "N"
       end if

       if xxnova = "S" then
       sxlink = "ObjTreeSubTemaCons.asp?"
       sHref  = ""
    '   if codtema <> rs("codtema") then
    '   codtema=rs("codtema")

        if rs("aplic") = "GLB" then
	   		sImage = "&nbsp;&nbsp;&nbsp;&nbsp;<font face='CorFlaticon' size=2 color='#f7aa22'>&#0252;</font>&nbsp;&nbsp;"
	   	else
	   		sImage = "&nbsp;&nbsp;&nbsp;&nbsp;<font face='CorFlaticon' size=2 color='#f7aa22'>&#060;</font>&nbsp;&nbsp;"
	    end if


      '    xTrColor="#eeeeff"
          sitem2   = ofv.ConvertirCar(rs("DESCRI")," ","_")
		 sxlink  = sxlink & "ID=" & rs("codtema") & "&sitem=" & sitem2 & "&stitem=" & REPLACE(stitem," ","_")
	      sxlink  = sxlink & "&stnivel=0&tipo=SUBTEMAS&TIPOVISTA=CONS"
	      sHref   = sxlink

          starget=""
	      if rs("pstipo")="R" THEN
	          sImage  = "<font face='corflaticon' size=3 color='#990000'>4</font>"
	        sHref   = rs("psURL") & " "
	        if instr(ucase(sHref),"HTTP://") = 0 then sHref   = "HTTP://" & sHref
	        starget=" targer=_new"
	      END IF


	   '  sLink2   = ofv.GenLink("","cmb" ,sHref,starget,"","","<u>" & replace(rs("DESCRI")," ","&nbsp;") & "&nbsp;</u>","si")

	       sLink2   = ofv.GenLink("","nt" ,sHref,starget,"","",replace(rs("DESCRI")," ","&nbsp;") & "&nbsp;","si")


          celda= ofv.GenCelda("","ut bgcolor='" & xTrColor & "' xstyle='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#cccccc;' ","","10%","","","",sImage & slink2,"si")

         Fila    = Fila & ofv.GenRow("","aaa ","","","20","","",Celda,"si")
     '  end if

  '     if  isubtema <> rs("subtema") and false then
  '          stitem   =  "XXSUBTEMAXX"
  '         xTrColor="#FFFFFF"
	'	   sitem    =  rs("subtema")
  ''         sitem2   = ofv.ConvertirCar(rs("subtema")," ","_")
  '         sxlink  = sxlink & "ID=" & rs("codprodserv") & "&sitem=" & sitem2 & "&stitem=" & REPLACE(stitem," ","_")
	'	   sxlink  = sxlink & "&stnivel=1"
	'	   sHref   = sxlink
  '         sLink2   = ofv.GenLink("","cmb",sHref,"","","","<u>" & sitem & "&nbsp;</u>","si")
''
  '         celda=ofv.GenCelda("","ut bgcolor='" & xTrColor & "'  ","","10%","","","","&nbsp;","si")
  '         celda= ofv.GenCelda("","ut bgcolor='" & xTrColor & "' style='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#cccccc;' ","","90%","","","","&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & sLink2,"si")
  '         Fila    = Fila & ofv.GenRow("","aaa ","","","20","","",Celda,"si")
  '         isubtema=rs("subtema")
  '     end if
'
  '    nivelx=""
  '    IF RS("NIVELC") & "" <> "" and false THEN
  '     xTrColor="#EEEEEE"
  '     sxlink  = "Consdocumentosweb.asp?"
  '     stitem   =  rs("TOBJ")
  '     sitem    = RS(2)
  '     sitem2   = ofv.ConvertirCar(rs(2)," ","_")
  '     sxlink  = sxlink & "ID=" & rs("cobj") & "&sitem=" & sitem2 & "&stitem=" &stitem
  '     sxlink  = sxlink & "&stnivel=" & RS("NIVELC")
	'   sHref   = sxlink
  '     sLink2   = ofv.GenLink("","cmb ",sHref,"","","","<u>" & sitem & "&nbsp;</u>","si")

  '     for i=0 to RS("NIVELC")
  '     nivelx="&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & nivelx
  '     next


    '   IF RS("NIVELC")=1 THEN
	 '   celda= ofv.GenCelda("","ut bgcolor='" & xTrColor & "'  ","","10%","","","","&nbsp;","si")
	 '   celda= ofv.GenCelda("","ut bgcolor='" & xTrColor & "' style='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#cccccc;' ","","90%","","","",nivelx & sLink2,"si")
	'    Fila    = Fila & ofv.GenRow("","aaa ","","","20","","",Celda,"si")
	 '   ISUBCARPETA=rs(2)
     '  END IF

    '  IF RS("NIVELC")=2 THEN
	 '   celda= ofv.GenCelda("","ut bgcolor='" & xTrColor & "'  ","","10%","","","","&nbsp;","si")
	 '   celda= ofv.GenCelda("","ut bgcolor='" & xTrColor & "'  ","","90%","","","",nivelx & sLink2,"si")
	 '   Fila    = Fila & ofv.GenRow("","aaa ","","","20","","",Celda,"si")
	 ''   ISUBCARPETA=rs(2)
     '  END IF
'
     'IF RS("NIVELC")=3 THEN
	 '   celda= ofv.GenCelda("","ut bgcolor='" & xTrColor & "'  ","","10%","","","","&nbsp;","si")
	 '   celda= ofv.GenCelda("","ut bgcolor='" & xTrColor & "'  ","","90%","","","",nivelx & sLink2,"si")
	 '   Fila    = Fila & ofv.GenRow("","aaa ","","","20","","",Celda,"si")
	 '   ISUBCARPETA=rs(2)
     '  END IF

    '  END IF
    end if
      RS.MOVENEXT
    loop

  end if






    Call OFV.cerrarconsulta(rs)
    Call OFV.cerrarconn(cn)


 if sformato = "ABR" THEN


 sHTML = sHTML & ofv.GenTabla("","aaa   ","","95% ","","0","","0","0","","",Fila,"si")
 ELSE
  sHTML = sHTML & ofv.GenTabla("","aaa   ","","60% ","","0","","0","0","","",Fila,"si")
  Agregar = ""
  Buscar  = ""
  Celda   = ""

  if session("SMPLMENU") = "S" then
     sAccion = "novedmenu.asp"
'     Celda   = ofv.BotonVolver(sAccion)

      xxxx    = CorGenBotonXp(16, "NO", "home", "tahoma", "NO", "Home", "_self", "Home")
      Celda  = ofv.GenBoton(16, "esp1Bot", "BT", "S", "", "S", "A", "N", "esp1Frm", sAccion & "")


  end if


  sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","",Agregar,Buscar,celda)
  END IF

  sHTML = sHTML & "</body></html>"




'''''''''''''

end if

response.write sHTML





%>
