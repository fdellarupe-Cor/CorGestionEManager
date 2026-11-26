<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/CorWKFNro.inc" -->
<%

ok=ofv.CheckUsuario()

if ok <> "Y" then
  sHTML = ok
else
  response.expires=0


' for each x in session.contents
' response.write x & " = " & session.contents(x) & "<br>"
'next

  CantRegMostrarx = 12
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))
  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)

  set cn = ofv.conectar(ofv.strconn5)
  set cn4 = ofv.conectar(ofv.strconn6)
  set cn2 = ofv.conectar(ofv.strconn0)

  ofv.ObtenerAtributos session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  SESSION("SEGCLIENTE") = "NO"
  SESSION("WKF") = ""

 CKT = request.querystring("CKT")
 DSD = request.querystring("DSD")

 if CKT = "" then
    CKT = session("CKT")
 else
    session("CKT") = CKT
 end if

  sHTML = ofv.FormHeader(session("FormN"))

'----------------------------------------------

  shtml = shtml & CorPnlFnc("N","")

  sHTML = sHTML &  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' >"

'-------------------------------------------------


  dim AUTabla (100,2)
  TUTabla = 0
  IUTabla = 0



 xs = "<script language=JavaScript>" & Chr(13)
    xs = xs & "function AbrirCons(Datos,nom) {" & Chr(13)
    xs = xs & " var h=570;w=790 ;"
    if  session("APLMDCLI")="BC" THEN
	xs = xs &  "if (window.screen) {w = window.screen.availWidth;h = window.screen.availHeight-60;} "
	END IF
    xs = xs & " myWin= open(Datos"

    xs = xs & "," & Chr(34) & "ConsP" & Chr(34) & " + nom," '& Chr(34)
   ' xs = xs & "'menubar=no,toolbar=no,height=h,width=w,resizable=yes'" '& Chr(34)
    xs = xs &"  'toolbar=no,scrollbars=yes,width='+w+',height='+h+',top=0,left=0,resizable=yes'" '& CHR(34)
    xs = xs & " ); " & Chr(13) & " } " & Chr(13) & "</script>" & Chr(13)




       Ghtml = ""




  StrSql = "Select * From SEGCIRCUITO WHERE SEGCKT = '" & SESSION("CKT") & "' "
  'response.write SESSION("CKT")
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then
  posi = 0

     WKFOBJTIPO  = RS(4)
     WKFOBJCODE  = RS(5)


    sform = "WKF" & rs(0)


       sc = "vt"



       celdas = ""


       sInput = rs(5)

       IF WKFOBJTIPO = "CLI" THEN
          StrSql = "Select NOMBRE From CLIENTE WHERE CODIGO = '" & rs(5) & "' "
          set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
          if not rs2.eof then
             sInput = "<B>" & RS2(0) & "</B>"
          END IF
       ELSEIF WKFOBJTIPO = "REQ" THEN
          StrSql = "Select TITULO From REQUERIMIENTO where codigo = '" & rs(5) & "'  "
          set rs2 = ofv.crearconsultaEx(StrSql,cn4,1,parametros)
          if not rs2.eof then
             sInput = "<B>" & RS2(0) & "</B>"
          END IF
       ELSE
          sInput = esqnomp (WKFOBJTIPO,WKFOBJCODE,ofv)

       END IF





       SINPUTn1 = REPLACE(SINPUT," ","&nbsp;")

       celdas = celdas & ofv.GenCelda("",sc,"","15%","","","",sInput,"si")



       sInputn2 = rs(8)
       celdas = celdas & ofv.GenCelda("",sc,"","50%","","","",sInput,"si")

 set cn2 = ofv.conectar(ofv.strconn0)

 ' StrSql = "Select * From USUARIOS WHERE ID = '" & RS(7) & "' "
'  set rs2 = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
 ' if not rs2.eof THEN

  '     sInputn3 = rs2(1)
'  END IF
       xxxx = getInfoUser(RS(7))
       if session("SGU_nombre")<> "" then    sInputn3 =  session("SGU_nombre")

        ' = rs2(1)
       celdas = celdas & ofv.GenCelda("",sc,"","50%","","","",sInput,"si")

       segCKT= RS(0)

       sinput = "<a class=tt TARGET=_blank href=WKFVER.asp?CKT=" & RS(0)
       sinput = sinput  & " ><b>VER</b></a>"
    '   celdas = celdas & ofv.GenCelda("","tt","center","","10%","","",sInput,"si")



  '  filas = filas & ofv.GenRow("","","","","","","",celdas,"si")



    posi = posi + 1

  END IF

    celdas = ""
    filas  = ""


  'Encabezados


  celdas = ofv.GenCelda("","VT ","","20%","","","","DOCUMENTO","si")
  celdas = celdas & ofv.GenCelda("","uti ","","80%","","","",sInputn1,"si")
  filas  = filas & ofv.GenRow("","","","","","","",celdas,"si")

  celdas = ofv.GenCelda("","VT  ","","","","","","CIRCUITO","si")
  celdas = celdas & ofv.GenCelda("","uti  ","","","","","",sInputn2,"si")
  filas  = filas & ofv.GenRow("","","","","","","",celdas,"si")

  celdas = ofv.GenCelda("","VT  ","","","","","","INICIANTE","si")
  celdas = celdas & ofv.GenCelda("","uti  ","","","","","",sInputn3,"si")
  filas  = filas & ofv.GenRow("","","","","","","",celdas,"si")


  sHTML = sHTML  & ofv.GenTabla("","AAA  ","","98%","","0","","0","0","","",filas,"si")

'''' PASOS

    sImagenO  = "<B><FONT FACE=Corflaticon COLOR=#6699cc SIZE=3>&#0128;</FONT></B>"
    sImagenI  = "<B><FONT FACE=Corflaticon COLOR=#AA0000 SIZE=3>&#0128;</FONT></B>"
    sImagenE  = "<B><FONT FACE=corflaticon COLOR=#f7aa22 SIZE=3>&#0240;</FONT></B>"
    sImagenC  = "<B><FONT FACE=Corflaticon COLOR=#008000 SIZE=3>R</FONT></B>"
	
    'sImagenP  = "<B><FONT FACE=Corflaticon COLOR=#336699 SIZE=3>2</FONT></B>"
	sImagenP  = "<B><i class='' COLOR=#336699 ></i></B>"
	
    sImagenRU = "<B><FONT FACE=Corflaticon COLOR=#f7aa22 SIZE=3>&#0128;</FONT></B>"
    sImagenRG = "<B><FONT FACE=corflaticon COLOR=#00aa00 SIZE=3>&#0148;</FONT></B>"
    sImagenRS = "<B><FONT FACE=Corflaticon COLOR=#0000aa SIZE=3>G</FONT></B>"
    sImagenFI = "<B><FONT FACE=Corflaticon COLOR=#6699cc SIZE=3>5</FONT></B>"
    sImagenFF = "<B><FONT FACE=Corflaticon COLOR=#993333 SIZE=3>6</FONT></B>"

    sImagenA  = "<FONT FACE=Corflaticon COLOR=#005500 SIZE=3>&#097;</FONT>"
    sImagenR  = "<FONT FACE=Corflaticon COLOR=#AA0000 SIZE=4>&#0114;</FONT>"
    sImagenK  = "<FONT FACE=Corflaticon COLOR=#F7AA22 SIZE=3>&#0161;</FONT>"
    sImagenIN = "<FONT FACE=Corflaticon COLOR=#336699 SIZE=3>&#0236;</FONT>"


  'Encabezados
  celdas =  ofv.GenCelda("","VT colspan=4 STYLE='border-TOP-style:solid;border-TOP-width:1;border-TOP-color:" & Session("CMFPNTIT") & ";' ","center","","","","","&nbsp;","si")
  filas  =  ofv.GenRow("","","","","","","",celdas,"si")

  PASOANT = " "
  IPASO   = 0

  StrSql = "Select SEGPASO.*,SEGCKTRESP.codempre  From SEGPASO "
  StrSql =StrSql  & " left join SEGCKTRESP ON SEGPASO.SEGCKT = SEGCKTRESP.SEGCKT AND SEGCKTRESP.ID = SEGPASO.SEGIDRESP "
  StrSql =StrSql  & "  WHERE SEGPASO.SEGCKT ='" & SESSION("CKT") & "' "
  StrSql =StrSql  & " ORDER BY segpaso.ID "
  'response.write strsql
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then
  posi = 0
  Do While (Not rs.EOF)

    sform = "WKF" & rs(0)


       sc = "ut"



    ''   celdas = ofv.GenCelda("","","","","","","","&nbsp;","si")


       IF LEFT(RS(10),INSTR(RS(10),":")) <> PASOANT THEN

    celdas = ofv.GenCelda("","aaa colspan=4 style='border-top-style:solid;border-top-width:2;border-top-color:#dddddd;' ","","","","","","&nbsp;","si")
   ' filas = filas & ofv.GenRow("","","","","","","",celdas,"si")

          sInput = "<b>" & sImagenp & "</b> " & rs(10)
          celdas = ofv.GenCelda("",sc & "  COLSpan=4 BGCOLOR=#f9f9f9 style='border-top-style:solid;border-top-width:2;border-top-color:#dddddd;' ","","","","","",sInput,"si")
           if rs(8) & "" = "NO" AND WKAPLIC = "OR" THEN
           ELSE
          filas = filas & ofv.GenRow("","","","","30","","",celdas,"si")
          END IF
          PASOANT = LEFT(RS(10),INSTR(RS(10),":"))
          IPASO   = 0
       END IF

       IPASO = IPASO + 1

       SINPUT = "." & IPASO

       xlink = "<A CLASS=CC href='&nbsp;'  "
       xlink = xlink & " onclick=" & chr(34) & "AbrirCons('SEGCKTESTPSC.asp?PSC=" & rs(2) & "'); return false;" & chr(34)
       xlink = xlink & " >" & SINPUT & "</A>"

       SINPUT = xlink


       celdas = ofv.GenCelda("","VT ROWSPAN=2 STYLE='border-TOP-style:solid;border-TOP-width:1;border-TOP-color:#6699CC;' ","RIGHT","3%","","","",SINPUT,"si")

       sInput = sImagenIN

       IF RS(5) = "NO" THEN
          sInput = sImagenK
       ELSE

	  Strsql = "SELECT * FROM SEGPASO WHERE SEGCKTPSC = '" & rs(6) & "' "
	  set rsxT = ofv.crearconsultaEx(StrSql,cn,1,parametros)
	  if NOT rsxT.eof then

             PSCMODEL = rsxT(3)
             PSCMOT   = "A"

          END IF

	  Strsql = "SELECT RCHPSC FROM WKFPASOS WHERE WKFPASCOD = '" & rs(3) & "' "
	  set rsxu = ofv.crearconsultaEx(StrSql,cn,1,parametros)
	  if NOT rsxU.eof then

             IF RSXU(0) = PSCMODEL THEN PSCMOT   = "R"

          END IF

          SELECT CASE PSCMOT
                 CASE "A"
                      sInput = sImagenA
                 CASE "R"
                      sInput = sImagenR
          END SELECT


       END IF

       xlink = "<A CLASS=NT href='&nbsp;'  "
       xlink = xlink & " onclick=" & chr(34) & "AbrirCons('SEGCKTESTPSC.asp?PSC=" & rs(2) & "'); return false;" & chr(34)
       xlink = xlink & " >" & SINPUT & "</A>"

       SINPUT = xlink

       celdas = celdas & ofv.GenCelda("",sc & " ROWSPAN=2 style='border-RIGHT-style:solid;border-RIGHT-width:1;border-RIGHT-color:#6699CC;border-TOP-style:solid;border-TOP-width:1;border-TOP-color:#6699CC;' ","","4%","","","",SINPUT,"si")


       sInput = GETRESPO(RS(7),RS(8))
       celdas = celdas & ofv.GenCelda("",sc,"","45%","","","",sInput,"si")

       sInput = "SIN ASIGNAR"
       sIFunc = ""
       IF RS(9) <> "NO" THEN

          set cn2 = ofv.conectar(ofv.strconn0)

       '   StrSql = "Select * From USUARIOS WHERE ID = '" & RS(9) & "' "
       '   set rs2 = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
        '  if not rs2.eof THEN
         '     sInput = rs2(1)
        '      sIFunc = GETUSUARIO(RS(9))
        '  END IF

        SELECT CASE session("APLMDCLI")


		 CASE "GP"

		   IF rs("codempre") & "" <>  "" THEN
		   IdDomE= rs("codempre")
		   else
		   IdDomE=  SESSION("UN")
		   end if

		   xxxx = getInfoUsereX(RS(9),IdDomE)
			CASE ELSE
		     xxxx = getInfoUser(RS(9))

           END SELECT

      '  xxxx = getInfoUser(RS(9))

	  sInput =  RS(9) & " - " & session("SGU_nombre")


        sIFunc = GETUSUARIO(RS(9))


       END IF

       sInput = "<b>" & sImageni & "</b> " & sInput
       celdas = celdas & ofv.GenCelda("",sc & " " & sIFunc,"","45%","","","",sInput,"si")
       filas = filas & ofv.GenRow("","","","","","","",celdas,"si")

       sInput = "<b>" & sImagenFI & "</b> " & rs(4)
       celdas = ofv.GenCelda("",sc,"","45%","","","",sInput,"si")

       sInput = "<b>" & sImagenFF & "</b> "

       IF RS(5) = "NO" THEN
          sInput = sInput & "<B><FONT COLOR=#F7AA66 SIZE=1>PENDIENTE</FONT></B>"
       ELSE
          sInput = sInput & rs(5)
       END IF

       celdas = celdas & ofv.GenCelda("",sc,"","45%","","","",sInput,"si")

    filas = filas & ofv.GenRow("","","","","","","",celdas,"si")





    posi = posi + 1
    rs.Movenext
  loop
  END IF


    celdas = ofv.GenCelda("","aaa colspan=4 style='border-top-style:solid;border-top-width:2;border-top-color:#dddddd;' ","","","","","","&nbsp;","si")
    filas = filas & ofv.GenRow("","","","","","","",celdas,"si")

  sHTML = sHTML  & ofv.GenTabla("","aaa  ","","100%","","0","","-2","-1","","",filas,"si")

'''' FIN PASOS
       sinput = "&nbsp;"
       celdas = ""

       celdas = celdas & ofv.GenCelda("","VT","center","","1%","","","&nbsp;" & sImagenA,"si")
       celdas = celdas & ofv.GenCelda("","VT","center","","1%","","","Aprueba","si")

       celdas = celdas & ofv.GenCelda("","VT","center","","1%","","","&nbsp;" & sImagenR,"si")
       celdas = celdas & ofv.GenCelda("","VT","center","","1%","","","Rechaza","si")

       celdas = celdas & ofv.GenCelda("","VT","center","","1%","","","&nbsp;" & sImagenK,"si")
       celdas = celdas & ofv.GenCelda("","VT","center","","1%","","","Pendiente","si")
       filas =  ofv.GenRow("","","","","","","",celdas,"si")

       sHTML = sHTML  & ofv.GenTabla("","aaa ","","30%","","0","","0","0","","",filas,"si")


    sHTML = sHTML  & Ghtml


   Agregar = ""

  'Boton test

 if wkaplic <> "OR" then

  xxxx    = CorGenBotonXp(5, "NO", "NO", "NO", "NO", "Ver grafico del circuito.", "NO", "Grafico")

  sAccion = "WKFVER.asp?CKT=" & segCKT
  Buscar  = ofv.BotonConsultar(sAccion)

 else
  buscar = ""
 end if

  if  DSD <> "REQ" then
      sAccion = "COHISTWKF.asp"
  else
      sAccion = "EDHISTWKF.asp"
  end if

  Celdav = ofv.BotonVolver(sAccion)





'--------------------------------------------------------------------------------
  sHTML = sHTML & "</div>"

  Celda = ofv.GenCelda("","vt","","60%","","","","&nbsp;&nbsp;","si")
  Fila = ofv.GenRow("","","","","","","",Celda,"si")
  sHTML = sHTML &  ofv.GenTabla("TCEX1","onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")

  'Estructura del Indice

  sHTML = sHTML & CorPnlIndx("N","")

'--------------------------------------------------------------------------------

sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","",Agregar,Buscar,Celdav)

  Whtml = ""
  Whtml = Whtml & "<script language=javascript>" & chr(13)
''  Whtml = Whtml & " var Fbtnas = 0; "
  Whtml = Whtml & " MoveBtnas(); " & chr(13)
  Whtml = Whtml & "</script>" & chr(13)

  sHTML = sHTML & Whtml

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)


  sHTML = sHTML  & "</body></html>"
end if

response.write sHTML


FUNCTION GETRESPO(RESPTIP,RESPCOD)

  set cn1 = ofv.conectar(ofv.strconn5)
  set cn2 = ofv.conectar(ofv.strconn0)
  set cn3 = ofv.conectar(ofv.strconn2)

  sInput = "SIN ASIGNAR"



    XRESPO = RESPCOD & ""

    if XRESPO <> "" then

       SELECT CASE RESPTIP

         CASE "USR"
           RESPTIP = "<b>" & sImagenRU & "</b> "
         '  strsql = "Select * From USUARIOS where ID = '" & XRESPO & "' "
         '  set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
          ' if not rsniv.eof then sInput = rsniv(1)
          ' call ofv.cerrarconsulta(rsniv)
           xxxx = getInfoUserex(XRESPO,rs("codempre"))

	      sInput =  XRESPO & " - " & session("SGU_nombre")



         CASE "GRP"
           RESPTIP = "<b>" & sImagenRG & "</b> "
           strsql = "Select * From GRUPOS where ID = '" & XRESPO & "' "
           set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
           if not rsniv.eof then sInput = rsniv(1)
           call ofv.cerrarconsulta(rsniv)

         CASE "SEC"
           RESPTIP = "<b>" & sImagenRS & "</b> "
           strsql = "Select * From ECO where ECO = " & XRESPO
           set rsniv = ofv.crearconsultaEx(StrSql,cn3,1,parametros)
           if not rsniv.eof then sInput = rsniv(1)
           call ofv.cerrarconsulta(rsniv)

       END SELECT
    ELSE
       sInput =  XRESPO & " - SIN DATOS"

    END IF




  call ofv.cerrarconn(cn1)
  call ofv.cerrarconn(cn2)
  call ofv.cerrarconn(cn3)

  GETRESPO = RESPTIP & " " & sInput

END FUNCTION

function EsqnomP(xtipo,xcodigo,ofv)

 set xcn=ofv.conectar(ofv.strconn2)


 IF xtipo = "RSK" THEN

    strsql = "SELECT * FROM  ORRSKCIRCUITO "
    strsql = strsql & " Where OBJID = 0" & xcodigo
    set xrs = ofv.crearconsultaEx(StrSql,xcn,1,parametros)
    if not xrs.eof then
       IF XRS("APLIC") = "OR" THEN
          xdescri = "CIRCUITO AUTOEVALUACION"
       ELSE
          xdescri = "PLAN DE CONTINGENCIAS"
       END IF

       Strsql = "Select * From objetos where codobjeto = 0" & XRS("CODOBJETO")
       set xrs = ofv.crearconsultaEx(StrSql,xcn,1,parametros)
       if not xrs.eof then
          xiden = xrs(1)
          xobjdes = xrs(3)
       end if


    END IF
    call ofv.cerrarconsulta(xrs)

 ELSE




    Strsql="Select descri From tipobjeto where tipobjeto ='" & xtipo & "' "
    set xrs=ofv.crearconsultaEx(StrSql,xcn,1,parametros)
    if not xrs.eof then xdescri = xrs(0)
    call ofv.cerrarconsulta(xrs)

    Strsql="Select * From objetos where codobjeto =" & xcodigo
    set xrs=ofv.crearconsultaEx(StrSql,xcn,1,parametros)
    if not xrs.eof then
       xiden = xrs(1)
       xobjdes = xrs(3)
    end if
    call ofv.cerrarconsulta(xrs)

 END IF





 call ofv.cerrarconn(xcn)

 esqnomP = xdescri & ":&nbsp;" & xiden & "&nbsp;-&nbsp;" & xobjdes & ""

end function



FUNCTION GETUSUARIO(XFUSER)

    SUSROK = "N"

    IF  XFUSER <> "" THEN
        SUSROK = "Y"
        For indT = 1 to TUTabla
            if AUTabla (indT,0) = XFUSER then
               SUSROK = "S"
               exit for
            end if
        next
    end if

    if  SUSROK = "Y" then

    GTUcase = session("USRPRV")
    SELECT CASE  GTUcase

     CASE "I"
          set xcn=ofv.conectar(ofv.strconn2)

          StrSql  = "Select * From USUARIOS WHERE ID = '" & XFUSER & "' "
         ''' response.write strsql & "<br>"
          set rs2 = ofv.crearconsultaEx(StrSql,xcn,1,parametros)
          if not rs2.eof THEN


           SUSROK = "S"
           IUTabla             = IUTabla + 1
           AUTabla (IUTabla,0) = XFUSER
           TUTabla             = IUTabla
           indT                = IUTabla

           XFUNOM = rs2(1)
           XFUEMP = rs2(29)
           XFUSTE = rs2(31)
           XFUSEC = rs2(21)
           XFUFUN = rs2(27)
           XMAIL = UCASE(rs2("EMAIL")&"")

           strsql = "Select * From tipoperfil where tipoperfil = '" & XFUFUN & "' "
           set rsniv = ofv.crearconsultaEx(StrSql,xcn,1,parametros)
           if  not rsniv.eof then XFUFUN = rsniv(1)
           call ofv.cerrarconsulta(rsniv)

           strsql = "Select * From clienteplan where id = 0" & XFUEMP
           set rsniv = ofv.crearconsultaEx(StrSql,xcn,1,parametros)
           if  not rsniv.eof then XFUEMP = rsniv(1)
           call ofv.cerrarconsulta(rsniv)

           sGrupo = "Sin SITE/SUCURSAL ASIGNADA"
           Strsql = "Select descri From SITES Where OBJID = 0" & XFUSTE
           set rsGrupo = ofv.crearconsultaEx(StrSql,xcn,1,parametros)
           if  not rsGrupo.eof then XFUSTE = rsGrupo(0)
           call ofv.cerrarconsulta(rsGrupo)

           sGrupo = "Sin sector asignado"
           Strsql = "Select descri From eco Where eco = 0" & XFUSEC
           set rsGrupo = ofv.crearconsultaEx(StrSql,xcn,1,parametros)
           if  not rsGrupo.eof then XFUSEC = rsGrupo(0)
           call ofv.cerrarconsulta(rsGrupo)


       Ghtml = Ghtml & "<DIV id='PDUSR" & REPLACE(XFUSER," ","_") & "' STYLE='display:none;z-index:99;margin-top:1;margin-left:1;position:absolute;left:100px;top:100px;border-style:solid;border-width:2;border-color:" & Session("CMFPNTIT") & ";background-color:" & session("APLBGD") & ";"
       Ghtml = Ghtml & "'>"
       Ghtml = Ghtml & "<table id='MTUSR" & REPLACE(XFUSER," ","_") & "'  class=aaa width=400 cellspacing=4 cellpadding=2 "
       Ghtml = Ghtml & " style='border-style:solid;border-width:3;border-color:#dddddd;background-color:#f9f9f9;' "
       Ghtml = Ghtml & " >"
       Ghtml = Ghtml & "<tr>"
       Ghtml = Ghtml & "<td class=ttt width=70% colspan=2 >&nbsp;&nbsp;<b>" & sImageni & "</b>&nbsp;&nbsp;" & XFUNOM
       Ghtml = Ghtml & "</td>"
       Ghtml = Ghtml & "<td class=btbs width=5% onclick=""document.getElementById('PDUSR" & REPLACE(XFUSER," ","_") & "').style.display='none';"" style=""border-style:solid;border-width:1;border-color:#999999;"" >"
       Ghtml = Ghtml & "<font class=btic >r</font>"
       Ghtml = Ghtml & "</td>"
       Ghtml = Ghtml & "</tr>"

       Ghtml = Ghtml & "<tr>"
       Ghtml = Ghtml & "<td class=tt   >&nbsp;Empresa</td>"
       Ghtml = Ghtml & "<td class=vti  colspan=2  >"
       Ghtml = Ghtml & "<font color=#999999 >" & XFUEMP & "</font>"
       Ghtml = Ghtml & "</td>"
       Ghtml = Ghtml & "</tr>"
       Ghtml = Ghtml & "<tr>"
       Ghtml = Ghtml & "<td class=tt   >&nbsp;SITE/SUCURSAL</td>"
       Ghtml = Ghtml & "<td class=vti  colspan=2  >"
       Ghtml = Ghtml & "<font color=#999999 >" & XFUSTE & "</font>"
       Ghtml = Ghtml & "</td>"
       Ghtml = Ghtml & "</tr>"
       Ghtml = Ghtml & "<tr>"
       Ghtml = Ghtml & "<td class=tt   >&nbsp;SECTOR</td>"
       Ghtml = Ghtml & "<td class=vti  colspan=2  >"
       Ghtml = Ghtml & "<font color=#999999 >" & XFUSEC & "</font>"
       Ghtml = Ghtml & "</td>"
       Ghtml = Ghtml & "</tr>"
       Ghtml = Ghtml & "<tr>"
       Ghtml = Ghtml & "<td class=tt   >&nbsp;PUESTO</td>"
       Ghtml = Ghtml & "<td class=vti  colspan=2  >"
       Ghtml = Ghtml & "<font color=#999999 >" & XFUFUN & "</font>"
       Ghtml = Ghtml & "</td>"
       Ghtml = Ghtml & "</tr>"
       Ghtml = Ghtml & "<tr>"
       Ghtml = Ghtml & "<td class=tt   >&nbsp;MAIL</td>"
       Ghtml = Ghtml & "<td class=vti  colspan=2  >"
       Ghtml = Ghtml & "<font color=#999999 >" & XMAIL & "</font>"
       Ghtml = Ghtml & "</td>"
       Ghtml = Ghtml & "</tr>"

       Ghtml = Ghtml & "<tr>"
       Ghtml = Ghtml & "<td class=ut width=25%  >&nbsp;</td>"
       Ghtml = Ghtml & "<td class=ut width=70%  >&nbsp;</td>"
       Ghtml = Ghtml & "<td class=ut width=5%   >&nbsp;</td>"
       Ghtml = Ghtml & "</tr>"
       Ghtml = Ghtml & "</table>"
       Ghtml = Ghtml & "</DIV>"

          END IF
    CASE ELSE

             XXX=GETINFOUSER(XFUSER)

              set xcn=ofv.conectar(ofv.strconn2)
              IF session("SGU_nombre") <> "NO" THEN


                  SUSROK = "S"
                  IUTabla             = IUTabla + 1
                  AUTabla (IUTabla,0) = XFUSER
                  TUTabla             = IUTabla
                  indT                = IUTabla

                  XFUNOM = session("SGU_nombre")
                  XFUEMP = ""
                  XFUSTE = ""
                  XFUSEC =session("SGU_idcostcenter")

                  IF CSTR(XFUSEC&"") ="0" THEN XFUSEC=""
                  XFUFUN = session("SGU_codigopuesto")
                  XMAIL = UCASE(session("SGU_mail")&"")

                  strsql = "Select * From tipoperfil where tipoperfil = '" & XFUFUN & "' "
                  'RESPONSE.WRITE STRSQL
                  set rsniv = ofv.crearconsultaEx(StrSql,xcn,1,parametros)
                  if  not rsniv.eof then XFUFUN = rsniv(1)
                  call ofv.cerrarconsulta(rsniv)

                  strsql = "Select * From clienteplan where id = 0" & XFUEMP

                  set rsniv = ofv.crearconsultaEx(StrSql,xcn,1,parametros)
                  if  not rsniv.eof then XFUEMP = rsniv(1)
                  call ofv.cerrarconsulta(rsniv)

                  sGrupo = "Sin SITE/SUCURSAL ASIGNADA"
                  Strsql = "Select descri From SITES Where OBJID = 0" & XFUSTE
                  set rsGrupo = ofv.crearconsultaEx(StrSql,xcn,1,parametros)
                  if  not rsGrupo.eof then XFUSTE = rsGrupo(0)
                  call ofv.cerrarconsulta(rsGrupo)

                  sGrupo = "Sin sector asignado"
                  Strsql = "Select descri From eco Where eco = 0" & XFUSEC
                  set rsGrupo = ofv.crearconsultaEx(StrSql,xcn,1,parametros)
                  if  not rsGrupo.eof then XFUSEC = rsGrupo(0)
                  call ofv.cerrarconsulta(rsGrupo)


              Ghtml = Ghtml & "<DIV id='PDUSR" & REPLACE(XFUSER," ","_") & "' STYLE='display:none;z-index:99;margin-top:1;margin-left:1;position:absolute;left:100px;top:100px;border-style:solid;border-width:2;border-color:" & Session("CMFPNTIT") & ";background-color:" & session("APLBGD") & ";"
              Ghtml = Ghtml & "'>"
              Ghtml = Ghtml & "<table id='MTUSR" & REPLACE(XFUSER," ","_") & "'  class=aaa width=400 cellspacing=4 cellpadding=2 "
              Ghtml = Ghtml & " style='border-style:solid;border-width:3;border-color:#dddddd;background-color:#f9f9f9;' "
              Ghtml = Ghtml & " >"
              Ghtml = Ghtml & "<tr>"
              Ghtml = Ghtml & "<td class=ttt width=70% colspan=2 >&nbsp;&nbsp;<b>" & sImageni & "</b>&nbsp;&nbsp;" & XFUNOM
              Ghtml = Ghtml & "</td>"
              Ghtml = Ghtml & "<td class=btbs width=5% onclick=""document.getElementById('PDUSR" & REPLACE(XFUSER," ","_") & "').style.display='none';"" style=""border-style:solid;border-width:1;border-color:#999999;"" >"
              Ghtml = Ghtml & "<font class=btic >r</font>"
              Ghtml = Ghtml & "</td>"
              Ghtml = Ghtml & "</tr>"


              Ghtml = Ghtml & "<td class=tt   >&nbsp;SECTOR</td>"
              Ghtml = Ghtml & "<td class=vti  colspan=2  >"
              Ghtml = Ghtml & "<font color=#999999 >" & XFUSEC & "</font>"
              Ghtml = Ghtml & "</td>"
              Ghtml = Ghtml & "</tr>"
              Ghtml = Ghtml & "<tr>"
              Ghtml = Ghtml & "<td class=tt   >&nbsp;PUESTO</td>"
              Ghtml = Ghtml & "<td class=vti  colspan=2  >"
              Ghtml = Ghtml & "<font color=#999999 >" & XFUFUN & "</font>"
              Ghtml = Ghtml & "</td>"
              Ghtml = Ghtml & "</tr>"
              Ghtml = Ghtml & "<tr>"
              Ghtml = Ghtml & "<td class=tt   >&nbsp;MAIL</td>"
              Ghtml = Ghtml & "<td class=vti  colspan=2  >"
              Ghtml = Ghtml & "<font color=#999999 >" & XMAIL & "</font>"
              Ghtml = Ghtml & "</td>"
              Ghtml = Ghtml & "</tr>"
              Ghtml = Ghtml & "<tr>"
	      Ghtml = Ghtml & "<td class=tt   >&nbsp;GRUPOS</td>"
	      Ghtml = Ghtml & "<td class=vti  colspan=2  >"
	      Ghtml = Ghtml & "<font color=#999999 >" & GRUPOSUSRT(XFUSER) & "</font>"
	      Ghtml = Ghtml & "</td>"
              Ghtml = Ghtml & "</tr>"

              Ghtml = Ghtml & "<tr>"
              Ghtml = Ghtml & "<td class=ut width=25%  >&nbsp;</td>"
              Ghtml = Ghtml & "<td class=ut width=70%  >&nbsp;</td>"
              Ghtml = Ghtml & "<td class=ut width=5%   >&nbsp;</td>"
              Ghtml = Ghtml & "</tr>"
              Ghtml = Ghtml & "</table>"
              Ghtml = Ghtml & "</DIV>"
          END IF

   END SELECT
    END IF

    if  SUSROK = "S" then

        XFFunc = " style='cursor:pointer;' onclick=""document.getElementById('PDUSR" & REPLACE(XFUSER," ","_") & "').style.display='block';"" "

    else
        XFFunc = ""

    end if

    GETUSUARIO = XFFunc

END FUNCTION


FUNCTION GRUPOSUSRT(Usuario)
 call getGroupsUser(Usuario)
 session("SGU_grupos")="DEF" &CHR(9) & session("SGU_grupos")
 GRUPOSUSRT= ChkGrp(session("SGU_grupos"))

END FUNCTION

function ChkGrp(sGRUPOS)

'    SESSION("USREXPIR") = 50

  SESSION("USREXPIR") = "0"

  'set XFCN = ofv.conectar(ofv.strconn0)

  XFStrSql = "Select id,sessexp,descripcion From Grupos where "

  XMFTGRP = split(sGRUPOS,chr(9))
  XMFSEP  = ""

  if session("USRPRV")  = "I" or session("USRPRV")  = "E"  then
  XFStrSql = XFStrSql & " id in ("
  else
  XFStrSql = XFStrSql & " mapgrp in ("
  end if

  for XMFI = 0 to ubound(XMFTGRP)

      XFStrSql  = XFStrSql  & XMFSEP & "'" & replace(XMFTGRP(XMFI) & "","'","''") & "'"
      XMFSEP  = ","
  next

  XFStrSql = XFStrSql & ") "
  XFStrSql = XFStrSql & " Order By id"
' response.write XFStrSql
'XFStrSql= "Select id,sessexp,descripcion From Grupos where  mapgrp in ('DEF','DEF','GADVapComunidadGestion','GAPFtprootInfdeu','GAPSharepointClaridad','GapInternetDBG','GAPSAMA','UPAECOMGciaServCorporativos','GGdwarehouse','UPAECOMREMEDY6900','UPAECOMSistDesaAnalista','GAPSpointROperPilotoCuanti','Soft','UPAECOMGciaSistemas','GAPClaridadRedactor','GAPVapEmpleadosTODOS','GAPWebSCISistEfectivos','GGDWuserfinal','UPMECOMUsrEcompany','GADVapGrupoSIDAG','GAPPlanySegCambio','POLNoShared','GAPVCMEcomComSucRdJ','GapVCMEcom2RedJefe','GAPWSIAreaCentral','GAPArchivoGarant','UPAECOMAreasCentralesSinCCC','GAPDocumentacionLector','GAPWSISistemas','GAPNop','UPAECOMHombres','GGAplicConsultaPou','GAPClaridadAdministrador','UPACMDB','GPS1241AQ','GAPVAPPilotoRHOL','Usuarios de Sistemas','UPAECOMEmpleados','POLSi95','UPAECOMAreasCentrales','UPAECOMPersonalTodos','GAPClaridadEditor','POLSi95Sistemas','UPAECOMSistDesarrollo','GAPNormaBasico','GAPVapEmpleadosTODOSBIS','')  Order By id "
  set XFrs = ofv.crearconsultaEx(XFStrSql,cn2,1,parametros)

  if  not XFrs.eof then

      do until XFRS.EOF


          ChkGrp = ChkGrp & XFRS(0) &": " &  XFRS(2) & "<BR>"

         XMFSEP  = chr(9)
      XFRS.MOVENEXT
      LOOP

      IF  session("SegGrpLst") = "" THEN session("SegGrpLst") = Session("SegGrp")

  else
    '  session("SegGrpLst") = Session("SegGrp")
     ' session("SegGrpdsc") = Session("SegGrp")
  end if
  call ofv.cerrarconsulta(XFrs)
 ' call ofv.cerrarconn(XFCN)



end function

%>