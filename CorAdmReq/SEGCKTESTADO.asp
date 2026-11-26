<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/CorWKFNro.inc" -->
<%

ok=ofv.CheckUsuario()

if ok <> "Y" then
  sHTML = ok
else
  response.expires=0




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


  dim AUTabla (100,2)
  TUTabla = 0
  IUTabla = 0


  sHTML = sHTML  & "<script language=javascript> "
  sHTML = sHTML  & "function AbrirCons(Datos) { "
  sHTML = sHTML  & " myWin= open(Datos," & chr(34) & "ConsP" & chr(34) & "," & chr(34) & "RESIZABLE=YES,menubar=no,toolbar=no,height=570,width=790" & chr(34) & " ); "
  sHTML = sHTML  & " } "
  sHTML = sHTML  & "</script>"


       Ghtml = ""




  StrSql = "Select * From SEGCIRCUITO WHERE SEGCKT = '" & SESSION("CKT") & "' "
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
          sInput = esqnom (WKFOBJTIPO,WKFOBJCODE,ofv)

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

    sImagenO  = "<B><FONT FACE=WEBDINGS COLOR=#6699cc SIZE=3>&#0128;</FONT></B>"
    sImagenI  = "<B><FONT FACE=WEBDINGS COLOR=#AA0000 SIZE=3>&#0128;</FONT></B>"
    sImagenE  = "<B><FONT FACE=WINGDINGS COLOR=#f7aa22 SIZE=3>&#0240;</FONT></B>"
    sImagenC  = "<B><FONT FACE=WINGDINGS COLOR=#008000 SIZE=3>R</FONT></B>"
    sImagenP  = "<B><FONT FACE=WebDINGS COLOR=#336699 SIZE=3>2</FONT></B>"
    sImagenRU = "<B><FONT FACE=WEBDINGS COLOR=#f7aa22 SIZE=3>&#0128;</FONT></B>"
    sImagenRG = "<B><FONT FACE=WEBDINGS COLOR=#00aa00 SIZE=3>&#0148;</FONT></B>"
    sImagenRS = "<B><FONT FACE=WEBDINGS COLOR=#0000aa SIZE=3>G</FONT></B>"
    sImagenFI = "<B><FONT FACE=WEBDINGS COLOR=#6699cc SIZE=3>5</FONT></B>"
    sImagenFF = "<B><FONT FACE=WEBDINGS COLOR=#993333 SIZE=3>6</FONT></B>"

    sImagenA  = "<FONT FACE=WINGDINGS COLOR=#005500 SIZE=3>&#0252;</FONT>"
    sImagenR  = "<FONT FACE=WINGDINGS COLOR=#AA0000 SIZE=4>&#0251;</FONT>"
    sImagenK  = "<FONT FACE=WINGDINGS COLOR=#F7AA22 SIZE=3>&#0184;</FONT>"
    sImagenIN = "<FONT FACE=WINGDINGS COLOR=#336699 SIZE=3>&#0236;</FONT>"


  'Encabezados
  celdas =  ofv.GenCelda("","VT colspan=4 STYLE='border-TOP-style:solid;border-TOP-width:1;border-TOP-color:#f7aa22;' ","center","","","","","&nbsp;","si")
  filas  =  ofv.GenRow("","","","","","","",celdas,"si")

  PASOANT = " "
  IPASO   = 0

  StrSql = "Select * From SEGPASO WHERE SEGCKT ='" & SESSION("CKT") & "' ORDER BY ID "
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
          filas = filas & ofv.GenRow("","","","","30","","",celdas,"si")
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
        
        xxxx = getInfoUser(RS(9))
	if session("SGU_nombre")<> "" then    sInput =  session("SGU_nombre")

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

  sHTML = sHTML  & ofv.GenTabla("","aaa  ","","100%","","0","","-2","-1","","",filas,"si") & "<BR>"

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

  xxxx    = CorGenBotonXp(5, "NO", "NO", "NO", "NO", "Ver grafico del circuito.", "NO", "Grafico")

  sAccion = "WKFVER.asp?CKT=" & segCKT
  Buscar  = ofv.BotonConsultar(sAccion)

  if  DSD <> "REQ" then
      sAccion = "COHISTWKF.asp"
  else
      sAccion = "EDHISTWKF.asp"
  end if

  Celda = ofv.BotonVolver(sAccion)

  
  sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","",Agregar,Buscar,Celda)

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
           xxxx = getInfoUser(XRESPO)
	   if session("SGU_nombre")<> "" then    sInput =  session("SGU_nombre")


               
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

    END IF




  call ofv.cerrarconn(cn1)
  call ofv.cerrarconn(cn2)
  call ofv.cerrarconn(cn3)

  GETRESPO = RESPTIP & " " & sInput

END FUNCTION

function Esqnom(xtipo,xcodigo,ofv)

set xcn=ofv.conectar(ofv.strconn2)

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

call ofv.cerrarconn(xcn)

esqnom = ""
esqnom = esqnom & xdescri & ":&nbsp;" & xiden & "&nbsp;-&nbsp;" & xobjdes & ""

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
       Ghtml = Ghtml & "<td class=btbs width=5% onclick=""document.all.PDUSR" & REPLACE(XFUSER," ","_") & ".style.display='none';"" style=""border-style:solid;border-width:1;border-color:#999999;"" >"
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
       Ghtml = Ghtml & "<td class=ut width=25%  >&nbsp;</td>"
       Ghtml = Ghtml & "<td class=ut width=70%  >&nbsp;</td>"
       Ghtml = Ghtml & "<td class=ut width=5%   >&nbsp;</td>"
       Ghtml = Ghtml & "</tr>"
       Ghtml = Ghtml & "</table>"
       Ghtml = Ghtml & "</DIV>"

          END IF

    END IF

    if  SUSROK = "S" then

        XFFunc = " style='cursor:hand;' onclick=""document.all.PDUSR" & REPLACE(XFUSER," ","_") & ".style.display='block';"" "

    else
        XFFunc = ""

    end if

    GETUSUARIO = XFFunc

END FUNCTION


%>