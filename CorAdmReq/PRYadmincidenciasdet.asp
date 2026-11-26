<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/CorWKFNro.inc" -->
<%

ok=ofv.CheckUsuario()

if ok <> "Y" then
  sHTML = ok
else
  response.expires=0
  session("ModMain")="ED"
  session("NCliente") = ""
  SESSION("VPLAN")    = 0

  sHTML = ofv.formHeader("")




  sHTML = sHTML  & "<script language=javascript> "
''  sHTML = sHTML  & "        document.getElementById('MDPTOP1').style.display='none';  document.getElementById('tbhMAX').style.display='none';        hdemenu();     "
  sHTML = sHTML  & "function AbrirCons(Datos) { "
  sHTML = sHTML  & " myWin= open(Datos," & chr(34) & "Cons" & chr(34) & "," & chr(34) & "status=yes,resizable=yes,menubar=no,toolbar=no,height=570,width=790" & chr(34) & " ); "
  sHTML = sHTML  & " } "
  sHTML = sHTML  & "</script>"

  sHTML = sHTML  & GENINCIDENCIASHED()



  CantRegMostrarx = 6
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))
  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)

  set cn = ofv.conectar(ofv.strconn6)
  set cn2 = ofv.conectar(ofv.strconn0)
  set cn3 = ofv.conectar(ofv.strconn2)

  ofv.ObtenerAtributos session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  sopcion = request.querystring("opcion")
  sopc    = request.querystring("opc")
  sTIPO   = request.querystring("TIPO")
  sRESP   = request.querystring("RESP")
  sSEL    = request.querystring("SEL")
  XOPT    = request.querystring("XOPT")
  XRESP   = request.querystring("RESP")

  XESTAD   = request.querystring("XESTAD")

''  response.write "tipo: " & stipo


     SESSION("STIPO") = "PRY"


  IF XESTAD <> "" THEN
     SESSION("XESTAD") = XESTAD
  ELSE
     XESTAD = SESSION("XESTAD")
  END IF

  IF  SESSION("XESTAD") = "" THEN
      SESSION("XESTAD") = 12
      XESTAD            = 12
  END IF


  IF sRESP <> "" THEN
     SESSION("sRESP") = sRESP
  ELSE
     sRESP = SESSION("sRESP")
  END IF

  IF XOPT <> "" THEN
     SESSION("XOPT") = XOPT
  ELSE
     XOPT = SESSION("XOPT")
  END IF

  IF XRESP <> "" THEN
     SESSION("XRESP") = XRESP
  ELSE
     XRESP = SESSION("XRESP")
  END IF

  IF sopc <> "" THEN
     SESSION("sopc") = sopc
  ELSE
     sopc = SESSION("sopc")
  END IF


 LISTA = ""
 DESTINO = ""


    STABLA = "PROYECTO"

    if sopcion = "SI" then
       prycod = OWKF.GENPROY("NEW")

     '  response.write prycod & "<BR>"

       sanio = datepart("yyyy",date())
       smes  = datepart("m",date())
       smesC = smes

       Strsql = "Select * From meses WHERE mes = " & smes
       set rs = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
       if not rs.eof then
          smesC = RS(2)
       END IF 

       session("NCliente") = prycod
       Session("VPLAN")    = 0 

       OWKF.GENPROP()

    '   response.write prycod & "-2<BR>"

       STRSQL = "UPDATE PROYPRTY SET CONF = 1 WHERE CODIGO = '" & session("NCliente") & "'  "
       set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

       OWKF.GENINITPRY()

    ''   YXETP = OWKF.GENPRYETP(Session("NCliente"), Session("VPLAN"), "10", "INCIDENCIAS REALIZADAS", "0", "0")

     '  response.write prycod & "-3<BR>"

       PRYTIT = "INCIDENCIAS " & session("segusuario") & " MES " & smesC & "-" & sanio

       StrSql = "UPDATE PROYECTO SET TITULO = '" & PRYTIT & "',CUENTA = 2, MODAL = 0  WHERE CODIGO = '" & prycod & "'"
       set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

       wdescri = "ADMINISTRACION DE INCIDENCIAS"

       XXXX = OWKF.PRYLNCH(prycod,wdescri,"S")

       StrSql = "DELETE From PROYTAREAS WHERE CODIGO = '" & Session("NCliente") & "' "
       set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

       XXXX =  GENSECTOROR() 

       XXXX =  GENUSRCONT()


     '  response.write prycod & "-4<BR>"


       session("NCliente") = ""


    end if


    IF SESSION("XESTAD") = 12 THEN XESTADDSC = "EN CURSO"
    IF SESSION("XESTAD") = 13 THEN XESTADDSC = "FINALIZADOS"


    LISTA = " PROYECTOS " & XESTADDSC
    DESTINO = " ASIGNADOS A " & SESSION("SEGUSUARIO")

    REQWHERE2 = " WHERE ASIGNADO  = '" & SESSION("USUARIO") & "' and estado = " & SESSION("XESTAD") & " AND MODAL = 0 ORDER BY FECHASOL,TITULO"
    REQWHERE = " WHERE ASIGNADO  = '" & SESSION("USUARIO") & "' and estado = " & SESSION("XESTAD") & " AND MODAL = 0 ORDER BY FECHASOL desc"


    celdas = ofv.GenCelda("","TTT ","","33%","","","",LISTA & DESTINO,"si")
    filas = ofv.GenRow("","","","","22","","",celdas,"si")
    sHTML = sHTML & ofv.GenTabla("","","","100%","","0","","0","0","","",filas,"si")


    celdas = ofv.GenCelda("","TTR ","center","33%","","","","<A CLASS=tt HREF='pryadmincidenciasdet.asp?XESTAD=12' >EN CURSO</A>","si")
    celdas = celdas & ofv.GenCelda("","TTR ","center","33%","","","","<A CLASS=tt HREF='pryadmincidenciasdet.asp?XESTAD=13' >FINALIZADOS</A>","si")


    filas = ofv.GenRow("","","","","22","","",celdas,"si")
    sHTML = sHTML & "<CENTER>" & ofv.GenTabla("","","","50%","","0","","0","0","","",filas,"si") & "</CENTER>"










    sImagenO  = "<B><FONT FACE=WEBDINGS COLOR=#6699cc SIZE=3>&#0128;</FONT></B>"
    sImagenI  = "<B><FONT FACE=WEBDINGS COLOR=#AA0000 SIZE=3>&#0128;</FONT></B>"
    sImagenE  = "<B><FONT FACE=WINGDINGS COLOR=#f7aa22 SIZE=3>&#0240;</FONT></B>"
    sImagenC  = "<B><FONT FACE=WINGDINGS COLOR=#008000 SIZE=3>R</FONT></B>"

    sImagenX  = "<B><FONT FACE=WEBDINGS COLOR=#f7aa22 SIZE=3>4</FONT></B>"
    sImagenIX = "<B><FONT FACE=WEBDINGS COLOR=#aa0000 SIZE=3>2</FONT></B>"

    sImagenpry = "<B><FONT FACE=WEBDINGS COLOR=#aa0000 SIZE=3>&#0143;</FONT></B>"
    sImagenreq = "<B><FONT FACE=WEBDINGS COLOR=#006600 SIZE=3>&#0153;</FONT></B>"




 DIM ATABLA(50,2) 

 ITAB=0
 TTAB=0 
 ITABB=0




    celdas = ""



  'Encabezados
  formulario = formulario & ofv.GenCelda("","UT colspan=8","","","","","","&nbsp;","si")





     StrSql = "Select * From " & STABLA & REQWHERE
   '  response.write strsql & "<br>"
     set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     if not rs.eof then
        rs.move(CantRegMoverx)
  posi = 0
        formulario = formulario & GENLISTA(RS)

     ELSE

        celdas = ofv.GenCelda("","VT colspan=8","center","","","","","<BR><BR><BR>NO HAY " & STABLA & "S PARA LA SELECCION<BR><BR><BR>","si")
        formulario = formulario & ofv.GenRow("","aaa  ","","","","","",celdas,"si")

     END IF





  sHTML = sHTML  & ofv.GenTabla("","","","98%","","0","","-1","-1","","",formulario,"si")


  'Botonera
  Agregar = ""
  if bAgregar then
    sAccion = "PRyadmincidenciasdet.asp?Opcion=SI&OPC=1&TIPO=" & STIPO & "&RESP=" & SRESP & "'  "
    Agregar = REPLACE(ofv.BotonAgregar(sAccion),"Agregar","NUEVO " & STABLA)
  end if

  Buscar = ""
  

  
  sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,"")

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)


  sHTML = sHTML  & "</body></html>"
end if


response.write sHTML








FUNCTION GENLISTA(RSX)





   xTrColor = "#FFFFFF"

  posi = 0
  prow = "N" 
  Do While (Not rsX.EOF) and (posi < CantRegMostrarx)

    sform = "REQ" & rsX(0)

    if prow = "S" then
       prow = "N"
       sr   = " bgcolor=#eeeeee "
    else 
       prow = "S"
       sr   = ""
    end if

       sr   = ""
       sc = "ut"




    if bSuprimir and SESSION("XESTAD") = 12 then

          StrSql = "Select ID From PROYTAREAS WHERE CODIGO = '" & rsX(1) & "' "
          set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
          IF RS2.EOF THEN
             sHref = "PRYREQDELETE.asp?VLV=INC&ID=" & rsX(1) & "&CantRegMover=" & CantRegMoverx
             CeldaS = ofv.BotonEliminar( " proyecto",sHref)
          ELSE
             CeldaS = ofv.GenCelda("","","","","","","","&nbsp;","si")
          END IF



    else
      CeldaS = ofv.GenCelda("","","","","","","","&nbsp;","si")
    end if





       xlink = "<A  href='&nbsp;' class=nt "

       xlink = xlink & " onclick=" & chr(34) & "AbrirCons('editarPRY.asp?codigo=" & rsx(1) & "&tipo=PRY'); return false;" & chr(34)
       xlink = xlink & " >" & sImagenpry & "</A>"

       xlink2 = "" & rsX(2) & "" 



      sInput = xlink

      celdas = celdas & ofv.GenCelda("","ut","","1%","","","",sInput,"si")



      sInput = xlink2

      celdas = celdas & ofv.GenCelda("","ut","","55%","","","",sInput,"si")



      sInput = "<b>" & sImagenE & "</b>"
      celdas = celdas & ofv.GenCelda("","ut","","1%","","","",sInput,"si")

    if clng(rsx(4)) = 0 then
      dFecha = "---"
    else
      dFecha = cdate(rsx(4) + clng(cdate("01-01-1900")))
    end if

      sInput = dFecha



      celdas = celdas & ofv.GenCelda("","ut","","15%","","","",sInput,"si")


      sInput = "<b>" & sImagenI & "</b>"
      celdas = celdas & ofv.GenCelda("","vt","","1%","","","",sInput,"si")

      sniv = "---"

                strsql = "Select * From USUARIOS where ID = '" & RSX(14) & "' "
                set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
                if not rsniv.eof then sniv = rsniv(1)
                call ofv.cerrarconsulta(rsniv)

      sInput = sniv
      celdas = celdas & ofv.GenCelda("","ut" ,"","15%","","","",sInput,"si")



       sinput = "<a class=nt  STYLE='CURSOR:HAND;' href='pryadmtskdet.asp?pry=" & rsx(1) & "&vplan=" & rsx(17) & "' "
       sinput = sinput  & " title='Seleccionar&nbsp;proyecto' "
       sinput = sinput  & " >&nbsp;SELECCIONAR&nbsp;</a>" 

      celdas = celdas & ofv.GenCelda("","ut STYLE='border-style:outset;border-width:2;border-color:#993333;' ","","10%","","","",sInput,"si")



    filas = ofv.GenRow("","aaa  BGCOLOR='" & xTrColor & "' ","","","","","",celdas,"si")




    formu = formu & filas

      celdas = ofv.GenCelda("","AAA COLSPAN=8 style='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#CCCCCC;FONT-SIZE:5pt' ","","1%","","","","&nbsp;","si")

    formu = formu & ofv.GenRow("","aaa  BGCOLOR='" & xTrColor & "' ","","","","","",celdas,"si")


                         if xTrColor = "#FFFFFF" then
                            xTrColor = "#F9F9F9" 
                         ELSE
                            xTrColor = "#FFFFFF" 
                         END IF

    posi = posi + 1
    rsX.Movenext
  loop



  GENLISTA = FORMU

END FUNCTION

%>