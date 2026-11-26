<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0


  GLOSDSC = request.querystring("GLOSDSC")







'''  sHTML = ofv.MENUHeader("","")
'''    sHTML = replace(sHTML,"<html>","<html><title>Consultas al Glosario - Descripcion detallada</title>")

  sHTML = "<html><title>Consultas al Glosario</title>" & ofv.formstyler() & "<body>"

  'Encabezados

    Celda = ofv.GenCelda("","TTr COLSPAN=2 ","","","","","","CONSULTAS AL GLOSARIO - DESCRIPCION ","si")

    Fila = ofv.genrow("","","","","25","","",Celda,"si")


    Celda = ofv.GenCelda("","VT","","20%","","","","BUSQUEDA:&nbsp;","si")

    Celda = Celda & ofv.gencelda("","VT","","80%","","","","<BIG>" & TARGET & "</BIG>","si")


    Fila = FILA & ofv.genrow("","","","","","","",Celda,"si")

    sHTML = sHTML & ofv.GenTabla("","","","100% ","","0","","0","0","","",FILA,"si") & "<BR>"

  Formulario = ofv.Encabezados("DESCRIPCION")





  set cn = ofv.conectar(ofv.strconn2)

  Strsql = "Select * From GLOSDSC WHERE CODWORD = " & GLOSDSC & " Order By FECVIG,HORVIG,OBJID "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then 
  posi = 0
  Do While (Not rs.EOF)
    sForm = "form" & posi


      Celda = ofv.GenCelda("","","","","","","","<font face=corflaticon size=3 color=#F7AA22 >4</font>","si")

  GLAUTOR = ""

  Strsql = "Select USRWORD From GLOSusr WHERE CODWORD = " & GLOSDSC & " and CODDSC = " & RS(0) & " AND TIPMOD = 'A' "
  set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs1.eof then 

      Strsql = "Select DESCRIPCION From USUARIOS WHERE ID= '" & RS1(0) & "' "
      set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
      if not rs1.eof then 

         GLAUTOR = "<B><font face=corflaticon size=3 color=#993333 >&#0128;</font></B>&nbsp;" & RS1(0) & "&nbsp;-&nbsp;"

      END IF

  END IF


    sInput1 = cdate(RS(2) + clng(cdate("01-01-1900")))

    s1d = datepart("d",sInput1)
    s1m = datepart("m",sInput1)
    s1a = datepart("yyyy",sInput1)

    sInput1 = s1d & "/" & s1m & "/" & s1a

    sInput2 = GETGLBHMS(RS(3))

    IF LEN(GLBSM) = 1 THEN GLBSM = "0" & GLBSM 
    IF LEN(GLBSS) = 1 THEN GLBSS = "0" & GLBSS 
   
    sInput2 = GLBSH & ":" & GLBSM & ":" & GLBSS

    sInput = GLAUTOR & "&nbsp;&nbsp;" & sInput1 & "&nbsp;" & sInput2
    Celda = Celda & ofv.gencelda("","VT style='border-style:solid;border-width:1;border-color:#f7aa22;' ","right","100%","","","",sInput,"si")

    Fila = ofv.genrow("","","","","","","",Celda,"si")


      Celda = ofv.GenCelda("","","","","","",""," ","si")


    'Tipo
    sInput = REPLACE(rs(4),CHR(13),"<BR>") & "<br><br>"


    Celda = Celda & ofv.gencelda("","UT","","100%","","","",sInput,"si")


    Fila = FILA & ofv.genrow("","","","","","","",Celda,"si")


    Formulario = Formulario & Fila

    posi = posi + 1
    rs.Movenext
  Loop

 ELSE


      Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")



    sInput = "<BR><BR>NO SE ENCONTRARON ELEMENTOS PARA LA SELECCION"
    Celda = Celda & ofv.gencelda("","VT","","100%","","","",sInput,"si")


    Fila = ofv.genrow("","","","","","","",Celda,"si")


    Formulario = Formulario & Fila

 END IF


  ftab = ofv.GenTabla("","","","100% ","","0","","0","0","","",Formulario,"si")

    Celda = ofv.GenCelda("","","center","","","","",ftab,"si")
    FilaS = filas & ofv.GenRow("","","","","","","",Celda,"si")

    sHTML = sHTML & ofv.GenTabla("","aa style='filter:shadow(color=" & Session("APLSHW") & ",direction=125);' ","","100%","","0","","0","0","","",filas,"si")




  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>