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

  set cn = ofv.conectar(ofv.strconn6)
  set cn2 = ofv.conectar(ofv.strconn0)
  set cn3 = ofv.conectar(ofv.strconn2)

  ofv.ObtenerAtributos session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar





  PRY = REQUEST.QUERYSTRING("PRY")

  IF PRY <> "" THEN  session("NCliente") = PRY

  StrSql = "Select * From PROYRESP where codigo = '" & session("NCliente") & "' AND TIPO = 'AN' "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if rs.eof then

     GENSECTOROR()


  END IF



        XCONTROL = "S" 
        XCPT     = "S" 
        XCADV    = "S" 






 sopcion = request.querystring("opcion")
 sACCION = request.querystring("ACCION")

if nFecini < clng(cdate("01-01-1900")) then
   nFecini = clng(cdate("01-01-1900"))
else
   nFecini = 0
end if

vNuevoValor = clng(DATE()) - nFecini

 FUNCION = ""




  sHTML = ofv.MenuHeader("",funcion)

''  shtml = replace(shtml,"<body ","<body " & funcion)

  sHTML = sHTML  & "<script language=javascript> "
  sHTML = sHTML  & "function AbrirCons(Datos) { "
  sHTML = sHTML  & " myWin= open(Datos," & chr(34) & "CPRY" & chr(34) & "," & chr(34) & "resizable=yes,menubar=no,toolbar=no,height=570,width=790" & chr(34) & " ); "
  sHTML = sHTML  & " } "
''  sHTML = sHTML  & "</script>"

''  shtml = shtml & "<script  language=javascript>" & chr(13)
  shtml = shtml & " var cnt = 0; " & chr(13)
  shtml = shtml & " function dos() {  " & chr(13)
  shtml = shtml & " if (document.all.odiv) { document.all.odiv.style.top=document.body.clientHeight - document.all.odiv.clientHeight; " & chr(13)
  shtml = shtml & " } else {  if (cnt > 10) {} else { cnt++; " & chr(13)
  shtml = shtml & " setTimeout (" & chr(34) & "dos()" & chr(34) & ", 350); }}" & chr(13)
  shtml = shtml & " } " & chr(13)
  shtml = shtml & " </script>" & chr(13)



    sImagenO = "<B><FONT FACE=WEBDINGS COLOR=#6699cc SIZE=3>&#0128;</FONT></B>"
    sImagenI = "<B><FONT FACE=WEBDINGS COLOR=#AA0000 SIZE=3>&#0128;</FONT></B>"
    sImagenE = "<B><FONT FACE=WINGDINGS COLOR=#f7aa22 SIZE=3>&#0240;</FONT></B>"
    sImagenC = "<B><FONT FACE=WINGDINGS COLOR=#008000 SIZE=3>R</FONT></B>"
    sImagenN = "<B><FONT COLOR=#993333 FACE=WEBDINGS SIZE=3>i</FONT></B>"
    sImagenR = "<B><FONT COLOR=#993333 FACE=WEBDINGS SIZE=3>&#0254;</FONT></B>"
    sImagenI3 = "<B><FONT FACE=WEBDINGS COLOR=#cc0000 SIZE=4>&#064;</FONT></B>"
    sImagenPR = "<B><FONT FACE=SYMBOL COLOR=#AA0000 SIZE=4>!</FONT></B>" 


CARGAR = "S"

  IF CARGAR = "S" THEN

  StrSql = "Select * From PROYECTO where codigo = '" & session("NCliente") & "' "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if NOT rs.eof then
     xCLIENTE = RS(2) 

     VPLAN = RS(17)
     SESSION("VPLAN") = VPLAN




    sHTML = sHTML & "<center>"

    sImagen = "<FONT FACE=WINGDINGS SIZE=4>4</FONT>"

  celdas = ofv.GenCelda("","cc colspan=3","","","","","",sImagen & "&nbsp;<BIG><B>PROPIEDADES DEL MODELO</B></BIG>","si")
  filas  = ofv.GenRow("","","","","10","","",celdas,"si")

  celdas = ofv.GenCelda("","AA COLSPAN=3","CENTER","","","","","&nbsp;","si")
  filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


    sImagen = "<FONT COLOR=#993333 FACE=WEBDINGS SIZE=2>4</FONT>"

  celdas = ofv.GenCelda("","vt","","25%","","","","CODIGO:","si")
  celdas = celdas & ofv.GenCelda("","VT","","1%","","","","&nbsp;","si")

  celdas = celdas & ofv.GenCelda("","vt","","70%","","","",rs(1),"si")
  filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")




    celdas = ofv.GenCelda("","vt","","","","","","TITULO:","si")
    celdas = celdas & ofv.GenCelda("","","","","","","",sImagen,"si")
    if bModificar then
      sInput = ofv.GenerarInput(rs(2).name,rs(2),"text","Usuarios","70","dt","")
    else
      sInput = ofv.GenerarInput("",rs(2),"readonly","","70","dt","")
    end if
    celdas = celdas & ofv.GenCelda("","vt","","","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")



    celdas = ofv.GenCelda("","vt","","","","","","CLASE:","si")


       SINPUT1 = sImagenN    
       SINPUT2 = "MODELO DE PROYECTO"


    celdas = celdas & ofv.GenCelda("","VT","","","","","",sInput1,"si")

    celdas = celdas & ofv.GenCelda("","vt","","","","","",sInput2,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


''''''''''''

 

      strsql = "Select * From PROYPLAN where VPLAN = " & VPLAN & " AND CODIGO = '" & session("NCliente") & "' "
        set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if not rsniv.eof then sniv = rsniv(3)
        call ofv.cerrarconsulta(rsniv)
      sInput = VPLAN & ". " & sniv

     SESSION("VPLAND") = sInput 




    celdas = ofv.GenCelda("","vt","","","","","","TIPO:","si")
    celdas = celdas & ofv.GenCelda("","","","","","","",sImagen,"si")
    if bModificar then
      strsql = "Select * From TIPOS Order By DESCRI"
      sInput = ofv.GenerarCombo("Usuarios",rs(8).name,rs(8),strsql,cn,0,1,"")
    else
      sniv = ""
      strsql = "Select * From TIPOS where ID = " & rs(8)
        set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if not rsniv.eof then sniv = rsniv(1)
        call ofv.cerrarconsulta(rsniv)
      sInput = ofv.GenerarInput("",sniv,"readonly","","40","dt","")
    end if
    celdas = celdas & ofv.GenCelda("","vt","","","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

    celdas = ofv.GenCelda("","vt","","","","","","ORIGEN:","si")
    celdas = celdas & ofv.GenCelda("","","","","","","",sImagen,"si")
    if bModificar then
      strsql = "Select * From ORIGEN Order By DESCRI"
      sInput = ofv.GenerarCombo("Usuarios",rs(7).name,rs(7),strsql,cn3,0,1,"")
    else
      sniv = ""
      strsql = "Select * From ORIGEN where OBJID = 0" & rs(7)
        set rsniv = ofv.crearconsultaEx(StrSql,cn3,1,parametros)
        if not rsniv.eof then sniv = rsniv(1)
        call ofv.cerrarconsulta(rsniv)
      sInput = ofv.GenerarInput("",sniv,"readonly","","40","dt","")
    end if
    celdas = celdas & ofv.GenCelda("","vt","","","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")





    celdas = ofv.GenCelda("","vt","","","","","","DESCRIPCION ORIGEN:","si")
    celdas = celdas & ofv.GenCelda("","","","","","","",sImagen,"si")
    if bModificar then
      sInput = ofv.GenerarInput(rs(12).name,rs(12),"text","Usuarios","70","dt","")
    else
      sInput = ofv.GenerarInput("",rs(12),"readonly","","70","dt","")
    end if
    celdas = celdas & ofv.GenCelda("","vt","","","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")





      celdas = ofv.GenCelda("","vt  ","","","","","","MODO DE PLANIFICACION:","si")
      celdas = celdas & ofv.GenCelda("","","","","","","",sImagen,"si")
    if bModificar AND (RS(11) = 1 OR RS(11) = 99)  then
        sInput = "<font class=at><select name=" & rs(18).name & " onchange=submit()>"
        sInput = sInput & "<option value=0 "
        if rs(18) = 0 then sInput = sInput & "selected"
        sInput = sInput & " >DIAS CORRIDOS<option value=1 "
        if rs(18) = 1 then sInput = sInput & "selected"
        sInput = sInput & " >DIAS HABILES</select></font>"
    else
        if rs(18) = 0 then sInput = "DIAS CORRIDOS"
        if rs(18) = 1 then sInput = "DIAS HABILES"
    end if
    celdas = celdas & ofv.GenCelda("","vt","","","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


    celdas = ofv.GenCelda("","vt","","","","","","PRIORIDAD:","si")
    celdas = celdas & ofv.GenCelda("","","","","","","",sImagenPR,"si")

    if bModificar then
       StrSql = "Select * From PRIORIDAD ORDER BY ID"
      sInput = ofv.GenerarCombo("Usuarios",rs(19).name,rs(19),strsql,cn,0,1,"")
    else
     StrSql = "Select * From PRIORIDAD where ID = '" & RS(19) & "' "
     set rsX2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     if NOT rsX2.eof then
        sInput  = rsX2(1)

     ELSE
        sInput  = "SIN DETERMINAR"

     END IF
    end if
    celdas = celdas & ofv.GenCelda("","vt","","","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

    celdas = ofv.GenCelda("","vt","","","","","","FECHA CREACION:","si")
    celdas = celdas & ofv.GenCelda("","","","","","","","","si")

    if clng(rs(3)) = 0 then
      dFecha = "---"
    else
      dFecha = cdate(rs(3) + clng(cdate("01-01-1900")))
    end if

      sInput = dFecha


    celdas = celdas & ofv.GenCelda("","vt","","20","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")









    celdas = ofv.GenCelda("","vt","","","","","","AUTOR:","si")
    celdas = celdas & ofv.GenCelda("","","","","","","","<B>" & sImagenO & "</B>","si")
    if RS(13) = "NO" then
      sInput = "SIN AUTOR"
      celdas = celdas & ofv.GenCelda("","vT","","","","","",sInput,"si")
    else
      strsql = "Select * From USUARIOS where ID = '" & rs(13) & "' "
        set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
        if not rsniv.eof then sniv = rsniv(1)
        call ofv.cerrarconsulta(rsniv)
      sInput = "<B>" & sniv & "</B>"
      celdas = celdas & ofv.GenCelda("","vt","","","","","",sInput,"si")
    end if

    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")



    if bModificar then
      sAccion = "asp/Actualizar.asp?FNAME=" & rs(0) & "&volver=" & vuelta & "&Tabla=PROYECTO"
      formulario0 = ofv.GenForm("Usuarios","",sAccion,"","post","",filas,"si")
    else
      formulario0 = ofv.GenForm("","","","","post","",filas,"si")
    end if







    filas = ofv.GenRow("","","","","","","",formulario0,"si")
    Celda = ofv.GenCelda("","abc colspan=3","center","100%","","","","&nbsp;","si")
    FilaS =  filaS & ofv.GenRow("","","","","","","",Celda,"si")

    FilaS = ofv.GenTabla("","","","100%","","0","","0","0","","",filas,"si")



    celdas = ""



       sinput = "<a class=tt  href=proyresp.asp"
       sinput = sinput  & " ><b><FONT FACE=WEBDINGS size=2 >Ä</FONT>  RESPONSABLES</b></a>" 
       celdas = celdas & ofv.GenCelda("","tt","center","","10%","","",sInput,"si")

       sinput = "<a class=tt  href=proytask.asp"
       sinput = sinput  & " ><b><FONT FACE=WEBDINGS size=2 >2</FONT>  TAREAS</b></a>" 
       celdas = celdas & ofv.GenCelda("","tt","center","","10%","","",sInput,"si")

       sinput = "<a class=tt  href=proyest.asp"
       sinput = sinput  & " ><b><FONT FACE=WEBDINGS size=2 >ù</FONT>  ESTADO</b></a>" 
       celdas = celdas & ofv.GenCelda("","tt","center","","10%","","",sInput,"si")

          sinput = "<a class=tt  href=PROYMODEL.ASP"
          sinput = sinput  & " ><b><FONT FACE=WEBDINGS size=2 >:</FONT>  VOLVER</b></a>" 
          celdas = celdas & ofv.GenCelda("","tt","center","","10%","","",sInput,"si")


    cel = ofv.GenRow("","","","","","","",celdas,"si")


    FilaS =  filaS & "<BR><BR>" & ofv.GenTabla("UNO","aaa style='border-top-style:ridge;border-top-width:2;border-top-color:#f7aa22;' ","","100%","","0","","4","0","","",cel,"si")



    Celda = ofv.GenCelda("","","center","","","","",FilaS,"si")
    FilaS = ofv.GenRow("","","","","","","",Celda,"si")


    sHTML = sHTML & ofv.GenTabla("","aa style='filter:shadow(color=#dddddd,direction=125);' ","","80%","","0","","0","0","","",filas,"si")
    sHTML = sHTML & "</center><br><br>"

 



 
  end if

 END IF

  fill1 = ofv.gencelda("","vt","left","","1%","","","<img src='../coress20/imagenes/icorontime.jpg' border=0>","si")
  fill = ofv.gencelda("","vt","right","","90%","","",fill & "<big>Cor&nbsp;<font color=#f7aa66 size=2><i>On</i></font>Time&nbsp;</big>&reg;&nbsp;","si")



    formu = ofv.GenRow("","aaa height=30","","","","","",fill1 & fill,"si")
    formu = ofv.GenTabla("","aaa style='border-top-style:ridge;border-top-width:2;border-top-color:#f7aa22;' ","","100%","","0","","-2","-1","","",formu,"si")





  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)


  sHTML = sHTML  & "</body></html>"
end if

response.write sHTML

FUNCTION GETTEMASUBTEMA(SSUBTEMA)

  xTEMA = 0
  xSUBTEMA = "SIN ASIGNAR"

  SQL2 = "SELECT * "
  SQL2 = SQL2 & " FROM csbuscasubtema where "
  SQL2 = SQL2 & " codPRODSERV = " & SSUBTEMA
  set rsX = ofv.crearconsultaEx(SQL2,cn3,1,parametros)
  if NOT rsX.eof then
     xTEMA = RSX(3)
     xSUBTEMA = RSX(1)
     SQL1 = "SELECT * From Tema WHERE CODTEMA = " & xTEMA
     set rsX = ofv.crearconsultaEx(SQL1,cn3,1,parametros)
     if NOT rsX.eof then
        xTEMA = RSX(1)
     END IF

  END IF

  GETTEMASUBTEMA = XTEMA & CHR(10) & XSUBTEMA

END FUNCTION


%>