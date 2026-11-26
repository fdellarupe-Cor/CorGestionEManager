<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()

if ok <> "Y" then
  sHTML = ok
else
  response.expires=0






  CantRegMostrarx = 12
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))
  vuelta = "../EDREQUER.ASP?ACCION=S_CANTREGMOSTRAR=1_CANTREGMOVER=1"

  set cn = ofv.conectar(ofv.strconn6)
  set cn2 = ofv.conectar(ofv.strconn0)
  set cn3 = ofv.conectar(ofv.strconn2)

  ofv.ObtenerAtributos session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

'  IF session("ModMain") <> "WF" THEN
'     bAgregar = FALSE
'     bModificar = FALSE
'     bSuprimir = FALSE
'  END IF

 sopcion = request.querystring("opcion")
 sACCION = request.querystring("ACCION")

if nFecini < clng(cdate("01-01-1900")) then
   nFecini = clng(cdate("01-01-1900"))
else
   nFecini = 0
end if

vNuevoValor = clng(DATE()) - nFecini


 FUNCION = ""

  sHTML = ofv.FormHeader(replace(session("FormN"),"_",FUNCION))

 





 

  StrSql = "Select * From REQUERIMIENTO where codigo = '" & session("NCliente") & "' "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if NOT rs.eof then
     xCLIENTE = RS(2) 






    sHTML = sHTML & "<center>"

    sImagen = "<FONT FACE=WINGDINGS SIZE=4>4</FONT>"

  celdas = ofv.GenCelda("","cc colspan=3","","","","","",sImagen & "&nbsp;<BIG><B>FECHA DE ENTREGA ESPERADA</B></BIG>","si")
  filas  = ofv.GenRow("","","","","10","","",celdas,"si")

  celdas = ofv.GenCelda("","AA COLSPAN=3","CENTER","","","","","&nbsp;","si")
  filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

  celdas = ofv.GenCelda("","tt","","","","","","clase:","si")
  celdas = celdas & ofv.GenCelda("","","","","","","","","si")

     StrSql = "Select * From TIPOREQ where TIPOREQ = '" & RS(15) & "' "
     set rsX2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     if NOT rsX2.eof then
        XTIPOREQ  = RS(15) & " - " & rsX2(2)

     ELSE
        XTIPOREQ  = "SIN DETERMINAR"

     END IF

  celdas = celdas & ofv.GenCelda("","cc","","","","","","<B>" & REPLACE(TRIM(XTIPOREQ)," ","&nbsp;") & "</B>","si")
  filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


  celdas = ofv.GenCelda("","tt","","","","","","Codigo:","si")
  celdas = celdas & ofv.GenCelda("","","","","","","","","si")

  celdas = celdas & ofv.GenCelda("","cc","","","","","",rs(1),"si")
  filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

    sImagen = "<FONT COLOR=#993333 FACE=WEBDINGS SIZE=2>4</FONT>"


    celdas = ofv.GenCelda("","tt","","","","","","TITULO:","si")
    celdas = celdas & ofv.GenCelda("","","","","","","","&nbsp;","si")



    celdas = celdas & ofv.GenCelda("","CC","","","","","","<B>" & rs(2) & "</B>","si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")



    celdas = ofv.GenCelda("","tt","","","","","","FECHA INICIO:","si")
    celdas = celdas & ofv.GenCelda("","","","","","","","","si")

    if clng(rs(3)) = 0 then
      dFecha = "---"
    else
      dFecha = cdate(rs(3) + clng(cdate("01-01-1900")))
    end if

      sInput = dFecha


    celdas = celdas & ofv.GenCelda("","cc","","20","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


    celdas = ofv.GenCelda("","tt","","","","","","FECHA ESPERADA:","si")
    celdas = celdas & ofv.GenCelda("","","","","","","",sImagen,"si")

    dFecha = vNuevoValor
    dFecha = cdate(dFecha + clng(cdate("01-01-1900")))



    sInput = ofv.GenerarInput(rs(16).name,dFecha,"agregar","Usuarios","70","dt","")
    celdas = celdas & ofv.GenCelda("","cc","","20","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

    celdas = ofv.GenCelda("","aaa colspan=3","","20","","","","","si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


    sInput = "<input  type=SUBMIT  size=10%  value='aceptar' class=bt >"
    celdas = ofv.GenCelda("","","","","","","","","si")
    celdas = celdas & ofv.GenCelda("","","","","","","","","si")
    Celdas = celdas & ofv.gencelda("","","center","","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


    if bModificar then
      sAccion = "asp/Actualizar.asp?FNAME=" & rs(0) & "&volver=" & vuelta & "&Tabla=REQUERIMIENTO"
      formulario0 = ofv.GenForm("Usuarios","",sAccion,"","post","",filas,"si")
    else
      formulario0 = ofv.GenForm("","","","","post","",filas,"si")
    end if


    filas1 = ofv.GenTabla("","AAA style='border-style:ridge;border-width:2;border-color:#f7aa22;'","","50%","","0","","0","0","","",formulario0,"si")

'''''''''
 cniv = request.querystring("Cniv")
  cnom = request.querystring("Cnom")
  cx = request.querystring("Cx")
  cA = request.querystring("CA")

  if cniv = "" then 
     cniv = session("ENC7")
     cnom = session("ENC8")
     cx = session("ENC9")
     cA = session("ENCA")
  else
     session("ENC7") = cniv
     session("ENC8") = cnom
     cx = cdbl(cx)
     session("ENC9") = cx
     session("ENCA") = cA
  end if






  posi = 0
    sForm = "form" & posi
       sclase = "CC" 

     Celda = ofv.GenCelda("","VT","","15%","","","","FECHA:","si")



    if clng(XFSR) = 0 then
       sInput1 = "------------"
    else
       sInput1 = cdate(clng(XFSR) + clng(cdate("01-01-1900")) )
       swd = datepart("w",sInput1)
       sInput1 = Fechador(XFSR)
    END IF


     Celda = Celda & ofv.gencelda("","VT colspan=3 ","","80%","","","",sInput1,"si")



    Fila = ofv.genrow("","","","","","","",Celda,"si")

  ''  filas = filas & ofv.GenTabla("","","","50%","","0","","0","0","","",Fila,"si")

    'Descripcion
    CELDA = ""

    sanio = datepart("yyyy",date())
    smes = datepart("m",date())

    SMESC = SMES

    if not isnull(XFSR) then

       if clng(XFSR) > 0 then

          smesc = cdate(clng(XFSR) + clng(cdate("01-01-1900")) )
          smesc = datepart("m",smesc)

       end if 


    end if

    IF CA <> "" THEN 
       smesc = request.form("smesc")
 '      response.write "smesc " & smesc & " ca " & ca
       if isnull(smesc) then  SMESC = ca
       if len(smesc) = 0 then  SMESC = ca
       ca =  SMESC
      
    else
       ca = smesc

    end if  

  '     response.write "smesc " & smesc & " ca " & ca

       session("ENCA") = smesc

    if cint(smesc) = cint(smes) then
       sanioc = sanio
    elseif cint(smesc) < cint(smes) then
       sanioc = sanio + 1
    else
       sanioc = sanio
    end if

    
      Celda = ofv.GenCelda("","VT","","15%","","","","MES:","si")

    'codigo

     Strsql = "Select * From meses Order By mes"
     sInput = ofv.GenerarCombo("document." & sForm & "A","smesc",smesc,Strsql,cn2,1,2,"sb")
     Celda = Celda & ofv.gencelda("","VT","","45%","","","",sInput,"si")

     Celda = Celda & ofv.GenCelda("","VT","","15%","","","","AÑO:","si")
     
     sinput = sanioc 
     Celda = Celda & ofv.gencelda("","VT","","20%","","","",sInput,"si")

    Fila = FILA & ofv.genrow("","","","","","","",Celda,"si")


    sAccion = "PROYTASKPRY.asp?cniv=" & cniv & "&cnom=" &  cnom & "&cx=" & cx & "&ca=" & smesc
    Formulario = ofv.genform(sForm & "A","",sAccion,"","post","",Fila,"si")

    filas = filas & ofv.GenTabla("","AAA style='border-style:ridge;border-width:2;border-color:#f7aa22;'","","50%","","0","","0","0","","",Formulario,"si")


    dim afech(33,1)
    dim afechs(42,1)

    CELDA = ""

    sfecds = "01-" & cstr(smesc) & "-" & cstr(sanioc)

    'sfecd = clng(cdate(sfecds))
    sfecd = clng(dateserial(sanioc,smesc,1))

    'swd = datepart("w",date(sfecds))
     swd = datepart("w",dateserial(sanioc,smesc,1))
     
    sfechs = "28-" & cstr(smesc) & "-" & cstr(sanioc)

    'sfech = clng(date(sfechs))
     sfech = clng(dateserial(sanioc,smesc,28))
 '   response.write " fd " & sfecd & " fh " & sfech

    yy = 0

    for ii = 1 to 3
        sfecha = sfech + ii
        sfecha = cdate(clng(sfecha))
        smess = datepart("m",sfecha)
        if cint(smesc) <> cint(smess) then
           exit for
        else
           yy = ii
        end if

    next  

    sfech = sfech + yy

 '  response.write " wd " & swd & " fd " & sfecd & " fh " & sfech
    srest = clng(cdate("01-01-1900"))
   ' response.write  sfech &" -"&sfecd 
    topefech = sfech - sfecd 

    yy = 0



    wswd = swd

    for ii = 0 to topefech
        
        afech(ii+1,0) = wswd
        wswd = wswd + 1
        if wswd > 7 then wswd = 1        

        afech(ii+1,1) = sfecd + ii


    next


    for ii = 1 to (cint(afech(1,0)) - 1)
        afechs(ii,0) = "NO"
        afechs(ii,1) = 0
    next

    yy = 1

 '   RESPONSE.WRITE "II " & II & " TOPE " & (ii + topefecH ) & " TOPEFECH " & topefecH
    for ii = ii to (ii + topefecH )
        afechs(ii,0) = yy
        afechs(ii,1) = afech(yy,1)
        yy = yy + 1
    next

'    RESPONSE.WRITE "II " & II

    for ii = ii to 42
        afechs(ii,0) = "NO"
        afechs(ii,1) = 0
    next

    CELDA = ""
    Celda = Celda & ofv.GenCelda("","Vt style='border-style:outset;border-width:2;'","CENTER","10%","","","","D","si")
    Celda = Celda & ofv.GenCelda("","Vt style='border-style:outset;border-width:2;'","CENTER","10%","","","","L","si")
    Celda = Celda & ofv.GenCelda("","Vt style='border-style:outset;border-width:2;'","CENTER","10%","","","","M","si")
    Celda = Celda & ofv.GenCelda("","Vt style='border-style:outset;border-width:2;'","CENTER","10%","","","","M","si")
    Celda = Celda & ofv.GenCelda("","Vt style='border-style:outset;border-width:2;'","CENTER","10%","","","","J","si")
    Celda = Celda & ofv.GenCelda("","Vt style='border-style:outset;border-width:2;'","CENTER","10%","","","","V","si")
    Celda = Celda & ofv.GenCelda("","Vt style='border-style:outset;border-width:2;'","CENTER","10%","","","","S","si")

    Fila = ofv.genrow("","","","","","","",Celda,"si")

    CELDA = ""
    yy = 0

    for ii = 1 to 43

        if yy = 7 then
           yy = 1
           Fila = fila & ofv.genrow("","","","","","","",Celda,"si")
           CELDA = ""
        else
           yy = yy + 1 
        end if

        if ii > 42 then exit for  

        if yy = 1 or yy = 7 then
           sclase = "tt3 style='border-style:outset;border-width:2;'" 
        else
           sclase = "TT style='border-style:outset;border-width:2;'" 
        end if 

        if afechs(ii,0) = "NO" then
           sinput = "&nbsp;"
           sclase = "CC " 
           Celda = Celda & ofv.GenCelda("","VT","CENTER","10%","","","",sinput,"si")
        else
           sinput1 = afechs(ii,0)

           sinput1 = cstr(sinput1)

           if len(sinput1) = 1 then sinput1 = "0" & sinput1 

           sinput1 = " " & sinput1 & " " 

           sAccion = "asp/Actualizar.asp?FNAME=" & rs(0) & "&volver=" & vuelta & "&Tabla=PROYECTO"

           t1 = ofv.GenerarInput("bvo1",sInput1,"submit","","10",sclase,"")
           t1 = t1 & ofv.GenerarInput(rs(4).name,(afechs(ii,1) - srest),"hidden","","2","dt","")
           t1 = ofv.GenCelda("","VT","center","10%","","","",t1,"si")
           t1 = ofv.GenForm("feac" & ii,"",sAccion,"","post","",t1,"si") 
           Celda = Celda & t1

        end if


    next

'''


   filas2 = ofv.GenTabla("","AAA style='border-style:ridge;border-width:2;border-color:#f7aa22;'","","50%","","0","","0","0","","",FILA,"si")

'''''''''







    Celda = ofv.GenCelda("","","center","","","","",FilaS1 & filas2,"si")
    FilaS = ofv.GenRow("","","","","","","",Celda,"si")


    sHTML = sHTML & ofv.GenTabla("","aa style='filter:shadow(color=" & Session("CMFMENU") & ",direction=125);' ","","80%","","0","","0","0","","",filas,"si")
    sHTML = sHTML & "</center><br><br>"

 


  end if




  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)


  sHTML = sHTML  & "</body></html>"
end if

response.write sHTML



%>