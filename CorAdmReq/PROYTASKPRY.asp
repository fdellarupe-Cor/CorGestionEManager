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

 sopcion = request.querystring("opcion")
 sACCION = request.querystring("ACCION")

if nFecini < clng(cdate("01-01-1900")) then
   nFecini = clng(cdate("01-01-1900"))
else
   nFecini = 0
end if

vNuevoValor = clng(DATE()) - nFecini



  sHTML = ofv.FormHeader("")



    sImagenO = "<B><FONT FACE=WEBDINGS COLOR=#6699cc SIZE=3>&#0128;</FONT></B>"
    sImagenI = "<B><FONT FACE=WEBDINGS COLOR=#AA0000 SIZE=3>&#0128;</FONT></B>"
    sImagenE = "<B><FONT FACE=WINGDINGS COLOR=#f7aa22 SIZE=3>&#0240;</FONT></B>"
    sImagenC = "<B><FONT FACE=WINGDINGS COLOR=#008000 SIZE=3>R</FONT></B>"
 


CARGAR = "S"

  IF CARGAR = "S" THEN

  StrSql = "Select * From PROYECTO where codigo = '" & session("NCliente") & "' "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if NOT rs.eof then
     xCLIENTE = RS(2) 
     xSTS     = RS(11) 





    sHTML = sHTML & "<center>"

    sImagen = "<FONT FACE=WINGDINGS SIZE=4>4</FONT>"

  celdas = ofv.GenCelda("","ttt colspan=3","","","","","",sImagen & "&nbsp;INICIO DEL PROYECTO " & rs(2) & " ","si")
  filas  = ofv.GenRow("","","","","10","","",celdas,"si")

    XFSP = RS(3)
    XFEP = RS(5)
    XFSR = RS(4)
    XFER = RS(6)

    IF XFSR = 0 THEN XFSR = XFSP

    StrSql = "Select tarea From PROYETPTSK where codigo = '" & session("NCliente") & "' and VPLAN = " & session("vplan") & " ORDER BY ORDEN desc,TSKORD DESC "
    set rs5 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    IF NOT rs5.EOF THEN
       OWKF.YFMAXTSK = rs5(0)

       SXFSTARTD = OWKF.GETFENDV(rs5(0), session("NCliente"))

       XIIX = OWKF.TIX


       FSPLAN = OWKF.ATABTSK(cint(XIIX),7)
  ''     FEPLAN = OWKF.ATABTSK(cint(XIIX),8)
       FSREAL = OWKF.ATABTSK(cint(XIIX),9)

       FEPLAN = XFEP
       FEREAL = OWKF.ZXMFMAX
       
    else
       SXFSTARTD = OWKF.GETFENDV(session("NCliente"), session("NCliente"))

       XIIX = OWKF.TIX


       FSPLAN = OWKF.ATABTSK(cint(XIIX),7)
''       FEPLAN = OWKF.ATABTSK(cint(XIIX),8)
       FSREAL = OWKF.ATABTSK(cint(XIIX),9)

       FEPLAN = XFEP
       FEREAL = OWKF.ATABTSK(cint(XIIX),10)

    end if


    call ofv.cerrarconsulta(rs5)



    celdas = ofv.GenCelda("","ut","","","","","","FECHA PLAN INICIAL:","si")
    celdas = celdas & ofv.GenCelda("","","","","","","","","si")

    if clng(XFSP) = 0 then
      dFecha = "---"
    else
      dFecha = cdate(XFSP + clng(cdate("01-01-1900")))
    end if

      sInput = dFecha


    celdas = celdas & ofv.GenCelda("","vt","","20","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

    

    celdas = ofv.GenCelda("","ut","","","","","","FECHA INICIO:","si")
    celdas = celdas & ofv.GenCelda("","","","","","","","","si")

    if clng(XFSR) = 0 then
      dFecha = "---"
    else
      dFecha = cdate(XFSR + clng(cdate("01-01-1900")))
    end if

      sInput = dFecha


    celdas = celdas & ofv.GenCelda("","vt","","20","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")






    celdas = ofv.GenCelda("","ut","","","","","","FECHA PLANIFICADA ENTREGA:","si")
    celdas = celdas & ofv.GenCelda("","","","","","","","","si")

    if clng(FEPLAN) = 0 then
      dFecha = "---"
    else
      dFecha = cdate(FEPLAN + clng(cdate("01-01-1900")))
    end if

      sInput = dFecha


    celdas = celdas & ofv.GenCelda("","vt","","20","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")



    celdas = ofv.GenCelda("","ut","","","","","","FECHA FINALIZACION:","si")
    celdas = celdas & ofv.GenCelda("","","","","","","","","si")

    if clng(FEREAL) = 0 then
      dFecha = "---"
    else
      dFecha = cdate(FEREAL + clng(cdate("01-01-1900")))
    end if

      sInput = dFecha


    celdas = celdas & ofv.GenCelda("","vt","","20","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


    FilaS = ofv.GenTabla("","AAA style='border-style:solid;border-width:1;border-color:#f7aa22;' ","","100%","","0","","-1","2","","",filas,"si") & "<BR>"

'''''''''''''''''''''

    IF xSTS = 1 THEN

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

    if smesc = 0 then smesc = smes

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
     sInput = ofv.GenerarCombo(sForm & "A","smesc",smesc,Strsql,cn2,1,2,"sb0")
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
           t1 = t1 & ofv.GenerarInput(rs(3).name,(afechs(ii,1) - srest),"hidden","","2","dt","")
           t1 = ofv.GenCelda("","VT","center","10%","","","",t1,"si")
           t1 = ofv.GenForm("feac" & ii,"",sAccion,"","post","",t1,"si") 
           Celda = Celda & t1

        end if


    next

'''


   filas = filas & ofv.GenTabla("","AAA style='border-style:ridge;border-width:2;border-color:#f7aa22;'","","50%","","0","","0","0","","",FILA,"si")

 
  END IF


'''''''''''''''''''''




    sHTML = sHTML & filas
    sHTML = sHTML & "</center><br><br>"

 


  end if

 END IF

  Agregar = ""
  Buscar  = ""
  sAccion = "PROYTASK.asp"
  Celda = ofv.BotonVolver(sAccion)



  sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","no",Agregar,Buscar,Celda)



  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)


  sHTML = sHTML  & "</body></html>"
end if

response.write sHTML




Function Fechador(wkfecha)


    
      fec = cdate(wkfecha + clng(cdate("01-01-1900")))
      ndia = ofv.cdia(Weekday(fec))
      nMes = ofv.cmes(Month(fec))
    
      Fechador = "  " & ndia & ", " & Day(fec) & " de " & nMes & " de " & Year(fec)


End Function

%>