<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else

 '' session("ModMain") = "CL"

  response.expires=-1441
  Response.AddHeader "PRAGMA", "NO-CACHE"

  XFNAME = request.querystring("XFNAME")
  IF XFNAME = "" THEN
     XFNAME = "PV"
  END IF

  XFSRD = request.querystring("XFSRD")
  IF XFSRD = "" THEN XFSRD = 0

  XFSRH = request.querystring("XFSRH")
  IF XFSRH = "" THEN XFSRH = 0

  IF XFNAME = "FH" THEN
     XFSRD = request.form("FECACT")
     XFSRH = XFSRD + 1
  ELSEIF XFNAME = "OK" THEN
     XFSRH = request.form("FECACT")
  ELSEIF XFNAME = "PV" THEN
     XFNAME = "OK"
     XFSRD = clng(DATE()) - clng(cdate("01-01-1900")) - 7
     XFSRH = XFSRD + 7
  END IF





  set cn = ofv.conectar(ofv.strconn2)




  sHTML = ofv.FormHeader( session("FormN"))
  sHTML = sHTML & "<center>"

  tt="&nbsp;"

  IF XFNAME = "OK" THEN

     sInput = "<big>SELECCIONE LOS PARAMETROS PARA LA CONSULTA</big>"
     Celda  = ofv.GenCelda("","ttt colspan=2 ","","","","","",sInput,"si")
     Fila   = ofv.genrow("","","","","30","","",Celda,"si")

     Celda  = ofv.GenCelda("","vt","","30%","","","","&nbsp;","si")
     Celda  = Celda & ofv.GenCelda("","vti","","70%","","","","&nbsp;","si")
     Fila   = Fila & ofv.GenRow("","","","","10","","",Celda,"si")

    sImagenE = "<B><FONT FACE=corflaticon COLOR=#F7AA22 SIZE=4>&#0166;</FONT></B>"

    sAccion = "COESTDSEGUI.asp?XFNAME=NO&Desde=" & sID & "&Hasta=" & sID2 & "&Fecha=" & sID3 & "&sorden=" & sorden & "&CantRegMostrar=10&CantRegMover=0"

    sinputE = "&nbsp;&nbsp;<a  class=UT href=" & sAccion
    sinputE = sinputE  & " TITLE='MODIFICAR FECHA' >" & sImagenE & "</a>" 

    Celda   = ""
    Celda   = Celda & ofv.GenCelda("","tt","","","","","","Desde:","si")
    sInput1 = ofv.GenerarInput("desde",XFSRD,"HIDDEN","","20","dt","")

    if clng(XFSRD) = 0 then
       sInput = "------------"
    else
       sInput = cdate(clng(XFSRD) + clng(cdate("01-01-1900")) )
       swd = datepart("w",sInput)
       sInput = Fechador(XFSRD)
    END IF

    Celda   =  Celda & ofv.GenCelda("","uti","","","","","",sInput & sInput1 & sinputE,"si")
    Fila    = Fila & ofv.GenRow("","","","","25","","",Celda,"si")

    Celda   = ""
    Celda   = Celda & ofv.GenCelda("","tt","","","","","","Hasta:","si")
    sInput1 = ofv.GenerarInput("hasta",XFSRH,"HIDDEN","","20","dt","")

    if clng(XFSRH) = 0 then
       sInput = "------------"
    else
       sInput = cdate(clng(XFSRH) + clng(cdate("01-01-1900")) )
       swd = datepart("w",sInput)
       sInput = Fechador(XFSRH)
    END IF

    Celda  = Celda & ofv.GenCelda("","uti","","","","","",sInput & sInput1,"si")
    Fila   = Fila & ofv.GenRow("","","","","25","","",Celda,"si")


    sInput = ofv.GenerarInput("","Aceptar","submit","","10","bt","")
    Celda  = ofv.GenCelda("","","Center colspan=2","","","","",sInput,"si")
    Fila   = Fila & ofv.GenRow("","","","","25","","",Celda,"si")

    sAccion = "COESTDLisSEG.asp?CantRegMostrar=10&CantRegMover=0&tipoC=ACCXESQ"
    Formulario = ofv.GenForm("Subconsulta","",sAccion,"","post","",Fila,"si")

    sHTML = sHTML & "<center><BR>"
    sHTML = sHTML & ofv.GenTabla("","aaa style='border-style:solid;border-width:1;border-color:#dddddd;' ","","50% ","","0","","7","0","","",Formulario,"si")
    sHTML = sHTML & "</center>"

   END IF

  IF XFNAME = "NO" OR XFNAME = "FH" THEN 

    IF XFNAME = "NO" THEN
       XFSR   = clng(date()) - clng(cdate("01-01-1900"))
       XFSRD  = XFSR
       XFSRH  = 0
       XFNAME2 = "FH"
       XRESTO = ""
       XDESC  = "DESDE:"
    ELSE
       XFSR   = XFSRH
       XFNAME2 = "OK"
       XRESTO = "&XFSRD=" & XFSRD
       XDESC  = "HASTA:"
    END IF




  set cn2 = ofv.conectar(ofv.strconn0)

'''''''''
  cniv = request.querystring("Cniv")
  cnom = request.querystring("Cnom")
  cx = request.querystring("Cx")
  cA = request.querystring("CA")

  if cniv = "" then 
     cniv = session("ENC7")
     cnom = session("ENC8")
     cx = session("ENC9")
   '  cA = session("ENCA")
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


      Celda = ofv.GenCelda("","ttt colspan=4 ","","","","","","SELECCIONAR FECHA " & XDESC,"si")
      Fila = ofv.genrow("","","","","25","","",Celda,"si")


    if clng(XFSR) = 0 then
       sInput1 = "------------"
    else
       sInput1 = cdate(clng(XFSR) + clng(cdate("01-01-1900")) )
       swd = datepart("w",sInput1)
       sInput1 = Fechador(XFSR)
    END IF


      Celda = ""
      Celda = Celda & ofv.GenCelda("","uti colspan=4 ","","","","","","ACTUAL:&nbsp;" & ucase(sInput1),"si")
     
     ' Celda = Celda & ofv.gencelda("","ut","","","","","",ucase(sInput1),"si")



    Fila = fila & ofv.genrow("","","","","","","",Celda,"si")


    'Descripcion
    CELDA = ""

    sanio = datepart("yyyy",date())
    smes = datepart("m",date())
    saniob = datepart("yyyy",date())

    SMESC = SMES

    if not isnull(XFSR) then

       if clng(XFSR) > 0 then

          sfecc = cdate(clng(XFSR) + clng(cdate("01-01-1900")) )
          smesc = datepart("m",sfecc)
          sanio = datepart("yyyy",sfecc)
       end if 


    end if

    IF CA <> "" THEN 
       smesc = request.form("smesc")
       sanioc = request.form("saniox")

       if isnull(smesc) then  SMESC = ca
       if len(smesc) = 0 then  SMESC = ca
       ca =  SMESC

       if isnull(sanioc) then  sanioc = sanio
       if len(sanioc) = 0 then  sanioc = sanio

      
    else
       sanioc = sanio

    end if  

    if smesc = 0 then smesc = smes

    session("ENCA") = smesc



    dim Aanios(9)

    for ian = 0 to 5
        Aanios(9 - ian) = saniob - ian        
    next 


    combanio = ofv.GenerarCombo3(sForm & "A", Aanios, "saniox", sanioc)



    
      Celda = ofv.GenCelda("","uTi","","5%","","","","MES:","si")

    'codigo

     Strsql = "Select * From meses Order By mes"
     sInput = ofv.GenerarCombo(sForm & "A","smesc",smesc,Strsql,cn2,1,2,"sb0")
     Celda = Celda & ofv.gencelda("","uT","","45%","","","",sInput,"si")

     SINPUT3 = ""

     IF XFNAME = "FH" THEN
        SINPUT3 = ofv.GenerarInput("fecact",XFSRD,"hidden","","2","dt","")
     END IF


     Celda = Celda & ofv.GenCelda("","uti","","5%","","","","AÑO:" & SINPUT3,"si")
     
     sinput = combanio 
     Celda = Celda & ofv.gencelda("","uT","","45%","","","",sInput,"si")

    Fila = FILA & ofv.genrow("","","","","","","",Celda,"si")


    sAccion = "COESTDSEGUI.asp?cniv=" & cniv & "&cnom=" &  cnom & "&cx=" & cx & "&ca=" & smesc & "&XFNAME=" & XFNAME & XRESTO
    Formulario = ofv.genform(sForm & "A","",sAccion,"","post","",Fila,"si")

    filas2 = ofv.GenTabla("","AAA style='border-style:solid;border-width:1;border-color:#dddddd;'","","70%","","0","","0","0","","",Formulario,"si")


    dim afech(33,1)
    dim afechs(42,1)

    CELDA = ""

    sfecds = "01-" & cstr(smesc) & "-" & cstr(sanioc)
    sfecd = clng(dateserial(sanioc,smesc,1))
    swd = datepart("w",dateserial(sanioc,smesc,1))
     
    sfechs = "28-" & cstr(smesc) & "-" & cstr(sanioc)
    sfech = clng(dateserial(sanioc,smesc,28))


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

  ''  srest = 0
    srest = clng(cdate("01-01-1900"))
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


    for ii = ii to (ii + topefecH )
        afechs(ii,0) = yy
        afechs(ii,1) = afech(yy,1)
        yy = yy + 1
    next


    for ii = ii to 42
        afechs(ii,0) = "NO"
        afechs(ii,1) = 0
    next

    CELDA = ""
    Celda = Celda & ofv.GenCelda("","ttr ","CENTER","10%","","","","D","si")
    Celda = Celda & ofv.GenCelda("","ttr ","CENTER","10%","","","","L","si")
    Celda = Celda & ofv.GenCelda("","ttr ","CENTER","10%","","","","M","si")
    Celda = Celda & ofv.GenCelda("","ttr ","CENTER","10%","","","","M","si")
    Celda = Celda & ofv.GenCelda("","ttr ","CENTER","10%","","","","J","si")
    Celda = Celda & ofv.GenCelda("","ttr ","CENTER","10%","","","","V","si")
    Celda = Celda & ofv.GenCelda("","ttr ","CENTER","10%","","","","S","si")

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
           sclase = "btr" 
        else
           sclase = "bt" 
        end if 

        if afechs(ii,0) = "NO" then
           sinput = "&nbsp;"
           sclase = "CC " 
           Celda = Celda & ofv.GenCelda("","uT","CENTER","10%","","","",sinput,"si")
        else
           sinput1 = afechs(ii,0)

           sinput1 = cstr(sinput1)

           if len(sinput1) = 1 then sinput1 = "0" & sinput1 

           sinput1 = " " & sinput1 & " " 

           sAccion = "COESTDSEGUI.asp?XFNAME=" & XFNAME2 & XRESTO

           t1 = ofv.GenerarInput("bvo1",sInput1,"submit","","10",sclase,"")
           t1 = t1 & ofv.GenerarInput("fecact",(afechs(ii,1) - srest),"hidden","","2","dt","")
           t1 = ofv.GenCelda("","uT","center","10%","","","",t1,"si")
           t1 = ofv.GenForm("feac" & ii,"",sAccion,"","post","",t1,"si") 
           Celda = Celda & t1

        end if


    next

   call ofv.cerrarconn(cn2)

'''
   filas2 = "<br>" & filas2 & "<br>" & ofv.GenTabla("","AAA style='border-style:solid;border-width:1;border-color:#dddddd;'","","30%","","0","","0","0","","",FILA,"si")

   sHTML = sHTML & filas2

  END IF

  'Botonera
  Agregar = ""
  Buscar  = ""
  Celda   = ""

  sHTML = sHTML & "</center>"

  sHTML = sHTML & ofv.botoneraESSX("0",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,celda)

  sHTML = sHTML & "</body></html>"

  call ofv.cerrarconn(cn)
end if

response.write sHTML

Function Fechador(wkfecha)


    
      fec = cdate(wkfecha + clng(cdate("01-01-1900")))
      ndia = ofv.cdia(Weekday(fec))
      nMes = ofv.cmes(Month(fec))
    
      Fechador = "  " & ndia & ", " & Day(fec) & " de " & nMes & " de " & Year(fec)


End Function

%>