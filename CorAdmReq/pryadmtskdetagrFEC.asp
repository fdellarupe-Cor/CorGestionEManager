<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/CorWKFNro.inc" -->
<%

ok=ofv.CheckUsuario()

if ok <> "Y" then
  sHTML = ok
else
  response.expires=0


 sopcion = request.querystring("opcion")




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



  sHTML = ofv.formHeader("")



  sHTML = sHTML  & "<script language=javascript> "



  sHTML = sHTML  & "function CFEC(F1,F2) { "
  sHTML = sHTML  & " document.all.FSTART.value = F1;"
  sHTML = sHTML  & " document.all.FSTARTX.innerHTML = F2;"
  sHTML = sHTML  & " } "
  sHTML = sHTML  & "</script>"

  IF SOPCION = "MI" THEN

    FSTART = REQUEST.FORM("FSTART") 
    StrSql = "UPDATE PROYTAREASTMP SET FECHASREAL = '" & FSTART & "' WHERE codigo = '" & session("NCliente") & "' AND TAREA = '" & SESSION("TAREA") & "' AND VPLAN = " & SESSION("VPLAN")
    CALL ofv.crearconsultaEx(StrSql,cn,1,parametros)

    response.redirect "pryadmtskdetagr.asp"

  END IF






  StrSql = "Select * From PROYTAREASTMP where codigo = '" & session("NCliente") & "' AND TAREA = '" & SESSION("TAREA") & "'  AND VPLAN = " & SESSION("VPLAN")
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)



    sHTML = sHTML & "<center>"

    sImagen = "<FONT FACE=WINGDINGS SIZE=4>4</FONT>"
    sImagenA = "<FONT FACE=WINGDINGS COLOR=#AA0000 SIZE=2>n</FONT>"
    sImagenN = "<FONT FACE=WINGDINGS COLOR=#88bbdd SIZE=2>n</FONT>"
    sImageS = "<B><FONT FACE=WEBDINGS COLOR=#993333 SIZE=3>4</FONT></B>"




      celdas = ofv.GenCelda("","TTT colspan=3","","","","","",sImagen & "&nbsp;<B>FECHA DE INICIO DE LA TAREA</B>","si")
      filas  = ofv.GenRow("","","","","10","","",celdas,"si")

      celdas = ofv.GenCelda("","AA COLSPAN=3","CENTER","","","","","&nbsp;","si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

      celdas = ofv.GenCelda("","vt  ","","","","","","TAREA:","si")
      celdas = celdas & ofv.GenCelda("","TTT COLSPAN=2 ","","","","","",rs(4) & "." & RS(5),"si")
    ''''  celdas = celdas & ofv.GenCelda("","TTT","","","","","",RS(5),"si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")







      XMSG = "NO"

         Xhoy = CLNG(DATE()) - clng(cdate("01-01-1900"))

         XFSR = CLNG("0" & rs(9))

         sclase = "CC" 

         if clng(XFSR) = 0 then
            sInput1 = "------------"
         else
            sInput1 = cdate(clng(XFSR) + clng(cdate("01-01-1900")) )
            swd = datepart("w",sInput1)
            sInput1 = Fechador(XFSR)
         END IF




      celdas = ofv.GenCelda("","UT COLSPAN=3 ","","","","",""," ","si")

      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


'''''''''''''''''''''''''''


     Celda = ofv.GenCelda("","TTR COLSPAN=2","","","","","","FECHA INICIO:","si")
     Fila = ofv.genrow("","","","","","","",Celda,"si")

     Celda = ""

         sInput2 = ofv.GenerarInput("FSTART",XFSR,"hidden","Usuarios","10","dtn","")
         sInput2 = sInput2 & ofv.GenerarInput("btn1","ACEPTAR","SUBMIT","Usuarios","10","BT","")

         if bModificar then
            sAccion = "pryadmtskdetagrFEC.asp?OPCION=MI"
            formulario0 = ofv.GenForm("Usuarios","",sAccion,"","post","",sInput2,"si")
         else
            formulario0 = ofv.GenForm("","","","","post","",filas,"si")
         end if
     Celda = Celda & ofv.gencelda("FSTARTX","VT","","80%","","","",sInput1,"si")

     Celda = Celda & ofv.gencelda("","VT","","20%","","","",formulario0,"si")



     Fila = FILA & ofv.genrow("","","","","","","",Celda,"si")


    'Descripcion
    CELDA = ""

    sanio = datepart("yyyy",date())
    smes = datepart("m",date())

    SMESC = SMES

   '' if not isnull(XFSR) then
   ''    if clng(XFSR) > 0 then
   ''       smesc = cdate(clng(XFSR) + clng(cdate("01-01-1900")) )
   ''       smesc = datepart("m",smesc)
   ''    end if 
   '' end if

    IF CA <> "" THEN 
       smesc = request.form("smesc")

       if isnull(smesc) then  SMESC = ca
       if len(smesc) = 0 then  SMESC = ca
       ca =  SMESC
      
    else
       ca = smesc

    end if  

    if smesc = 0 then smesc = smes

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
     sInput = ofv.GenerarCombo("document." & sForm & "A","smesc",smesc,Strsql,cn2,1,2,"sb0")
     Celda = Celda & ofv.gencelda("","VT","","45%","","","",sInput,"si")

     Celda = Celda & ofv.GenCelda("","VT","","15%","","","","AÑO:","si")
     
     sinput = sanioc 
     Celda = Celda & ofv.gencelda("","VT","","20%","","","",sInput,"si")



    Formulario = Fila

    filas2 = ofv.GenTabla("","AAA ","","50%","","0","","0","0","","",Formulario,"si")





    dim afech(50,1)
    dim afechs(52,1)

    CELDA = ""

   ' sfecds = "01-" & cstr(smesc) & "-" & cstr(sanioc)
   ' sfecd = clng(dateserial(sanioc,smesc,1))

    sfecd = xhoy


    sfecd = sfecd - 40

    sInput1 = cdate(clng(sfecd))

    sanio1 = datepart("yyyy",sInput1)
    smes1 = datepart("m",sInput1)
    sdia1 = datepart("d",sInput1)

    swd = datepart("w",dateserial(sanio1,smes1,sdia1))

   ''' response.write "f " & sanio1 & "-" & smes1 & "-" & sdia1 & " d " & swd
     
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

  '  srest = 0
    srest = clng(cdate("01-01-1900"))
  '  topefech = sfech - sfecd 


    topefech = 42 


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
        sInput1      = cdate(clng(afech(yy,1)) )
        sdia1        = datepart("d",sInput1)
        afechs(ii,0) = sdia1
        afechs(ii,1) = afech(yy,1)
        yy = yy + 1
    next


    for ii = ii to 52
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

    xcolor = "background-color:#f7aa22;"

    for ii = 1 to 53

        if yy = 7 then
           yy = 1
           Fila = fila & ofv.genrow("","","","","","","",Celda,"si")
           CELDA = ""
        else
           yy = yy + 1 
        end if

        if ii > 52 then exit for  

        if afechs(ii,0) = "1" then xcolor = ""

        if yy = 1 or yy = 7 then
           sclase = "tt3 style='border-style:outset;border-width:2;'" 
        else
           sclase = "TT style='border-style:outset;border-width:2;" & xcolor & "'" 
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

           sAccion = " "

           sf = "CFEC('" & (afechs(ii,1) - srest) & "','"  & Fechador((afechs(ii,1) - srest)) & "');"

           t1 = ofv.GenerarInput("bvo1",sInput1,"BUTTON","","10",sclase,sf)
     '''      t1 = t1 & ofv.GenerarInput(rs(5).name,(afechs(ii,1) - srest),"hidden","","2","dt","")
           t1 = ofv.GenCelda("","VT","center","10%","","","",t1,"si")
      '''     t1 = ofv.GenForm("feac" & ii,"",sAccion,"","post","",t1,"si") 
           Celda = Celda & t1

        end if


    next



   filas2 = "<br>" & filas2 & "<br>" & ofv.GenTabla("","AAA style='border-style:SOLID;border-width:1;border-color:#888888;'","","50%","","0","","0","0","","",FILA,"si")



   filas2 = filas2 & "<br>" 

      celdas = ofv.GenCelda("","vt COLSPAN=3 style='border-style:SOLID;border-width:1;border-color:#888888;' ","CENTER","","","","",filas2,"si")

      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")





    sHTML = sHTML & ofv.GenTabla("","","","70%","","0","","0","0","","",filas,"si")
    sHTML = sHTML & "</center><br><br>"






  Agregar = ""
  Buscar  = ""

  sAccion = "pryadmtskdetAGR.asp"
  Celda = ofv.BotonVolver(sAccion)


  sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","no",Agregar,Buscar,Celda)



  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)


  sHTML = sHTML  & "</body></html>"
end if

response.write sHTML

FUNCTION GETPDV(XPDV)

 XPDI = 0

 XTA = CINT(XPDV / 10)

 FOR XPDI=1 TO XTA
     GETPDV = GETPDV & SIMAGENA
 NEXT


 FOR XPDI=XTA + 1  TO 10
     GETPDV = GETPDV & SIMAGENN
 NEXT

END FUNCTION


Function Fechador(wkfecha)


    
      fec = cdate(wkfecha + clng(cdate("01-01-1900")))
      ndia = ofv.cdia(Weekday(fec))
      nMes = ofv.cmes(Month(fec))
    
      Fechador = "  " & ndia & ", " & Day(fec) & " de " & nMes & " de " & Year(fec)


End Function



%>