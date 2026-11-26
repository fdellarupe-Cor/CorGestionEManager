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
'response.write request.form
  Server.ScriptTimeout=60 * 10
  XFNAME = request.querystring("XFNAME")
  IF XFNAME = "" THEN XFNAME = "NO"

  XFSRD = request.querystring("XFSRD")
  IF XFSRD = "" THEN XFSRD = 0

  XFSRH = request.querystring("XFSRH")
  IF XFSRH = "" THEN XFSRH = 0

  IF XFNAME = "OK" THEN
     XFSRD = request.form("FECACT")
  END IF

  IF XFNAME = "ACT" THEN
     XFSRD = request.form("FACTA")
     XFNRO = request.form("NACTA")
  END IF

    selact=Request.querystring("selact")
    sorden = Request.querystring("sorden")
	estado = request.querystring("estado")
    Sfiltro = Request.querystring("sfiltro")

    CantRegMostrar = CDbl(Request.querystring("CantRegMostrar"))
    CantRegMover = CDbl(Request.querystring("CantRegMover"))

    sID = Request.querystring("Desde")
    sID2 = Request.querystring("Hasta")
    sID3 = Request.querystring("Fecha")
    Addqry = "&Desde=" & sID & "&Hasta=" & sID2 & "&Fecha=" & sID3

    If Len(sorden) = 0 Then
      sorden = "N"
    End If

    If Len(Sfiltro) = 0 Then
      Sfiltro = "TODOS"
    End If

    Sfiltro = Replace(Sfiltro, "_", " ")



  set cn = ofv.conectar(ofv.strconn2)




  sHTML = ofv.FormHeader( session("FormN"))
  sHTML = sHTML & "<center>"

  tt="&nbsp;"

  IF XFNAME = "ACT" THEN



        sverestado = " and objnivel >= " & Session("SegNivel") & " And estcod In (9,3)"
        XIN = session("XIN")

        cdesde = clng(sID) - 1
        chasta = clng(sID2) + 1

        xinc = 0


        strsql = "select * from csporfechas "
        IF  selact <> "TODOS" THEN
        strsql = strsql & " WHERE (fecvigen > " & cdesde & " and fecvigen < " & chasta & ") AND NROACTA IS  NULL"
        ELSE
        strsql = strsql & " WHERE (fecvigen > " & cdesde & " and fecvigen < " & chasta & ") "
        END IF
        strsql = strsql & sverestado
        strsql = strsql & xin

        response.write "sql: " & strsql & "<br>"
        Set rs = OFV.crearconsultaEx(strsql, cn, 1, "")
        do until rs.eof
           strsql = "UPDATE OBJPRTY SET FECACTA = " & XFSRD & ", NROACTA = '" & XFNRO & "' WHERE CODOBJETO = " & rs(3)
         '  response.write "sql: " & strsql & "<br>"
           call OFV.crearconsultaEx(strsql, cn, 1, "")
           xinc = xinc + 1

        rs.movenext
        loop
       '' response.write " esquemas actualizados: " & xinc & "<br>"

        response.redirect "EstActaLisConf.asp?Desde=" & sID & "&Hasta=" & sID2 & "&Fecha=" & sID3 & "&sorden=" & sorden & "&CantRegMostrar=10&CantRegMover=0&XFNAME=ACT&FACTA=" & XFSRD & "&NACTA=" & XFNRO & "&selact=TODOS&estado=" & estado

   END IF

  IF XFNAME = "OK" THEN

    Fila = ""

    sImagenE = "<B><FONT FACE=corflaticon COLOR=#F7AA22 SIZE=4>&#0167;</FONT></B>"

    sAccion = "ESTACTACONFIRMAR.asp?XFNAME=NO&Desde=" & sID & "&Hasta=" & sID2 & "&Fecha=" & sID3 & "&sorden=" & sorden & "&CantRegMostrar=10&CantRegMover=0&selact=" & selact & "&estado=" & estado

    sinputE = "&nbsp;&nbsp;<a  class=UT href=" & sAccion
    sinputE = sinputE  & " TITLE='MODIFICAR FECHA' >" & sImagenE & "</a>"

    Celda = ofv.GenCelda("","abc colspan=3","","","","","","&nbsp;","si")
    Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")
    Celda = ofv.GenCelda("","VT colspan=3 ","CENTER","","","","","ASIGNACION DE FECHA Y NRO. DE ACTA","si")
    Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")
    Celda = ofv.GenCelda("","abc colspan=3 STYLE='BORDER-TOP-STYLE:RIDGE;BORDER-TOP-WIDTH:2;BORDER-TOP-COLOR:#F7AA22;' ","","","","","","&nbsp;","si")
    Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")
    Celda = ofv.GenCelda("","abc colspan=3","","","","","","&nbsp;","si")
    Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")

    Celda = ofv.GenCelda("","","","20%","","","","&nbsp;","si")
    Celda = Celda & ofv.GenCelda("","tt","","25%","","","","&nbsp;FECHA ACTA:&nbsp;&nbsp;","si")
    sInput1 = ofv.GenerarInput("FACTA",XFSRD,"HIDDEN","","20","dt","")

    if clng(XFSRD) = 0 then
       sInput = "------------"
    else
       sInput = cdate(clng(XFSRD) + clng(cdate("01-01-1900")) )
       swd = datepart("w",sInput)
       sInput = Fechador(XFSRD)
       sInput = "<font size=2>" & sinput & "</font>"
    END IF



    Celda = Celda & ofv.GenCelda("","VT","","","","","",sInput & sInput1 & sinputE,"si")
    Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")
    Celda = ofv.GenCelda("","abc colspan=3","","","","","","&nbsp;","si")
    Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")

    Celda = ofv.GenCelda("","","","20%","","","","&nbsp;","si")
    Celda = Celda & ofv.GenCelda("","tt","","25%","","","","&nbsp;NUMERO ACTA:&nbsp;&nbsp;","si")
    sInput = ofv.GenerarInput("NACTA","","AGREGAR","","40","dt","")
    Celda = Celda & ofv.GenCelda("","","","","","","",sInput,"si")
    Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")

    Celda = ofv.GenCelda("","abc colspan=3","","","","","","&nbsp;","si")
    Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")
    Celda = ofv.GenCelda("","abc colspan=3","","","","","","&nbsp;","si")
    Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")

    sInput = ofv.GenerarInput("","Aceptar","submit","","10","bt","")
    Celda = ofv.GenCelda("","","Center colspan=3","","","","",sInput,"si")
    Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")

  sFuncion = "onsubmit=" & chr(34) & "if (this.NACTA"
  sFuncion = sFuncion & ".value.length==0 ) "
  sFuncion = sFuncion & "{alert('Numero de acta no debe ser nulo!!!');this.NACTA.focus();return false;}" & chr(34)

    sAccion = "EstActaCONFIRMAR.asp?Desde=" & sID & "&Hasta=" & sID2 & "&Fecha=" & sID3 & "&sorden=" & sorden & "&CantRegMostrar=10&CantRegMover=0&XFNAME=ACT" & "&selact=" & selact & "&estado=" & estado
    Formulario = ofv.GenForm("Subconsulta","",sAccion,"","post",sFuncion,Fila,"si")

    sHTML = sHTML & "<CENTER><BR><BR><BR><BR>" & ofv.GenTabla("","AAA STYLE='BORDER-STYLE:RIDGE;BORDER-WIDTH:2;BORDER-COLOR:#F7AA22;' ","","70% ","","0","","0","0","","",Formulario,"si") & "</CENTER>"

   END IF

  IF XFNAME = "NO"  THEN

    IF XFNAME = "NO" THEN
       XFSR   = clng(date()) - clng(cdate("01-01-1900"))
       XFSRD  = XFSR
       XFSRH  = 0
       XFNAME2 = "OK"
       XRESTO = ""
       XDESC  = "ACTA:"

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


      Celda = ofv.GenCelda("","VT colspan=2 ","","","","","","SELECCIONAR FECHA " & XDESC,"si")
      Celda = Celda & ofv.GenCelda("","VT","","","","","","ACTUAL:","si")

    if clng(XFSR) = 0 then
       sInput1 = "------------"
    else
       sInput1 = cdate(clng(XFSR) + clng(cdate("01-01-1900")) )
       swd = datepart("w",sInput1)
       sInput1 = Fechador(XFSR)
    END IF

      Celda = Celda & ofv.gencelda("","VT","","","","","",sInput1,"si")










    Fila = ofv.genrow("","","","","","","",Celda,"si")

  ''  filas = filas & ofv.GenTabla("","","","50%","","0","","0","0","","",Fila,"si")

    'Descripcion
    CELDA = ""

    sanio = datepart("yyyy",date())
    smes = datepart("m",date())

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



 '   if cint(smesc) = cint(smes) then
 '      sanioc = sanio
  '  elseif cint(smesc) < cint(smes) then
  '     sanioc = sanio + 1
  '  else
  '     sanioc = sanio
  '  end if


    dim Aanios(20)

    for ian = 0 to 20
        Aanios(20 - ian) = ( sanio) - ian
    next


   ' combanio = ofv.GenerarCombo3("document." & sForm & "A", Aanios, "saniox", sanioc)

    combanio = ofv.GenerarCombo3( sForm & "A", Aanios, "saniox", sanioc)



      Celda = ofv.GenCelda("","VT","","5%","","","","MES:","si")

    'codigo

     Strsql = "Select * From meses Order By mes"
     'sInput = ofv.GenerarCombo("document." & sForm & "A","smesc",smesc,Strsql,cn2,1,2,"sb0")
     sInput = ofv.GenerarCombo(sForm & "A","smesc",smesc,Strsql,cn2,1,2,"sb0")
     Celda = Celda & ofv.gencelda("","VT","","45%","","","",sInput,"si")

     SINPUT3 = ""

     IF XFNAME = "FH" THEN
        SINPUT3 = ofv.GenerarInput("fecact",XFSRD,"hidden","","2","dt","")
     END IF


     Celda = Celda & ofv.GenCelda("","VT","","5%","","","","AÑO:" & SINPUT3,"si")

     sinput = combanio
     Celda = Celda & ofv.gencelda("","VT","","45%","","","",sInput,"si")

    Fila = FILA & ofv.genrow("","","","","","","",Celda,"si")


    sAccion = "ESTACTACONFIRMAR.asp?cniv=" & cniv & "&cnom=" &  cnom & "&cx=" & cx & "&ca=" & smesc & "&Desde=" & sID & "&Hasta=" & sID2 & "&Fecha=" & sID3 & "&sorden=" & sorden & "&CantRegMostrar=10&CantRegMover=0&XFNAME=" & XFNAME & XRESTO &  "&selact=" & selact & "&estado=" & estado
    formulario = ofv.genform(sForm & "A","",sAccion,"","post","",Fila,"si")

    filas2 = ofv.GenTabla("","AAA style='border-style:ridge;border-width:2;border-color:#f7aa22;'","","100%","","0","","0","0","","",Formulario,"si")


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

           sAccion = "ESTACTACONFIRMAR.asp?Desde=" & sID & "&Hasta=" & sID2 & "&Fecha=" & sID3 & "&sorden=" & sorden & "&CantRegMostrar=10&CantRegMover=0&XFNAME=" & XFNAME2 & XRESTO & "&selact=" & selact & "&estado=" & estado

           t1 = ofv.GenerarInput("bvo1",sInput1,"submit","","10",sclase,"")
           t1 = t1 & ofv.GenerarInput("fecact",(afechs(ii,1) - srest),"hidden","","2","dt","")
           t1 = ofv.GenCelda("","VT","center","10%","","","",t1,"si")
           t1 = ofv.GenForm("feac" & ii,"",sAccion,"","post","",t1,"si")
           Celda = Celda & t1

        end if


    next

   call ofv.cerrarconn(cn2)

'''
   filas2 = "<br>" & filas2 & "<br>" & ofv.GenTabla("","AAA style='border-style:ridge;border-width:2;border-color:#f7aa22;'","","50%","","0","","0","0","","",FILA,"si")

   sHTML = sHTML & filas2
  END IF
   sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","","","","")
  sHTML = sHTML & "</center></body></html>"

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