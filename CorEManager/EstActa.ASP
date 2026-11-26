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
  IF XFNAME = "" THEN XFNAME = "NO"

  XFSRD = request.querystring("XFSRD")
  IF XFSRD = "" THEN XFSRD = 0

  XFSRH = request.querystring("XFSRH")
  IF XFSRH = "" THEN XFSRH = 0

  IF XFNAME = "FH" THEN
     XFSRD = request.form("FECACT")
     XFSRH = XFSRD + 1
  ELSEIF XFNAME = "OK" THEN
     XFSRH = request.form("FECACT")
  END IF

session("XIN")=""

'response.write request.querystring

  set cn = ofv.conectar(ofv.strconn2)

XXXX = NEWEMPRE("NO")

      session("ASP")="SUBCONSULTA.ASP"

      sUN = request.querystring("EMPRE")
    'response.write "sEMPRE: " & sEMPRE
       IF sUN = ""  AND SESSION("UN") = "" THEN

         sUN = ATREMP(0, 0)
         SESSION("UN") = sUN & ""
         If (SESSION("UN") & "") = "" Then SESSION("UN") = 1
       elseIF sUN <> "" THEN
        XXXX = NEWEMPRE(sUN)
        SESSION("UN")=sUN
       ELSE
        XXXX = NEWEMPRE(SESSION("UN"))
        sUN=SESSION("UN")
   end if


  sHTML = ofv.FormHeader( session("FormN"))
  sHTML = sHTML & "<center>"
'response.write session("UN")
 TablaH = ofv.GenTabla("","","","20%","","0","0","0","25","","",selecEmpre(),"si")


shtml = shtml & tablah

  tt="&nbsp;"

  IF XFNAME = "OK" THEN

    Celda = ofv.GenCelda("","","","40%","","","","&nbsp;","si")
    Celda = Celda & ofv.GenCelda("","tt","","10%","","","","Desde:","si")
    sInput1 = ofv.GenerarInput("desde",XFSRD,"HIDDEN","","20","dt","")

    if clng(XFSRD) = 0 then
       sInput = "------------"
    else
       sInput = cdate(clng(XFSRD) + clng(cdate("01-01-1900")) )
       swd = datepart("w",sInput)
       sInput = Fechador(XFSRD)
    END IF


    Celda = Celda & ofv.GenCelda("","uti","","","","","",sInput & sInput1,"si")
    Fila = ofv.GenRow("","","","","25","","",Celda,"si")

    Celda = ofv.GenCelda("","abc colspan=3","","","","","","&nbsp;","si")
    Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")

    Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
    Celda = Celda & ofv.GenCelda("","tt","","","","","","Hasta:","si")
    sInput1 = ofv.GenerarInput("hasta",XFSRH,"HIDDEN","","20","dt","")

    if clng(XFSRH) = 0 then
       sInput = "------------"
    else
       sInput = cdate(clng(XFSRH) + clng(cdate("01-01-1900")) )
       swd = datepart("w",sInput)
       sInput = Fechador(XFSRH)
    END IF

    Celda = Celda & ofv.GenCelda("","uti","","","","","",sInput & sInput1,"si")
    Fila = Fila & ofv.GenRow("","","","","25","","",Celda,"si")

    Celda = ofv.GenCelda("","abc colspan=3","","","","","","&nbsp;","si")
    Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")


'   sInput = "<select name=fecha><option value=todos>TODOS</option>"
'   sInput = sInput & "<option value=sinact>SIN ACTA INFORMADA</option>"
'   sInput = sInput & "</select><br><br>"

    scomK = "todos" & chr(9) & "SINACT"
    scomD = "Todos" & chr(9) & "SIN ACTA INFORMADA"


    sInput = ofv.gencomboFX(sForm,"selact","todos",scomK,scomD,"no")

    Celda = ofv.GenCelda("","","","40%","","","","&nbsp;","si")
    Celda = Celda & ofv.GenCelda("","tt","","10%","","","","Seleccion:","si")

    Celda = Celda & ofv.GenCelda("","uti","","","","","",sInput,"si")
    Fila = Fila & ofv.GenRow("","","","","25","","",Celda,"si")

    Celda = ofv.GenCelda("","abc colspan=3","","","","","","&nbsp;","si")
    Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")
    '''
    Celda = ofv.GenCelda("","","","40%","","","","&nbsp;","si")
	Celda = Celda & ofv.GenCelda("","tt","","10%","","","","Estado:","si")

	sql = "select * from estados order by estado"
	Set rs = ofv.crearconsultaEx(sql,cn,1,parametros)

	if not rs.eof then estados = "TODOS"

	scomK = "TODOS"
	scomD = "TODOS"
	scomSEP = chr(9)

	do while not rs.eof

	scomK = scomK & scomSEP & rs(0)
	scomD = scomD & scomSEP & rs(1)

	rs.movenext
	loop

	sInput = ofv.gencomboFX(sForm,"estado",estados,scomK,scomD,"no")

	Celda = Celda & ofv.GenCelda("","uti","","","","","",sInput,"si")
	Fila = Fila & ofv.GenRow("","","","","25","","",Celda,"si")

	Celda = ofv.GenCelda("","abc colspan=3","","","","","","&nbsp;","si")
    Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")
''

    sInput = ofv.GenerarInput("","Aceptar","submit","","10","bt","")
    Celda = ofv.GenCelda("","","Center colspan=3","","","","",sInput,"si")
    Fila = Fila & ofv.GenRow("","","","","25","","",Celda,"si")

    sAccion = "EstActaLis.asp?CantRegMostrar=10&CantRegMover=0"
    Formulario = ofv.GenForm("Subconsulta","",sAccion,"","post","",Fila,"si")

    sHTML = sHTML & ofv.GenTabla("","","","100% ","","0","","0","0","","",Formulario,"si")

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


      Celda = ofv.GenCelda("","VT colspan=2 ","","","","","","SELECCIONAR FECHA DE VIGENCIA " & XDESC,"si")
      Celda = Celda & ofv.GenCelda("","VT","","","","","","ACTUAL:","si")

    if clng(XFSR) = 0 then
       sInput1 = "------------"
    else
       sInput1 = cdate(clng(XFSR) + clng(cdate("01-01-1900")) )
       swd = datepart("w",sInput1)
       sInput1 = Fechador(XFSR)
    END IF

      Celda = Celda & ofv.gencelda("","VT","","","","","",sInput1 ,"si")










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
  '  RESPONSE.WRITE sanio
    if smesc = 0 then smesc = smes

    session("ENCA") = smesc

    dim Aanios(20)

    for ian = 0 to 20
        if XFNAME<> "FH" THEN
          Aanios(20 - ian) = ( sanio) - ian
        ELSE
          Aanios(20 - ian) = (20 + sanio) - ian
        END IF
    next


    combanio = ofv.GenerarCombo3(sForm & "A", Aanios, "saniox", sanioc)




      Celda = ofv.GenCelda("","VT","","5%","","","","MES:","si")

    'codigo

     Strsql = "Select * From meses Order By mes"
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


    sAccion = "ESTACTA.asp?cniv=" & cniv & "&cnom=" &  cnom & "&cx=" & cx & "&ca=" & smesc & "&XFNAME=" & XFNAME & XRESTO
    Formulario = ofv.genform(sForm & "A","",sAccion,"","post","",Fila,"si")

    filas2 = ofv.GenTabla("","AAA style='margin-bottom:20px;'","","100%","","0","","0","0","","",Formulario,"si")


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
    Celda = Celda & ofv.GenCelda("","Vt style='border-style:outset;border-width:2;color:" & Session("CMFPNBOTFT") & ";background-color:" & Session("CMFPNTIT") & "'","CENTER","10%","","","","D","si")
    Celda = Celda & ofv.GenCelda("","Vt style='border-style:outset;border-width:2;color:" & Session("CMFPNBOTFT") & ";background-color:" & Session("CMFPNTIT") & "'","CENTER","10%","","","","L","si")
    Celda = Celda & ofv.GenCelda("","Vt style='border-style:outset;border-width:2;color:" & Session("CMFPNBOTFT") & ";background-color:" & Session("CMFPNTIT") & "'","CENTER","10%","","","","M","si")
    Celda = Celda & ofv.GenCelda("","Vt style='border-style:outset;border-width:2;color:" & Session("CMFPNBOTFT") & ";background-color:" & Session("CMFPNTIT") & "'","CENTER","10%","","","","M","si")
    Celda = Celda & ofv.GenCelda("","Vt style='border-style:outset;border-width:2;color:" & Session("CMFPNBOTFT") & ";background-color:" & Session("CMFPNTIT") & "'","CENTER","10%","","","","J","si")
    Celda = Celda & ofv.GenCelda("","Vt style='border-style:outset;border-width:2;color:" & Session("CMFPNBOTFT") & ";background-color:" & Session("CMFPNTIT") & "'","CENTER","10%","","","","V","si")
    Celda = Celda & ofv.GenCelda("","Vt style='border-style:outset;border-width:2;color:" & Session("CMFPNBOTFT") & ";background-color:" & Session("CMFPNTIT") & "'","CENTER","10%","","","","S","si")

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

           sAccion = "ESTACTA.asp?XFNAME=" & XFNAME2 & XRESTO

           t1 = ofv.GenerarInput("bvo1",sInput1,"submit","","10",sclase,"")
           t1 = t1 & ofv.GenerarInput("fecact",(afechs(ii,1) - srest),"hidden","","2","dt","")
           t1 = ofv.GenCelda("","VT","center","10%","","","",t1,"si")
           t1 = ofv.GenForm("feac" & ii,"",sAccion,"","post","",t1,"si")
           Celda = Celda & t1

        end if


    next

   call ofv.cerrarconn(cn2)

'''
   filas2 = "<br>" & filas2 & ofv.GenTabla("","AAA","","50%","","0","","0","0","","",FILA,"si")

   sHTML = sHTML & filas2

  END IF


  sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","","","","")


  sHTML = sHTML & "</center></body></html>"

  call ofv.cerrarconn(cn)
end if

function selecEmpre()

sImagenx = "<B><FONT FACE=CorFlatIcon COLOR=#aa0000 SIZE=4>x</FONT></B>"

   IF  Session("EMPMUN") = "S" THEN

   Fila    = ""

   SINPUT3 = "&nbsp;<A class=tt style='color:#000000;' title='SELECCIONE EMPRESA' HREF='javascript:void(0);' onclick='shwDem();' >EMPRESA:&nbsp;&nbsp;</a>"

   Fila    = ""
   Celda   = ""
   sHref   = ""

   sImage  = ofv.GenImage("","","","","..\CORESS20\EMPRESA\" & ATREMP(0,5),"0","")
   sLink   = ofv.GenLink("","ttt",sHref,"ftemas","","",sImage,"si")
   Celda   = Celda & ofv.GenCelda("","tt ","","8%","","","",SINPUT3,"si")
   Celda   = Celda & ofv.GenCelda("","ttt ","center","5%","","","",sLink,"si")
   Celda   = Celda & ofv.GenCelda("","ttt  ","","","","","","&nbsp;<BIG>" & ATREMP(0,4) & "</BIG>","si")

   Fila    = fila & ofv.GenRow("","","top","","","","",Celda,"si")

   Tabla   = ofv.GenTabla("","aaa style='"  & Session("CBRDABL") & "' ","","100% ","","0","","-1","-2","","",Fila,"si")





   shtml=shtml & "<script language=javascript>" & chr(13)
   shtml=shtml & "function shwDem(){" & chr(13)
   shtml=shtml & "document.all.DEM1.style.display='block';" & chr(13)
   shtml=shtml & "}" & chr(13)
   shtml=shtml & "</script>" & chr(13)
 '----------------------------------------------
   sinput2 = "<a  style='text-decoration:none;' onmouseover='this.style.textDecoration=" & chr(34)
   sinput2 = sinput2  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
   sinput2 = sinput2  & "none" & chr(34) & ";' href='javascript:void(0);' ONCLICK=" & chr(34) & "document.all.DEM1.style.display='none';" & chr(34)
   sinput2 = sinput2  & " TITLE='OCULTAR' ><b>" & sImagenx & "</b></a>"
   TABLA = Tabla & "<DIV id='DEM1' STYLE='display:none;z-index:1;margin-top:1;margin-left:1;background-color:#ffffff;position:absolute;left:1px;top:240px;left:20px;height:250px;width:350px;border-style:inset;border-color:#f7aa22;border-width:2;overflow:auto;'>"


   TABLAX = TablaX & OFV.gentabla("", "aaa style='filter:shadow(color=#dddddd,direction=125);' ", "", "100%", "", "0", "", "-2", "-1", "", "", "", "no")
   TABLAX = TablaX & "<tr><td  "
   TABLAX = TablaX & "class=vT align=left colspan=3 STYLE='border-style:outset;border-width:2;background-color:#eeeeee;' >" & sinput2 & "&nbsp;SELECCIONE EMPRESA</td></tr>"
   TABLAX = TablaX & "<tr><td colspan=3 "
   TABLAX = TablaX & " >&nbsp;</td></tr>"

   FOR IAEMP=1 TO TAEMP

    '''  RESPONSE.WRITE  ATREMP(IAEMP,2)

      IF ATREMP(IAEMP,2) = 1 THEN

       IF ATREMP(IAEMP,0) = ATREMP(0,0) THEN
          TABLAX = TablaX & "<tr BGCOLOR=#EEEEEE >"
       ELSE
          TABLAX = TablaX & "<tr>"
       END IF

       TABLAX = TablaX & "<td width=5 "
       IF ATREMP(IAEMP,0) = ATREMP(0,0) THEN
          TABLAX = TablaX & "class=vt  >" & sImagen4si & "</td>"
       ELSE
          TABLAX = TablaX & "class=vt  >" & sImagenI3 & "</td>"
       END IF

       sImage = ofv.GenImage("","","","","..\CORESS20\EMPRESA\" & ATREMP(IAEMP,5),"0","")
       TABLAX = TablaX & "<td width=5 "
       TABLAX = TablaX & "class=vt  >" & sImage & "</td>"

       TABLAX = TablaX & "<td class=vt  >"

       'sHref = "subconsulta.asp?Menu=" & sMenu & "&Nombre=" & sNombren & "&CantRegMostrar=10&CantRegMover=0&EMPRE=" & ATREMP(IAEMP,0)
 		sHref = "EstActa.asp?EMPRE=" & ATREMP(IAEMP,0)
     '''  sHref = "OBJtreeTema.asp?Tipo=" & sTipo & "&Nombre=" & LNombre & "&Menu=" & SESSION("MENU") & "&EMPRE=" & ATREMP(IAEMP,0)
       sLink = ofv.GenLink("","cc",sHref,"","","","<b>" & ATREMP(IAEMP,4) & "</b>","si")
       TABLAX = TablaX & sLink
       TABLAX = TablaX & "</td>"
       TABLAX = TablaX & "</tr>"
      END IF

   NEXT

   TABLAX = TablaX & OFV.genFtabla()

   TABLA = Tabla & TABLAX

   TABLA = Tabla & "</DIV>"

   sHTML = sHTML & TABLA

  end if

end function

response.write sHTML

Function Fechador(wkfecha)



      fec = cdate(wkfecha + clng(cdate("01-01-1900")))
      ndia = ofv.cdia(Weekday(fec))
      nMes = ofv.cmes(Month(fec))

      Fechador = "  " & ndia & ", " & Day(fec) & " de " & nMes & " de " & Year(fec)


End Function

%>