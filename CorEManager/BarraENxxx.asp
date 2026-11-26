<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0
  bModificar = False
  ofv.ObtenerAtributos Session("Form2"),""
  bModificar = ofv.AModificar
  if session("Nomodificar") = "S" then bModificar = False

  sTipo    = request.querystring("Tipo")
  sCodigo  = request.querystring("Codigo")
  sPerfil  = request.querystring("Perfil")
  sFuncion = request.querystring("Funcion")
  sVolver  = request.querystring("Volver")
  sSClick  = request.querystring("SSClick")
  sNIV1    = request.querystring("NIV1")
  sNIV2    = request.querystring("NIV2")
  sNIV3    = request.querystring("NIV3")
  sNIV4    = request.querystring("NIV4")
  sNIV5    = request.querystring("NIV5")
  sMAINAPL = request.querystring("MAINAPL")


 '' response.write "recibe " & request.querystring & "<br>"

  set cn = ofv.conectar(ofv.strconn2)
  scodi = sCodigo
  stipi = sTipo
  srestric = "N"

  session("NCliente") = sCodigo
  session("TCliente") = sTipo
  SESSION("CODVINOBJ") = ""

  XXXX = GENOBJALOG()



  if sFuncion = "E" then
     XXXX = GETOBJALOGUPD("ED")
  else
     XXXX = GETOBJALOGUPD("CON")
  end if


'xxxxxxxxxxxxxxxxxx


    redim preserve sPath(0)
    i = 1

 IF SMAINAPL = "OR" OR SMAINAPL = "PL" THEN

 ELSE

    do

       strsql = "select * from csbarr where codojasoc = " & sCodigo & " and tipobjasoc = '" & sTipo & "'"



      set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
      if not rs.eof then
        redim preserve sPath(i)
        sCodigo = rs(1)
        sTipo = rs(2)
        if not isnull(isnull(rs(0))) then
          sPath(i) = rs(0)
        else
          i = i + 1
          exit do
        end if
        i = i + 1
        call ofv.cerrarconsulta(rs)
      else
        call ofv.cerrarconsulta(rs)
        exit do
      end if

    loop




    strsql = "SELECT PRODSERV.descri, TEMA.descri FROM TEMA, RELAOBJ, PRODSERV where "
    strsql = strsql & " RELAOBJ.codobjmaster = PRODSERV.codprodserv and "
    strsql = strsql & " RELAOBJ.tipobjmaster = PRODSERV.marcaps and "
    strsql = strsql & "TEMA.codtema = PRODSERV.codtema and codprodserv = " & sCodigo
    set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    if not rs.eof then
      redim preserve sPath(i + 2)
      sPath(i + 1) = rs(0)
      sPath(i + 2) = rs(1)

    end if
    call ofv.cerrarconsulta(rs)

  if sFuncion <> "E" then

    strsql = "SELECT * FROM PRODSERV where codprodserv = " & sCodigo
    set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    if not rs.eof then
       if rs("segaceso") <> 0 then srestric = "S"
    end if
    call ofv.cerrarconsulta(rs)



    if srestric = "S" then
            StrSql = "Select * From subtemaskr "
            StrSql = StrSql & "Where codprodserv = " & sCodigo
            StrSql = StrSql & " and (grpusr = 'USR-" & Ucase(session("usuario")) & "' "

'''''''''''''''''''''''''''''''''''''''''''
            if  Session("DELLOG") = "S" then

                XMFTGRP = split(session("SegGrpLst"),chr(9))
                xcount  = ubound(XMFTGRP) + 1
                XMFSEP  = ""
                StrSql = StrSql & " or grpusr in ("

                for XMFI = 0 to ubound(XMFTGRP)
                    StrSql = StrSql  & XMFSEP & "'GRP-" & XMFTGRP(XMFI) & "'"
                    XMFSEP  = ","
                next

                StrSql = StrSql & ") "

            else
                StrSql = StrSql & " or grpusr = 'GRP-" & UCase(Session("seggrp")) & "' "
            end if


''''''''''''''''''''''''''''''''''''''''''

            StrSql = StrSql & ")"

'response.write strsql & "<br>"
       set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
       if not rs.eof then
          srestric = "N"
       else
          srestric = "S"
       end if
       call ofv.cerrarconsulta(rs)
    end if




 end if

 END IF

'xxxxxxxxxxxxxxxxxx

  sCodigo = scodi
  sTipo = stipi

  sHTML = "<html>"
'  sHTML = sHTML & ofv.FormStyler()

     wxs = "<script language=JavaScript>" & Chr(13)
     wxs = wxs & " function IniWin() { " & Chr(13)
     wxs = wxs & " document.ICONACT.submit(); " & Chr(13)
     wxs = wxs & " } " & Chr(13)
     wxs = wxs & "function AbrirWin() { " & Chr(13)
     wxs = wxs & " document.festad.submit(); " & Chr(13)
     wxs = wxs & " } " & Chr(13)
     wxs = wxs & "function closWin() { " & Chr(13)
     wxs = wxs & " document.init.submit(); " & Chr(13)
     wxs = wxs & " } " & Chr(13)
     wxs = wxs & "</script>" & Chr(13)

    sHTML = sHTML & wxs


  if sFuncion = "E" then
     sHTML = sHTML & "<body  onload='IniWin();' onunload='closWin();AbrirWin();' >"
  else
     sHTML = sHTML & "<body  onload='IniWin();' onunload='AbrirWin();' >"
  end if

  ''''   sHTML = sHTML & "<P>" & SESSION("SEGUSUARIO") & " - " & SESSION("SEGGRP") & "</P>"

  perfilsp = "S"

  siconmas = "ICorAcciones.jpg"


 IF SMAINAPL = "OR" OR SMAINAPL = "PL" THEN

     Session("Segct") = "S"
     Session("Segmo") = "S"
     Session("Segco") = "S"
     Session("Segce") = "S"

     sPerfil = "OBJMAIN"
     sAcceso = "S"
     sResol  = "S"

     if session("ModMain") = "WK"  then
        xstrsol     = 0
     else
        xstrsol     = sSClick
     end if

     IF SMAINAPL = "OR" THEN
        sEsquema = "CIRCUITO AUTOEVALUACION"
     ELSE
        sEsquema = "PLAN DE CONTINGENCIAS"
     END IF

     xestado = 1

     sCodigoVINC = 0

     Strsql = "Select estado,CODOBJETOBS,CODOBJETO From ORRSKCIRCUITO Where OBJID = 0" & sCodigo
     set rsx1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     if not rsx1.eof then
        xestado = rsx1(0)
        sCodigoVINC = rsx1(1)
        sCodigoproc = rsx1(2)
     end if
     call ofv.cerrarconsulta(rsx1)

 ELSE


  Strsql = "Select * From Tipobjeto Where tipobjeto ='" & stipo & "'"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then
     sEsquema = rs(1)
     siconmas = rs("archivo")
     sPerfil = rs(3)
     sAcceso = rs(16)
     sResol = rs(17)
  end if

  call ofv.cerrarconsulta(rs)

 sPrint = "S"


 if sAcceso = "N" then
    sPrint = "N"
 elseif session("USACC") = "R" and sAcceso = "L" then
    sPrint = "N"
 elseif session("USRES") = "B" and sResol = "A" then
    sPrint = "N"
 end if



  xestado = 1

  Strsql = "Select estado From Objprty Where codobjeto =" & sCodigo & " and tipobjeto ='" & stipo & "' "
 ' response.write strsql
  set rsx1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rsx1.eof then
     xestado = rsx1(0)
  end if

'''  if session("ModMain") = "WK" and sFuncion = "E" then
  if session("ModMain") = "WK"  then
     xstrsol     = 0
  else
     xstrsol     = sSClick
  end if



  if sFuncion <> "E" then

     Session("Segct") = "S"
     Session("Segmo") = "S"
     Session("Segco") = "S"
     Session("Segce") = "S"

   Strsql = "Select * From OBJSDK Where tipobjeto='" & sTipo & "' and codobjeto = " & sCodigo
   set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
   if not rs.eof then


      if Session("Usuario") <> rs(7) then

         if rs(3)  then Session("Segct") = "N"
         if rs(4)  then Session("Segmo") = "N"
         if rs(5)  then Session("Segco") = "N"
         if rs(6)  then Session("Segce") = "N"

         session("verrs")= "rs3=" & rs(3) & "  rs4=" & rs(4) & " rs5=" & rs(5) & " rs6=" & rs(6)

         call ofv.cerrarconsulta(rs)


         Strsql = "Select * From OBJSDKDET Where tipobjeto='" & sTipo
         Strsql = Strsql & "' and codobjeto = "  & sCodigo
         Strsql = Strsql & " and ("
         Strsql = Strsql & "(grpusr = 'USR-" & Session("Usuario") & "') "

'''''''''''''''''''''''''''''''''''''''''''
            if  Session("DELLOG") = "S" then

                XMFTGRP = split(session("SegGrpLst"),chr(9))
                xcount  = ubound(XMFTGRP) + 1
                XMFSEP  = ""
                StrSql = StrSql & " or ( grpusr in ("

                for XMFI = 0 to ubound(XMFTGRP)
                    StrSql = StrSql  & XMFSEP & "'GRP-" & XMFTGRP(XMFI) & "'"
                    XMFSEP  = ","
                next

                StrSql = StrSql & ")) "

            else
                Strsql = Strsql & " or (grpusr = 'GRP-" & Session("Seggrp") & "')"
            end if


''''''''''''''''''''''''''''''''''''''''''


         Strsql = Strsql & ") "

'response.write strsql & "<br>"
         set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
         do until rs.eof

            select case rs(3)
                   case "1"
                        Session("Segct") = "S"
                        Session("Segmo") = "S"
                        Session("Segco") = "S"
                        Session("Segce") = "S"

                   case "2"
                        Session("Segmo") = "S"
                        Session("Segco") = "S"

                   case "3"
                        Session("Segco") = "S"

                   case "4"
                        Session("Segco") = "S"
                        Session("Segce") = "S"

            end select
         rs.movenext
         loop
      end if
    end if
    call ofv.cerrarconsulta(rs)

    if srestric = "S" then
       Session("Segct") = "N"
       Session("Segmo") = "N"
       Session("Segco") = "N"
       Session("Segce") = "N"
    end if
  end if

 END IF

 IF SMAINAPL = "OR" OR SMAINAPL = "PL" THEN

    xtipobj = STIPO


  bstitulox = GENPROCSECUENEX(sCodigoproc)

    bstitulo = "&nbsp;" & bstitulox

    xtittemp =  bstitulox


  Strsql = "Select * From Objetos Where codobjeto = 0" & sCodigoVINC
'  response.write "sql " & strsql & "<br>"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then

 '   if session("FULLDOCN") = "1" then
 '      bstitulox = sEsquema & ": " & rs(3)
 '   elseif session("FULLDOCN") = "2" then
 '      bstitulox = sEsquema & ": " & rs(1) & " - " & rs(3)
 '   else
 '      bstitulox = sEsquema & ": " & rs(1) & " - " & rs(3)
 '   end if

 '        bstitulox = bstitulox & "aaaaaaaaa"

 '   if sFuncion = "E" then
 '     bstitulo = sEsquema & ": <b>" & rs(1) & " - " & rs(3) & "</b>"
 '   else
 '     bstitulo = sEsquema & ": <b>" & rs(1) & " - " & rs(3) & "</b>"
 '   end if
 '   bstitulo = "&nbsp;" & Ucase(bstitulo)

 '   xtittemp =  UCase(rs(1) & " - " & rs(3)) & ""

 END IF



 ELSE

  Strsql = "Select * From Objetos Where codobjeto=" & sCodigo
'  response.write "sql " & strsql & "<br>"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then

     xtipobj = rs(2)

    if session("FULLDOCN") = "1" then
       bstitulox = sEsquema & ": " & rs(3)
    elseif session("FULLDOCN") = "2" then
       bstitulox = sEsquema & ": " & rs(1) & " - " & rs(3)
    else
       bstitulox = sEsquema & ": " & rs(1) & " - " & rs(3)
    end if



    if sFuncion = "E" then
      bstitulo = sEsquema & ": <b>" & rs(1) & " - " & rs(3) & "</b>"
    else
      bstitulo = sEsquema & ": <b>" & rs(1) & " - " & rs(3) & "</b>"
    end if
    bstitulo = "&nbsp;" & Ucase(bstitulo)

    xtittemp =  UCase(rs(1) & " - " & rs(3)) & ""

  END IF


 END IF

 ' response.write "eof " & rs.eof & "<br>"

   IF  Session("FULLGUI") = "N" THEN
       styl = "background-color:#888888;"
   else
       styl = XspEx("#888888","#cccccc",0) & ";"
   end if

   sHTML = sHTML & OFV.GenTabla("DBpanel", "aaa style='border-style:solid;border-color:#ffffff;border-width:1;" & styl & "' ", "", "100%", "", "0", "", "-2", "2", "", "", "", "no")


    if len(xtittemp) > 60 then xtittemp = left(xtittemp,57) & "..+"

        sInput = "<a>" & xtittemp & "</a>"
        xsclase = "at"

    h1 = OFV.gencelda("", "at style='padding-left:3px;' ", "", "2", "", "", "", "<IMG valign=middle border=0 src='../coress20/imagenes/logocorPnewicoSm1.gif'  style='width:18px;' >", "si")
    h1 = h1 & ofv.GenCelda("",xsclase,"","","","","",sInput,"si")


  if  Session("Segco") = "S" and sPrint = "S" then
      sImagen1   = "<FONT FACE=corflaticon COLOR='#ffffff' SIZE=3 onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='#ffffff';"" >&#0202;</FONT>"
      t2 = "<a class='nt' style='color:#ffffff;' href='Reportes.asp?ID=" & sCodigo & "&Titulo=Esquemas_Normativos&perfil="
      t2 = t2 & sPerfil & "&tipobj=" & xtipobj & "&GLBSOLN=" & "ESPREP"  & "' Target=_blank>"
      t2 = t2 & "<font color='#ff0000' size=2 >&#0124;</font>&nbsp;&nbsp;IMPRIMIR&nbsp;&nbsp;<font color='#ff0000' size=2 >&#0124;</font>" & "</a>"
    '''  t2 = t2 & sImagen1 & "</a>"
      SBOTTIT = "Vista&nbsp;previa&nbsp;de&nbsp;impresión&nbsp;del&nbsp;Esquema."
      h1 = h1 & ofv.GenCelda("",xsclase & " title='" & SBOTTIT & "' ","center","1%","","","",t2,"si")

  end if


        sInput = "&nbsp;"

  if sFuncion = "E" then

     strsql = "SELECT * from OBJOBSCMT where  "
     strsql = strsql & " codobjeto = " & sCodigo & " and tipobjeto = '" & sTipo & "'"
'     strsql = strsql & " codobjeto = " & sCodigo & " and tipobjeto = '" & sTipo & "' and "
'     strsql = strsql & " actpend is null "

' response.write strsql
     set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)

     if not rs2.eof then

        sInput = "<font face='corflaticon' color='#bb0000' size=3 title='Hay observaciones pendientes.' style='background-color:#f0f0f0;border-style:solid;border-color:#993333;border-width:1;cursor:pointer;' >&#093;</font>"
        h1 = h1 & ofv.GenCelda("",xsclase,"center","1%","","","",sInput,"si")

     end if

 end if






  sHTML = sHTML & ofv.GenRow("","","","","","","",h1,"si")

    sHTML = sHTML & ofv.GenFTabla()
  call ofv.cerrarconsulta(rs)

  xsCorEsq = ucase(bstitulox)

  sReporte = "Reportes"

  sMenu="x"
  sSector="x"

  if sFuncion = "C" or sFuncion = "D" or sFuncion = "CD" then
    agregar = ""
    sMenu = request.querystring("sMenu")
    if sMenu = "PORTEMA" then
      xLista = request.querystring("lista")
      xEsquema = request.querystring("esquema")
      agregar = agregar & ofv.GenerarInput("lista",xLista,"hidden","","10","bt","")
      agregar = agregar & ofv.GenerarInput("esquema",xEsquema,"hidden","","10","bt","")
    elseif sMenu = "PORINSTRUMENTO" then
      xinstrumento = request.querystring("instrumento")
      xbuscar = request.querystring("buscar")
      agregar = agregar & ofv.GenerarInput("instrumento",xinstrumento,"hidden","","10","bt","")
      agregar = agregar & ofv.GenerarInput("buscar",xbuscar,"hidden","","10","bt","")
    elseif sMenu = "PORFECHA" then
      xdesde = request.querystring("Desde")
      xhasta = request.querystring("Hasta")
      xfecha = request.querystring("Fecha")
      agregar = agregar & ofv.GenerarInput("Desde",xdesde,"hidden","","10","bt","")
      agregar = agregar & ofv.GenerarInput("Hasta",xhasta,"hidden","","10","bt","")
      agregar = agregar & ofv.GenerarInput("Fecha",xfecha,"hidden","","10","bt","")
    else
      xLista = request.querystring("lista")
      sSector = xLista
      xbuscar = request.querystring("buscar")
      agregar = agregar & ofv.GenerarInput("lista",xLista,"hidden","","10","bt","")
      agregar = agregar & ofv.GenerarInput("buscar",xbuscar,"hidden","","10","bt","")
    end if
  end if



  if sFuncion = "E" then

    sHTML = sHTML & ofv.GenTabla("","aaa style='display:none;' ","","100%","",0,0,0,0,"","","","no")

    for h = UBound(sPath) to 1 step - 1
      if trim(spath(h)) <> "" then
        if h > 1 then
          sRoot = sRoot & ofv.convertircar(sPath(h)," ","&nbsp;") & "\"
        else
          sRoot = sRoot & ofv.convertircar(sPath(h)," ","&nbsp;")
        end if
      end if
    next
    if right(sRoot,1) = "\" then sRoot = left(sRoot,len(sRoot)-1)
    sHTML = sHTML & ofv.GenRow("","","","","","","",ofv.GenCelda("","cc","","100%","","","",sRoot,"si"),"si")


    xverasp = left(Session("ASP"),4)
    xverasp = Ucase(xverasp)

    if xverasp = "NOVE" then
       xverasp = Session("ASP") & "' target='info"
    else
       xverasp = Session("ASP") & "' target='fesquemas"
    end if

    sHTML = sHTML & ofv.GenFTabla()


    sHTML = sHTML & "<form name=init action='" & xverasp & "' "
    sHTML = sHTML & " method=post style='display:none;' >"
    sHTML = sHTML & "<input type=hidden size=10 name=boton value=menu></input>" & chr(13)
    sHTML = sHTML & "</form>"




  end if

    xverasp = "BarraEnCierra2.asp?"
    xverasp = xverasp & "USR=" & REPLACE(Session("Usuario") & "*"," ","_")
    xverasp = xverasp & "&NETUSR=" & REPLACE(Session("UsuarNT") & "*"," ","_")
    xverasp = xverasp & "&DMN=" & REPLACE(Session("DOMINIO") & "*"," ","_")
    xverasp = xverasp & "&LID=" & REPLACE(SESSION("LOGIDUSRDET") & "*"," ","_")
    xverasp = xverasp & "&LIDOBJ=" & REPLACE(SESSION("LOGIDDET") & "*"," ","_")

    sHTML = sHTML & "<form name=festad action='" & xverasp & "' "
    sHTML = sHTML & " method=post target=Estad style='display:none;' >"
    sHTML = sHTML & "<input type=hidden size=10 name=boton value=menu></input>" & chr(13)
    sHTML = sHTML & "</form>"

    xSolasp = ""
    xSolasp = xSolasp & "BarraIcon.asp?Tipo=" & sTipo & "&Codigo="
    xSolasp = xSolasp & sCodigo & "&Menu=" & sMenu & "&Sector=" & sSector & "&MAINAPL=" & SMAINAPL
    xSolasp = xSolasp & "&Funcion=" & sFuncion & "&SSCLICK=" & xstrsol
    xSolasp = xSolasp & "&NIV1=" & sNIV1  & "&NIV2=" & sNIV2  & "&NIV3=" & sNIV3  & "&NIV4=" & sNIV4  & "&NIV5=" & sNIV5

    sHTML = sHTML & "<form name=ICONACT action='" & xSolasp  & "' "
    sHTML = sHTML & " method=post target='WK2' style='display:none;' >"
    sHTML = sHTML & "<input type=hidden size=10 name=boton value=menu></input>" & chr(13)
    sHTML = sHTML & "</form>"

     wxs = "<script language=JavaScript>" & Chr(13)
     wxs = wxs & " window.top.CorEsq = '" & xsCorEsq & "'; " & Chr(13)
     wxs = wxs & " window.top.focus(); " & Chr(13)
     wxs = wxs & "</script>" & Chr(13)

    sHTML = sHTML & wxs



  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"


end if

Response.Write sHTML

 FUNCTION GENPROCSECUENEX(xxcodobj)

    GENPROCSECUENEX = ""

    Strsql = "Select * From CSRELPATHDESCRI  where COBJ = " & xxcodobj
 '' response.write strsql & "<br>"
    set Wrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    if NOT Wrs.eof then

       WKITOPE = 8

       FOR WKI=0 TO 7

           WKIPROC = WRS(WKI) & ""

           IF WKIPROC <> "" THEN
              WKITOPE = WKI
              EXIT FOR
           END IF


       NEXT

       WKICCOL = 7 - WKITOPE

       WKINCOL = 0

       FilaW = ""

           Celda = ""
           Celda2 = ""

       FOR WKI=WKITOPE TO 7




           Celda = Celda & Celda2 & WRS(WKI)

           Celda2 = " - "



       NEXT


       IF WKITOPE < 8 THEN

          GENPROCSECUENEX = Celda

       END IF



    END IF


 END FUNCTION


%>
