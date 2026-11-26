<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires = 0
  bModificar       = False
  ofv.ObtenerAtributos Session("Form"),""
  bModificar       = ofv.AModificar

  sSClick          = request.querystring("SSClick")


  dim stabla(20)
  dim solord(20)

  solord(0) = 0
  solord(1) = 1
  solord(2) = 2
  solord(3) = 3
  solord(4) = 4
  solord(5) = 5
  solord(6) = 6
  solord(7) = 7
  solord(8) = 8
  solord(9) = 9
  solord(10) = 10
  solord(11) = 11
  solord(12) = 12
  solord(13) = 13
  solord(14) = 14
  solord(15) = 15
  solord(16) = 16
  solord(17) = 17
  solord(18) = 18
  solord(19) = 19
  solord(20) = 20


  dim Nombres(20)

  dim ASPs(20)

  ASPs(0) = "CoWKFREQUER.asp?SKIND=1"


  if sFuncion = "E" then
     ASPs(1) = "EDREQUER.asp?SKIND=1"
  else
     ASPs(1) = "EDREQUER.asp?SKIND=1"
  end if

'  ASPs(2) = "proyprop.asp"
'  ASPs(3) = "PROYRESP.asp"
'  ASPs(4) = "PROYTASK.asp"
'  ASPs(5) = "PROYEST.asp"
'  ASPs(6) = "PROYEXECUTIVE.asp"
'  ASPs(7) = "PRYARCHIVOS.asp"
'  ASPs(8) = "proyRPLLST.asp"

  Nombres(0) = "WORKFLOW"
  Nombres(1) = "REQUERIMIENTO"

'  Nombres(2) = "PROYECTO"
'  Nombres(3) = "RESPONSABLES"
'  Nombres(4) = "TAREAS"
'  Nombres(5) = "ESTADOS"
'  Nombres(6) = "RESUMEN EJECUTIVO"
'  Nombres(7) = "DOCUMENTOS"
'  Nombres(8) = "PLANIFICACIONES"




  set cn = ofv.conectar(ofv.strconn2)



  sHTML = "<html>"

     wxs = "<script language=JavaScript>" & Chr(13)
     wxs = wxs & " var Fbtini = 0;" & Chr(13)
     wxs = wxs & " function IniWin() { " & Chr(13)
     wxs = wxs & "  if (document.all.SSOLACT) { " & chr(13)
     wxs = wxs & " document.SSOLACT.submit(); " & Chr(13)
     wxs = wxs & " } else { " & chr(13)
     wxs = wxs & "           if (Fbtini > 10) {} else { Fbtini++; " & chr(13)
     wxs = wxs & " setTimeout (""IniWin()"", 100); } " & chr(13)
     wxs = wxs & "  } " & chr(13)
     wxs = wxs & " } " & Chr(13)
     wxs = wxs & " </script>" & Chr(13)

    sHTML = sHTML & wxs

  sHTML = sHTML & "<body onload='IniWin();' >"


  set cn = ofv.conectar(ofv.strconn6)


    sImagenP = "<B><FONT FACE=WEBDINGS COLOR=#993333 SIZE=3>&#0143;</FONT></B>"
    sImagenR = "<B><FONT FACE=WEBDINGS COLOR=#993333 SIZE=3>2</FONT></B>"  



'xxxxxxxxxxxxxxxxxx
  StrSql = "Select * From REQUERIMIENTO WHERE CODIGO = '" & SESSION("NCLIENTE") & "' "
     set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     if not rs.eof then
        TITULO = rs(2)
        ASIGNA = rs(14)
        treq   = rs(15)
        cinter = rs(19) & ""
     end if

dim atab(15,1)
dim atabN(15,1)


    for xx = 1 to 15
           atab(xx,0) = 99
           atabN(xx,0) = 99
    next

    for xx = 1 to 9

           atab(XX,0) = xx + 2



    next



     StrSql = "Select * From TIPOREQ WHERE TIPOREQ = '" & treq & "' "
     set rsX = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     if not rsX.eof then




        for yy = 1 to 15
 
            if clng(atab(yy,0)) = 99 then exit for
               xx = atab(yy,0)

               IF RSX(XX) THEN
                  atabN(yy,0) = XX
                  atabN(yy,1) = rsX(xx + 15)
               ELSE
                  atabN(yy,0) = 99
               END IF 
        next

     end if

    if cinter = "" then cinter = session("NCliente")

    IF ASIGNA = SESSION("USUARIO") THEN

       StrSql = "Select * From PROYECTO WHERE CODIGO = '" & SESSION("NCLIENTE") & "' "
       set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
       if not rs.eof then
          t3 = "<a class=cc href='BARRAPRY.asp"
          t3 = t3 & "' <B><U>"
          t3 = t3 & "VER PROYECTO ASOCIADO&nbsp;" & sImagenP
          t3 = t3 & "</U></B></A>"
         h1 = h1 & ofv.GenCelda("","VT","RIGHT","30%","","","",t3,"si")
       ELSE 
         h1 = h1 & ofv.GenCelda("","VT","CENTER","30%","","","","<b>CODIGO:&nbsp;" & cinter & "</b>","si")
       end if
    ELSE 
         h1 = h1 & ofv.GenCelda("","VT","CENTER","30%","","","","<b>CODIGO:&nbsp;" & cinter & "</b>","si")
    end if

        titc =  1


        for yy = 1 to 15
 
            if clng(atabN(yy,0)) < 99 then
                titc = titc + 1
                Nombres(titc) = replace(Ucase(atabN(yy,1))," ","&nbsp;")
                ASPs(titc)    = "CoSolapa.asp?Solapa=" & atabN(yy,0) & "&NSOLAPA=" & replace(Ucase(atabN(yy,1))," ","&nbsp;")

           ' t3 = "<a class=cc href='CoSolapa.asp?Solapa=" & atabN(yy,0) & "&NSOLAPA=" & replace(Ucase(atabN(yy,1))," ","&nbsp;")
           ' t3 = t3 & "' target=Principal><B>"
           ' t3 = t3 & "&nbsp;" & replace(Ucase(atabN(yy,1))," ","&nbsp;") & "&nbsp;"
           ' t3 = t3 & "</B></A>"


            end if


        next



  dim Solapas(20)
  dim Solapas2(20)


  for z=0 to 20
      Solapas(z) = "N"
  next  


  if session("ModMain") = "WK"  then
     Solapas(0)  = Nombres(0)
     Solapas2(0) = ASPs(0)
     xstrsol     = 0
  else
     xstrsol     = 1
  end if




  Solapas(1)  = Nombres(1)
  Solapas2(1) = ASPs(1)


        titc =  1


        for yy = 1 to 15
 
            if clng(atabN(yy,0)) < 99 then
                titc = titc + 1
                Nombres(titc) = replace(Ucase(atabN(yy,1))," ","&nbsp;")
                ASPs(titc)    = "CoSolapa.asp?Solapa=" & atabN(yy,0) & "&NSOLAPA=" & replace(Ucase(atabN(yy,1))," ","&nbsp;")

                Solapas(titc)  = Nombres(titc)
                Solapas2(titc) = ASPs(titc)


           ' t3 = "<a class=cc href='CoSolapa.asp?Solapa=" & atabN(yy,0) & "&NSOLAPA=" & replace(Ucase(atabN(yy,1))," ","&nbsp;")
           ' t3 = t3 & "' target=Principal><B>"
           ' t3 = t3 & "&nbsp;" & replace(Ucase(atabN(yy,1))," ","&nbsp;") & "&nbsp;"
           ' t3 = t3 & "</B></A>"


            end if


        next







'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



  xfinal = ubound(Solapas)



  xSolasp = ""  
  xSolasp = xSolasp & "BarraREQSol.asp?tsk=" & tsk & "&SSCLICK="

 r1 = ""


     If ofv.ModTitCls = "1" Or ofv.ModTitCls = "0" Then
        xGenConta = ofv.GetTit(clng(ofv.ModTitCls) + 5)
        xGenConta = ofv.FTitDat1

        xsclase = "xts"

     else



        xsclase = "ttr"

     End If

      xsimagenf  = "<font FACE=WingDINGS COLOR='#ffffff' SIZE=2 onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='#ffffff';"" >&#0167;</font>&nbsp;&nbsp;" 
      xsimagend  = "<font FACE=WingDINGS COLOR='#ffffff' SIZE=2 onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='#ffffff';"" >0</font>&nbsp;&nbsp;" 



  xsolini = 99



  for x = xstrsol to xfinal
      if solapas(solord(x)) <> "N" then
         sSolapa =  solapas(solord(x))

         SOPC = "" 

         if  xsolini = 99 then  xsolini = solord(x)

         if  sSClick = 99 then
             sSClick = xsolini
             IF  xsolini = 1 THEN
                 SOPC = "&OPCION=RELOAD" 
             END IF 
         END IF 

         If ofv.ModTitCls = "1" Or ofv.ModTitCls = "0" Then
         '  xGenCont = xGenConta
         '  xGenConti = replace(Ucase(sSolapa),"*"," ")

         '  xGenCont = Replace(xGenCont, "XXXDATOXXX", xGenConti)
         '  xGenCont = Replace(xGenCont, "XXXWIDTHXXX", "98%")
         '  xGenCont = Replace(xGenCont, "XXXIDXXX", "DBED1R")

         '  sInput = xGenCont



        sInput = replace(Ucase(sSolapa),"*"," ")
        xsclase = "at style='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#dddddd;padding-left:3px;'"


         else

           sInput = "<b>" & "&nbsp;" & replace(Ucase(sSolapa),"*","&nbsp;") & "&nbsp;" & "</B>"

         End If


'''''''''''''''''''''''

         t3 = xsimagenf & "<a  href='" & xSolasp & solord(x) & "' target='WK1' >"
         t3 = t3 & sInput
         t3 = t3 & "</A>"



         t4 = ofv.GenCelda("",xsclase,"","","","","",t3,"si")

         r1 = r1 & ofv.GenRow("","","","","20","","",t4,"si")




         if clng(sSClick) = clng(solord(x)) then

            xverasp = solapas2(solord(x)) & SOPC
            xverasp = xverasp & "&GLBSOLN=S" & solord(x)

            sHTML1 = ""
            sHTML1 = sHTML1 & "<form name=SSOLACT action='" & xverasp & "' "  
            sHTML1 = sHTML1 & " method=post target='Principal' style='display:none;' >" & chr(13)
            sHTML1 = sHTML1 & "<input type=hidden size=10 name=boton value=menu></input>" & chr(13)
            sHTML1 = sHTML1 & "</form>" & chr(13)

            sHTML1 = sHTML1 & "<script language=JavaScript>" & Chr(13)
            sHTML1 = sHTML1 & " window.top.CorSol='" & replace(Ucase(sSolapa),"*"," ") & "'; " & Chr(13)
            sHTML1 = sHTML1 & "</script>" & Chr(13)


        IF  Session("FULLGUI") = "N" THEN
            styl = "background-color:#888888;"
        else     
            styl = XspEx("#888888","#cccccc",0) & ";"
        end if

'''
''

          xlfun = " onclick="" document.all.MDBEDL.style.display='block'; document.all.DBEDTD1.style.display='none'; document.all.DBEDTD2.style.display='block'; document.all.MM2.style.display='block'; "" "
          xlfun2 = " onclick=""document.all.MDBEDL.style.display='none'; document.all.DBEDTD1.style.display='block'; document.all.DBEDTD2.style.display='none'; document.all.MM2.style.display='none'; "" "

           macel = ofv.GenCelda("","at style='border-right-style:solid;border-right-width:1;border-right-color:#dddddd;padding-left:3px;' " & xlfun,"","90%","","","",replace(Ucase(sSolapa),"*"," "),"si")

          xsclase2 = "at style='border-right-style:solid;border-right-width:1;border-right-color:#dddddd;display:block;font-family:webdings;font-size:10pt;color:#cc0000;cursor:hand;"


          macel = macel & OFV.gencelda("DBEDTD1", xsclase2 & "display:block;' ", "center title='Ver solapas.' " & xlfun, "center", "1", "", "", "1", "si")
          macel = macel & OFV.gencelda("DBEDTD2", xsclase2 & "display:none;' ", "center title='Ocultar Solapas.' " & xlfun2, "center", "1", "", "", "0", "si")

           marow = ofv.GenRow("","","","","20","","",macel,"si")

           sHTML = sHTML & OFV.GenTabla("DBED", "aaa style='border-style:solid;border-color:#ffffff;border-width:1;" & styl & "' ", "", "100%", "", "0", "", "-2", "-2", "", "", marow, "si")



         end if 




''''''''''''''''''''''





      end if
  next



''''''''''''''''''''''''''''''''''''''''''''
  sHTML = sHTML & sHTML1


           sHTML = sHTML & "<div id='MDBEDL' name='MDBEDL' style='display:none;height:240px;overflow:auto;" & WStyl & "' >"

        IF  Session("FULLGUI") = "N" THEN
            styl = "background-color:#888888;"
        else     
            styl = XspEx("#888888","#cccccc",0) & ";"
        end if

            styl = ""

        If ofv.ModTitCls = "1" Then

           sHTML = sHTML & OFV.GenTabla("DBEDL", "aaa style='border-style:solid;border-color:#ffffff;border-width:1;" & styl & "' ", "", "100%", "", "0", "", "-2", "-2", "", "", r1, "si")

        else

           sHTML = sHTML & OFV.GenTabla("DBEDL", "aaa style='border-style:solid;border-color:#ffffff;border-width:1;" & styl & "' ", "", "100%", "", "0", "", "-2", "-2", "", "", r1, "si")

        end if

           sHTML = sHTML & "</div>"

  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"


end if

Response.Write sHTML

%>
