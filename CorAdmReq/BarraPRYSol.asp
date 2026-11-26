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

  tsk              = request.querystring("tsk")
  sSClick          = request.querystring("SSClick")


  dim stabla(8)
  dim solord(8)

  solord(0) = 0
  solord(1) = 1
  
  'response.write session("ModMain")="CK"
  
  if  session("ModMain")="CK" then
   solord(2) = 2
   solord(5) = 5
  else
   solord(2) = 2
   solord(5) = 5
  end if
  
  solord(3) = 3
  solord(4) = 4
  'solord(5) = 2
  solord(6) = 6
  solord(7) = 7
  solord(8) = 8





  dim Nombres(8)

  dim ASPs(8)

  ASPs(0) = "CoWKFselect.asp?SKIND=1"


  if sFuncion = "E" then
     ASPs(1) = "proytaskadv.asp"
  else
     ASPs(1) = "proytaskadv.asp"
  end if

  ASPs(2) = "proyprop.asp"
  ASPs(3) = "PROYRESP.asp"
  ASPs(4) = "PROYTASK.asp"
  ASPs(5) = "PROYEST.asp"
  ASPs(6) = "PROYEXECUTIVE.asp"
  ASPs(7) = "PRYARCHIVOS.asp"
  ASPs(8) = "proyRPLLST.asp"

  Nombres(0) = "WORKFLOW"
  Nombres(1) = "AVANCE"
  Nombres(2) = "PROYECTO"
  Nombres(3) = "RESPONSABLES"
  Nombres(4) = "TAREAS"
  Nombres(5) = "ESTADOS"
  Nombres(6) = "RESUMEN EJECUTIVO"
  Nombres(7) = "DOCUMENTOS"
  Nombres(8) = "PLANIFICACIONES"




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

stilo1 = " td.a1 {border-style:solid;border-width:1;border-color:#f7aa22;position:relative;top:+5;}"
stilo2 = " td.a2 {border-color:#f7aa22;border-top-style:inset;border-left-style:inset;border-right-style:inset;border-width:1;position:relative;top:+2;}"
stilo3 = " td.a3 {border-color:#f7aa22;border-bottom-style:inset;border-left-style:inset;border-right-style:inset;border-width:1;}"
stilo4 = " td.a4 {background-color:#f7aa22;border-style:solid;border-width:1;border-color:#f7aa22;position:relative;top:+5;}"






  XCONTROL = "N" 
  XCPT     = "N" 
  XCADV    = "N" 

  StrSql = "Select * From PROYRESP where codigo = '" & session("NCliente") & "' AND USRAFC = '" & SESSION("USUARIO") & "' AND TIPO = 'AN' "
  set rsX1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  if rsx1.eof then

     StrSql = "Select * From PROYRESP where codigo = '" & session("NCliente") & "' AND USRAFC = '" & SESSION("USUARIO") & "' "
     set rsX1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  end if
  if NOT rsX1.eof then

     IF RSX1(11) THEN
        XCONTROL = "S" 
        XCPT     = "S" 
        XCADV    = "S" 
     ELSE
        IF RSX1(13) THEN XCPT      = "S"
        IF RSX1(12) THEN XCADV     = "S"
     END IF

  END IF

  IF XCONTROL = "N" THEN

     bAgregar   = 0
     bModificar = 0
     bSuprimir  = 0

  END IF



'xxxxxxxxxxxxxxxxxx
  StrSql = "Select * From PROYECTO WHERE CODIGO = '" & SESSION("NCLIENTE") & "' "
     set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     if not rs.eof then
        TITULO = rs(2)
        ASIGNA = rs(14)
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



     StrSql = "Select * From TIPOREQ WHERE TIPOREQ = '100' "
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

    IF ASIGNA = SESSION("USUARIO") THEN

       StrSql = "Select * From requerimiento WHERE CODIGO = '" & SESSION("NCLIENTE") & "' "
       set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
       if not rs.eof then
          t3 = "<a class=cc href='BARRAen.asp"
          t3 = t3 & "' <B><U>"
          t3 = t3 & "VER REQUERIMIENTO ASOCIADO&nbsp;" & sImagenR
          t3 = t3 & "</U></B></A>"
         h1 = h1 & ofv.GenCelda("","VT","RIGHT","30%","","","",t3,"si")
       ELSE 
         h1 = h1 & ofv.GenCelda("","VT","CENTER","30%","","","","<b>CODIGO:&nbsp;" & session("NCliente") & "</b>","si")
       end if
    ELSE 
         h1 = h1 & ofv.GenCelda("","VT","CENTER","30%","","","","<b>CODIGO:&nbsp;" & session("NCliente") & "</b>","si")
    end if







  dim Solapas(8)
  dim Solapas2(8)


  for z=0 to 8
      Solapas(z) = "N"
  next  


  if session("ModMain") = "WK" and sFuncion = "E" then
     Solapas(0)  = Nombres(0)
     Solapas2(0) = ASPs(0)
     xstrsol     = 0
  else
     xstrsol     = 1
  end if




            clase = "a1"

  if session("ModMain") = "WK" and tsk <> "NO" then

             xstrsol     = 1

  IF XCONTROL = "S" OR XCADV = "S"  THEN


             Solapas(1)  = Nombres(1)
             Solapas2(1) = ASPs(1)



  END IF

             Solapas(2)  = Nombres(2)
             Solapas2(2) = ASPs(2)



  else

             xstrsol     = 2

             Solapas(2)  = Nombres(2)
             Solapas2(2) = ASPs(2)



  end if

  IF XCONTROL = "S" THEN

             Solapas(3)  = Nombres(3)
             Solapas2(3) = ASPs(3)

             Solapas(4)  = Nombres(4)
             Solapas2(4) = ASPs(4)





  END IF

  IF XCONTROL = "S" OR XCPT = "S" OR XCADV = "S"  THEN


             Solapas(5)  = Nombres(5)
             Solapas2(5) = ASPs(5)




  END IF

  IF XCONTROL = "S" OR XCPT = "S" THEN


             Solapas(6)  = Nombres(6)
             Solapas2(6) = ASPs(6)


             Solapas(7)  = Nombres(7)
             Solapas2(7) = ASPs(7)

             Solapas(8)  = Nombres(8)
             Solapas2(8) = ASPs(8)


  END IF

'''''''' VER PARA REQUERIMIENTOS

'        for yy = 1 to 0
'             if clng(atabN(yy,0)) = 99 then exit for
'            titc = titc + 1
'            clase = "a1"
'            StrSql = "Select * From reqinfo WHERE CODIGO = '" & SESSION("NCLIENTE") & "' and tipo = " & atabN(yy,0)
'            set rsX = ofv.crearconsultaEx(StrSql,cn,1,parametros)
'            if not rsX.eof then
'               clase = "a4"
'            end if
'            t3 = "<a class=CC href='CoSolapa.asp?Solapa=" & atabN(yy,0) & "&NSOLAPA=" & replace(Ucase(atabN(yy,1))," ","&nbsp;")
'            t3 = t3 & "' target=Principal><B>"
'            t3 = t3 & "&nbsp;" & replace(Ucase(atabN(yy,1))," ","&nbsp;") & "&nbsp;"
'            t3 = t3 & "</B></A>"
'            T2 = T2 & ofv.GenCelda("",clase & " id=xs" & titc & " name=" & clase & "  ","center","","","","",t3,"si")
'        next

''''''''''''''''''''''''''''''''''




'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



  xfinal = ubound(Solapas)



  xSolasp = ""  
  xSolasp = xSolasp & "BarraPRYSol.asp?tsk=" & tsk & "&SSCLICK="

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

         if  xsolini = 99 then  xsolini = solord(x)

         if  sSClick = 99 then  sSClick = xsolini

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



         t3 = xsimagenf & "<a  href='" & xSolasp & solord(x) & "' target='WK1' >"
         t3 = t3 & sInput
         t3 = t3 & "</A>"



         t4 = ofv.GenCelda("",xsclase,"","","","","",t3,"si")

         r1 = r1 & ofv.GenRow("","","","","20","","",t4,"si")




         if clng(sSClick) = clng(solord(x)) then

            xverasp = solapas2(solord(x)) & "?tsk=" & tsk
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


      end if
  next



''-----------------------------------------------------------------------------------------------------


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
