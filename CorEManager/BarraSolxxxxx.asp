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

  sTipo     = request.querystring("Tipo")
  sCodigo   = request.querystring("Codigo")
  sFuncion  = request.querystring("Funcion")
  sSClick   = request.querystring("SSClick")
  sMenu     = request.querystring("Menu")
  sSector   = request.querystring("Sector")
  sNIV1     = request.querystring("NIV1")
  sNIV2     = request.querystring("NIV2")
  sNIV3     = request.querystring("NIV3")
  sNIV4     = request.querystring("NIV4")
  sNIV5     = request.querystring("NIV5")
  sMAINAPL = request.querystring("MAINAPL")

''  response.write "recibe " & request.querystring & "<br>"

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


  dim stabla(20)
  dim solord(20)
  dim Nombres(20)
  dim ASPs(20)
  dim Solapas(20)
  dim Solapas2(20)


  ASPs(0) = "CoWKFselect.asp?SKIND=1"


 IF SMAINAPL = "OR" OR SMAINAPL = "PL" THEN





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



  perfilsp = "S"

     IF SMAINAPL = "OR" THEN
        sEsquema = "CIRCUITO AUTOEVALUACION"
     ELSE
        sEsquema = "PLAN DE CONTINGENCIAS"
     END IF

     xestado = 1

     Strsql = "Select estado,CODOBJETO,TIPOBJETO,RELLINK From ORRSKCIRCUITO Where OBJID = 0" & sCodigo
     set rsx1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     if not rsx1.eof then
        xestado = rsx1(0)
        sCodigoVINC = rsx1(1)
        sTIPOVINC = rsx1(2)
        sRELVINC = rsx1(3)
     end if
     call ofv.cerrarconsulta(rsx1)

	'response.Write(session("corapli"))

	if (sFuncion = "CD") then
		ASPs(1) = "ORMATRIXIMPACTO.asp?FILSEL=CRT&PROCESO=" & sCodigoVINC & "&SUBFILTRO=ALL&PROCMOD=WKF&SOLOCONSULTA=0"
	else
		ASPs(1) = "ORMATRIXIMPACTO.asp?FILSEL=CRT&PROCESO=" & sCodigoVINC & "&SUBFILTRO=ALL&PROCMOD=WKF&SOLOCONSULTA=1"
	end if




  for z=0 to 20
      Solapas(z) = "N"
  next


    for xx = 1 to 20
        Nombres(xx) = "N"
    next

     'Solapas(1)  = REPLACE("DETALLE DEL CIRCUITO"," ","*")
     Solapas(1)  = REPLACE("MATRIZ DE AUTOEVALUACION"," ","*")
     Solapas2(1) = ASPs(1)
     xstrsol      = 1
     xfinal       = 1
  '''   sSClick      = 99



  if session("ModMain") = "WK"  then
     Solapas(0)  = "WORKFLOW"
     Solapas2(0) = ASPs(0)
     xstrsol     = 0
  else
     xstrsol     = 1
  end if



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

 ELSE


  ASPs(1) = "edPropiedades.asp?SKIND=1"

  if sFuncion = "E" then
    ''' ASPs(2) = "edConsideraciones.asp?SKIND=1"
     ASPs(2) = "EDNWEDT.asp?SKIND=1"
  else
 '''    ASPs(2) = "CoConsideraciones.asp?SKIND=1"
     ASPs(2) = "CONWEDT.asp?SKIND=1"
  end if

  if sFuncion = "E" then
    ASPs(3) = "edSectores.asp?SKIND=1"
  else
    ASPs(3) = "CoSectores.asp?SKIND=1"
  end if

  ASPs(4) = "edEventos.asp?SKIND=1"

  if sFuncion = "E" then
    ASPs(5) = "edProcedimientos.asp?SKIND=1"
  else
    ASPs(5) = "CoProcedimientos.asp?SKIND=1"
  end if

  if sFuncion = "E" then
    ASPs(16) = "CorVerTodoseventos.asp?SKIND=1"
  else
    ASPs(16) = "CorVerTodoseventos.asp?SKIND=1"
  end if

  if sFuncion = "E" then
     ASPs(6) = "edObjDsc.asp?SKIND=1"
  else
     ASPs(6) = "CoObjDsc.asp?SKIND=1"
  end if

  if sFuncion = "E" then
    ASPs(7) = "edControles.asp?SKIND=1"
  else
    ASPs(7) = "CoControles.asp?SKIND=1"
  end if
  if sFuncion = "E" then
    ASPs(8) = "edEsquemas.asp?SKIND=1"
  else
    ASPs(8) = "CoEsquemas.asp?SKIND=1"
  end if
  if sFuncion = "E" then
    ASPs(9) = "edTablas.asp?SKIND=1"
  else
    ASPs(9) = "CoTablasCons.asp?SKIND=1"
  end if

  if sFuncion = "E" then
    ASPs(10) = "edarchivos.asp?SKIND=1"
  else
    ASPs(10) = "CoArchCons.asp?SKIND=1"
  end if

  if sFuncion = "E" then
    ASPs(11) = "edestr.asp?SKIND=1"
  else
  ''  ASPs(11) = "Coestrex.asp?SKIND=1"
    ASPs(11) = "CoestrCons.asp?SKIND=1"
  end if

  if sFuncion = "E" then
    ASPs(12) = "edProcedimientos.asp?SKIND=1"
  else
    ASPs(12) = "CursogramaLnch.asp?SKIND=1"
  end if

  if sFuncion = "E" then
    ASPs(13) = "edarchivos.asp?SKIND=2"
  else
    ASPs(13) = "CoArchCons.asp?SKIND=2"
  end if

  if sFuncion = "E" then
    ASPs(14) = "edFAQ.asp?SKIND=1"
  else
    ASPs(14) = "CoFAQ.asp?SKIND=1"
  end if


  if sFuncion = "E" then
    ASPs(15) = "edGeneral.asp?SKIND=1"
  else
    ASPs(15) = "CoGeneral.asp?SKIND=1"
  end if



  solord(0) = 0
  solord(1) = 15
  solord(2) = 1
  solord(3) = 6
  solord(4) = 3
  solord(5) = 2
  solord(6) = 4
  solord(7) = 5
  solord(8) = 7
  solord(9) = 8
  solord(10) = 11
  solord(11) = 10
  solord(12) = 9
  solord(13) = 12
  solord(14) = 13
  solord(15) = 14
  solord(16) = 16
  solord(17) = 17
  solord(18) = 18
  solord(19) = 19
  solord(20) = 20




  perfilsp = "S"


  siconmas = "ICorAcciones.jpg"
  Strsql = "Select * From Tipobjeto Where tipobjeto ='" & stipo & "'"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then
     sEsquema = rs(1)
     sPerfil = rs(3)
  end if

  call ofv.cerrarconsulta(rs)


  for z=0 to 20
      Solapas(z) = "N"
  next




  Strsql = "Select * From Objperfildoc Where codobjeto =" & sCodigo & " and tipobjeto ='" & stipo & "' "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if rs.eof then
     perfilsp = "N"
  end if

  if perfilsp = "N" then
     call ofv.cerrarconsulta(rs)
     Strsql = "Select * From Objperfil Where tipoperfil='" & sPerfil & "'"
     set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  end if

  if Session("Segco") <> "S" then
    for xx = 1 to 20
        Nombres(xx) = "N"
    next

     Solapas(15)  = REPLACE(RS(31)," ","*")
     Solapas2(15) = ASPs(15)
     xstrsol      = 1
     xfinal       = 1
     sSClick      = 99


  else

    for xx = 1 to 20
        solord(xx) = 17
        Nombres(xx) = "N"
    next



    for xx = 2 to 16
        stabla (xx - 1) = rs(xx + 30)
    next



    yy = 0
    for xx = 1 to 15

        if cint(stabla (xx)) < 99 then
           yy = yy + 1
           solord (stabla (xx)) = xx
        end if
    next
    tyy = yy



'''  if session("ModMain") = "WK" and sFuncion = "E" then
  if session("ModMain") = "WK"  then
     Solapas(0)  = "WORKFLOW"
     Solapas2(0) = ASPs(0)
     xstrsol     = 0
  else
     xstrsol     = 1
  end if





  X=0
  for z = 2 to 16

    if rs(z) Then

      if (sFuncion = "E"  ) or (sFuncion = "C" and z <> 2 and z <> 4 and z <> 5 ) or (sFuncion = "CD" and z <> 2 and z <> 5 )  or (sFuncion = "D"  and z = 4 ) then
          x = x + 1
          Solapas(z-1)  = REPLACE(RS(z + 15)," ","*")
          Solapas2(z-1) = ASPs(z-1)


          if sFuncion <> "E" and z = 6 then

             Solapas(16)  = "CIRCUITO COMPLETO"
             Solapas2(16) = ASPs(16)

             for xiord=0 to 19

                 if solord(xiord) = 5 then

                    xmovey = 16

                    for xiord2=xiord + 1 to 19
                        xmoven           = solord(xiord2)
                        solord(xiord2)   = xmovey
                        xmovey           = xmoven
                    next



                    exit for
                 end if
             next


          end if

      end if
    end if

  next


  xfinal = ubound(Solapas)

 end if

' response.write Solapas2(0) & xestado & sFuncion & "<br>"

  call ofv.cerrarconsulta(rs)


 END IF



  xSolasp = ""
  xSolasp = xSolasp & "BarraSol.asp?Tipo=" & sTipo & "&Codigo="
  xSolasp = xSolasp & sCodigo & "&Menu=" & sMenu & "&Sector=" & sSector & "&MAINAPL=" & SMAINAPL
  xSolasp = xSolasp & "&Funcion=" & sFuncion
  xSolasp = xSolasp & "&NIV1=0&NIV2=0&NIV3=0&NIV4=0&NIV5=0"
  xSolasp = xSolasp & "&SSCLICK="

 r1 = ""


     If ofv.ModTitCls = "1" Or ofv.ModTitCls = "0" Then
        xGenConta = ofv.GetTit(clng(ofv.ModTitCls) + 5)
        xGenConta = ofv.FTitDat1

        xsclase = "xts"

     else



        xsclase = "xts"

     End If

      xsimagenf  = "<font FACE=corflaticon COLOR='#ffffff' SIZE=2 onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='#ffffff';"" >&#0167;</font>&nbsp;&nbsp;"
      xsimagend  = "<font FACE=corflaticon COLOR='#ffffff' SIZE=2 onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='#ffffff';"" >0</font>&nbsp;&nbsp;"



  xsolini = 99


  for x = xstrsol to xfinal
      if solapas(solord(x)) <> "N" then
         sSolapa =  solapas(solord(x))

         if  xsolini = 99 then  xsolini = solord(x)

         if  sSClick = 99 then  sSClick = xsolini

         If ofv.ModTitCls = "1" Or ofv.ModTitCls = "0" Then


            sInput  = replace(Ucase(sSolapa),"*"," ")
            xsclase = "at style='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#dddddd;padding-left:3px;'"


         else

            sInput  = "&nbsp;" & replace(Ucase(sSolapa),"*","&nbsp;") & "&nbsp;"
            xsclase = "at style='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#dddddd;padding-left:3px;'"

         End If


         t3 = xsimagenf & "<a  href='" & xSolasp & solord(x) & "' target='WK1' >"
         t3 = t3 & sInput
         t3 = t3 & "</A>"



         t4 = ofv.GenCelda("",xsclase,"","","","","",t3,"si")

         r1 = r1 & ofv.GenRow("","","","","20","","",t4,"si")

         if clng(sSClick) = clng(solord(x)) then
            session("codSol")=sSClick
            xverasp = solapas2(solord(x)) & "&Tipobjeto=" & sTipo & "&Codobjeto="
            xverasp = xverasp & sCodigo & "&Menu=" & sMenu & "&Sector=" & sSector & "&MAINAPL=" & SMAINAPL
            xverasp = xverasp & "&GLBSOLN=S" & solord(x)
            xverasp = xverasp & "&NIV1=" & sNIV1  & "&NIV2=" & sNIV2  & "&NIV3=" & sNIV3  & "&NIV4=" & sNIV4  & "&NIV5=" & sNIV5

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

          xlfun = " onclick=""document.all.MDBEDL.style.display='block'; document.all.DBEDTD1.style.display='none'; document.all.DBEDTD2.style.display='block'; document.all.MM2.style.display='block'; "" "
          xlfun2 = " onclick=""document.all.MDBEDL.style.display='none'; document.all.DBEDTD1.style.display='block'; document.all.DBEDTD2.style.display='none'; document.all.MM2.style.display='none'; "" "

           macel = ofv.GenCelda("","at style='border-right-style:solid;border-right-width:1;border-right-color:#dddddd;padding-left:3px;' " & xlfun,"","90%","","","",replace(Ucase(sSolapa),"*"," "),"si")

          xsclase2 = "at style='border-right-style:solid;border-right-width:1;border-right-color:#dddddd;display:block;font-family:corflaticon;font-size:10pt;color:#cc0000;cursor:pointer;"


          macel = macel & OFV.gencelda("DBEDTD1", xsclase2 & "display:block;' ", "center title='Ver solapas.' " & xlfun, "center", "1", "", "", "1", "si")
          macel = macel & OFV.gencelda("DBEDTD2", xsclase2 & "display:none;' ", "center title='Ocultar Solapas.' " & xlfun2, "center", "1", "", "", "0", "si")

           marow = ofv.GenRow("","","","","20","","",macel,"si")

           sHTML = sHTML & OFV.GenTabla("DBED", "aaa style='border-style:solid;border-color:#ffffff;border-width:1;" & styl & "' ", "", "100%", "", "0", "", "-2", "-2", "", "", marow, "si")



         end if


      end if
  next

 IF SMAINAPL = "OR" OR SMAINAPL = "PL" THEN

 ELSE

''-----------------------------------------------------------------------------------------------------

''''''''''''''''''''''''''''''''''''''''''''
'' Accesos directos del documento ''''''''''
''''''''''''''''''''''''''''''''''''''''''''

  dim ATABSCPT(200,2)

  ISCPT = 0
  TSCPT = 0

    ''    response.write "entro1"

      if sFuncion <> "X"   then

      ''  response.write "entro2"

         ReDim Preserve bConsultar(4)

      if sFuncion = "E"   then

         bConsultar(2) = 0
         bConsultar(4) = 0

         bConsultar(1) = 1
         bConsultar(3) = 1

      else

         bConsultar(2) = 0
         bConsultar(4) = 0

         bConsultar(1) = 0
         bConsultar(3) = 1

      end if


         sVerEstado = "  and Objetos.objnivel >= " & Session("SegNivel") & " "

         sverestado = sverestado & " and Objprty.Estado In (9,"
         For i = 1 To 4
             If bConsultar(i) Then sVerEstado = sVerEstado & i & ","
         Next
         sVerEstado = Left(sVerEstado, Len(sVerEstado) - 1) & ") "

'''''''''''''''''''''''''''''''''''''''''



  sActcon  = "Consultar.asp?Codigo=NNNN&Tipo=TTTT"
  sActcon  = sActcon & "&Datos=*smenu=PORESTADO*lista=3"
  sHref    = "&nbsp;"
' response.write sActcon
  sFuncion = " onclick=" & Chr(34) & " window.top.abreWin('" & sActcon & "','TTTT'); return false;" & Chr(34)


  Strsql = "SELECT * FROM objacd  "
  Strsql = Strsql & " Where codobjeto=" & sCodigo
  Strsql = Strsql & " And tipobjeto = '" & stipo & "' "

  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  do until rs.eof

    ''   response.write "entro3"


     if  rs(3) = "S" then
'''' Carpeta superior

       if  rs(6) = "S" then

''----------------

         strsql  = "select * from csbarr where codojasoc = " & sCodigo
         Strsql  = Strsql & " and tipobjasoc = '" & sTipo & "'"
         set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
         if  not rs.eof then
             essCodigo = rs1(1)
             essTipo   = rs1(2)

         else
             essCodigo = 0
             essTipo   = "NO"
         end if
         call ofv.cerrarconsulta(rs1)


         sNIVE   = "no"

         StrSql = "Select archivo, descri, perfil, TIPOCPT From Tipobjeto "
         StrSql = StrSql & "Where Tipobjeto.tipobjeto = '" & essTipo & "'"
         Set rsIcono = OFV.crearconsultaEx(StrSql, cn, 1, "")
         If Not rsIcono.EOF Then

                sIcono  = rsIcono(0)
                sTipoE  = rsIcono(1)
                sPerfil = rsIcono(2)
                sNIVE   = rsIcono(3)

         End If
         Call OFV.cerrarconsulta(rsIcono)

         if sNIVE   = "C"  then

            strsql  = "select * from csbarr where codojasoc = " & essCodigo
            Strsql  = Strsql & " and tipobjasoc = '" & essTipo & "'"
            set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
            if  not rs.eof then

                essCodigo = rs1(1)
                essTipo   = rs1(2)

            end if
            call ofv.cerrarconsulta(rs1)


'''''''''''''''''''''''''''''''''''''''''

         StrSql = "SELECT DISTINCT OBJETOS.codobjeto, OBJETOS.descri, OBJETOS.tipobjeto, Estados.estado, "
         StrSql = StrSql & "Relaobj.tipobjasoc, OBJPRTY.fecemis, OBJPRTY.fecvigen, OBJPRTY.fecbaja,"
         StrSql = StrSql & " OBJPRTY.autor,OBJPRTY.aprobo, OBJPRTY.elimino, OBJPRTY.fecvto, OBJETOS.OBJUN"
         StrSql = StrSql & " FROM RELAOBJ, OBJETOS, OBJPRTY, Estados "
         StrSql = StrSql & " where "
         StrSql = StrSql & " RELAOBJ.codojasoc = OBJETOS.codobjeto and "
         StrSql = StrSql & " OBJPRTY.estado    = Estados.objid and "
         StrSql = StrSql & " OBJETOS.codobjeto = OBJPRTY.codobjeto  and "
         StrSql = StrSql & " (objetos.objun =  " & session("UN") & " or "
         StrSql = StrSql & "  objetos.objun =  1 )  and "
         StrSql = StrSql & " RELAOBJ.codobjmaster = " & essCodigo
         StrSql = StrSql & " and relaobj.tipobjmaster = '" & essTipo & "' " & sVerEstado
         StrSql = StrSql & " order by OBJETOS.tipobjeto, OBJETOS.descri, Estados.estado "
      ''  response.write "<br>" & strsql & "<br>"
         sVino = "N"
         Set rs2 = OFV.crearconsultaEx(StrSql, cn, 1, "")
         If  not rs2.EOF Then

        Do Until rs2.EOF

             if cstr(rs2(0) & "")  <> cstr(scodigo & "") then

             sVinook = "N"
             if  cstr(rs2(12) & "") = cstr(session("UN") & "") then
   '' response.write "<br><br>entro"
                 sVino   = "S"
                 sVinook = "S"
             end if

          if SESSION("EMPMUN") = "S" THEN
             StrSql = "Select OBJMEMPALC.CODEMPRE  from OBJMEMPALC "
             StrSql = StrSql & " where "
             StrSql = StrSql & " OBJMEMPALC.codobjeto = " & rs2(0) & " and "
             StrSql = StrSql & " OBJMEMPALC.CODEMPRE = " & Session("UN") & " "
             Set rsmas2 = OFV.crearconsultaEx(StrSql, cn, 1, "")
             If  Not rsmas2.EOF Then
                 sVino   = "S"
                 sVinook = "S"
             end if
          end if

          sNIVE   = "D"

        if  sVinook = "S" then

            stitem = rs2(4)
            StrSql = "Select archivo, descri, perfil, TIPOCPT From Tipobjeto "
            StrSql = StrSql & "Where Tipobjeto.tipobjeto = '" & rs2(4) & "'"
            Set rsIcono = OFV.crearconsultaEx(StrSql, cn, 1, "")
            If Not rsIcono.EOF Then
                sIcono  = rsIcono(0)
                sTipoE  = rsIcono(1)
                sPerfil = rsIcono(2)
                sNIVE   = rsIcono(3)

                sitem = sTipoE & ": " & rs2(1)

            End If
            Call OFV.cerrarconsulta(rsIcono)

            if sNIVE   = "D"  then


               sSolapa =  sitem

               If ofv.ModTitCls = "1" Or ofv.ModTitCls = "0" Then

                  sInput = replace(Ucase(sSolapa),"*"," ")

               else

                  sInput = "&nbsp;" & replace(Ucase(sSolapa),"*","&nbsp;") & "&nbsp;"

               End If

               sFuncionDSP = REPLACE(sFuncion,"NNNN",rs2(0))
               sFuncionDSP = REPLACE(sFuncionDSP,"TTTT",rs2(2))

               t3 = xsimagend & "<a  href='javascript:void(0);' " & sFuncionDSP &  " >"
               t3 = t3 & sInput
               t3 = t3 & "</A>"

               t4 = ofv.GenCelda("",xsclase,"","","","","",t3,"si")

               r1 = r1 & ofv.GenRow("","","","","20","","",t4,"si")

            else





            end if


         end if

        end if

         rs2.movenext
         loop


       end if



'''''''''''''''''''''''''''''''''''''''

        end if

''----------------



       end if





     elseif  rs(3) = "P" then
'''' Carpeta primaria

       if  rs(6) = "S" then

         strsql  = "select * from csbarr where codojasoc = " & sCodigo
         Strsql  = Strsql & " and tipobjasoc = '" & sTipo & "'"
         set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
         if  not rs.eof then
             essCodigo = rs1(1)
             essTipo   = rs1(2)

         else
             essCodigo = 0
             essTipo   = "NO"
         end if
         call ofv.cerrarconsulta(rs1)


         sNIVE   = "no"

         StrSql = "Select archivo, descri, perfil, TIPOCPT From Tipobjeto "
         StrSql = StrSql & "Where Tipobjeto.tipobjeto = '" & essTipo & "'"
         Set rsIcono = OFV.crearconsultaEx(StrSql, cn, 1, "")
         If Not rsIcono.EOF Then

                sIcono  = rsIcono(0)
                sTipoE  = rsIcono(1)
                sPerfil = rsIcono(2)
                sNIVE   = rsIcono(3)

         End If
         Call OFV.cerrarconsulta(rsIcono)




         StrSql = "SELECT DISTINCT OBJETOS.codobjeto, OBJETOS.descri, OBJETOS.tipobjeto, Estados.estado, "
         StrSql = StrSql & "Relaobj.tipobjasoc, OBJPRTY.fecemis, OBJPRTY.fecvigen, OBJPRTY.fecbaja,"
         StrSql = StrSql & " OBJPRTY.autor,OBJPRTY.aprobo, OBJPRTY.elimino, OBJPRTY.fecvto, OBJETOS.OBJUN"
         StrSql = StrSql & " FROM RELAOBJ, OBJETOS, OBJPRTY, Estados "
         StrSql = StrSql & " where "
         StrSql = StrSql & " RELAOBJ.codojasoc = OBJETOS.codobjeto and "
         StrSql = StrSql & " OBJPRTY.estado    = Estados.objid and "
         StrSql = StrSql & " OBJETOS.codobjeto = OBJPRTY.codobjeto  and "
         StrSql = StrSql & " (objetos.objun =  0" & session("UN") & " or "
         StrSql = StrSql & "  objetos.objun =  1 )  and "
         StrSql = StrSql & " RELAOBJ.codobjmaster = " & essCodigo
         StrSql = StrSql & " and relaobj.tipobjmaster = '" & essTipo & "' " & sVerEstado
         StrSql = StrSql & " order by OBJETOS.tipobjeto, OBJETOS.descri, Estados.estado "
     ''   response.write "<br>" & strsql & "<br>"
         sVino = "N"
         Set rs2 = OFV.crearconsultaEx(StrSql, cn, 1, "")
         If  not rs2.EOF Then

        Do Until rs2.EOF

             if cstr(rs2(0) & "")  <> cstr(scodigo & "") then

             sVinook = "N"
             if  cstr(rs2(12) & "") = cstr(session("UN") & "") then
   '' response.write "<br><br>entro"
                 sVino   = "S"
                 sVinook = "S"
             end if

          if SESSION("EMPMUN") = "S" THEN
             StrSql = "Select OBJMEMPALC.CODEMPRE  from OBJMEMPALC "
             StrSql = StrSql & " where "
             StrSql = StrSql & " OBJMEMPALC.codobjeto = " & rs2(0) & " and "
             StrSql = StrSql & " OBJMEMPALC.CODEMPRE = " & Session("UN") & " "
             Set rsmas2 = OFV.crearconsultaEx(StrSql, cn, 1, "")
             If  Not rsmas2.EOF Then
                 sVino   = "S"
                 sVinook = "S"
             end if
          end if

          sNIVE   = "D"

        if  sVinook = "S" then

            stitem = rs2(4)
            StrSql = "Select archivo, descri, perfil, TIPOCPT From Tipobjeto "
            StrSql = StrSql & "Where Tipobjeto.tipobjeto = '" & rs2(4) & "'"
            Set rsIcono = OFV.crearconsultaEx(StrSql, cn, 1, "")
            If Not rsIcono.EOF Then
                sIcono  = rsIcono(0)
                sTipoE  = rsIcono(1)
                sPerfil = rsIcono(2)
                sNIVE   = rsIcono(3)

                sitem = sTipoE & ": " & rs2(1)

            End If
            Call OFV.cerrarconsulta(rsIcono)

            if sNIVE   = "D"  then


               sSolapa =  sitem

               If ofv.ModTitCls = "1" Or ofv.ModTitCls = "0" Then
                  sInput = replace(Ucase(sSolapa),"*"," ")
               else
                  sInput = "&nbsp;" & replace(Ucase(sSolapa),"*","&nbsp;") & "&nbsp;"
               End If

               sFuncionDSP = REPLACE(sFuncion,"NNNN",rs2(0))
               sFuncionDSP = REPLACE(sFuncionDSP,"TTTT",rs2(2))

               t3 = xsimagend & "<a  href='javascript:void(0);' " & sFuncionDSP &  " >"
               t3 = t3 & sInput
               t3 = t3 & "</A>"

               t4 = ofv.GenCelda("",xsclase,"","","","","",t3,"si")

               r1 = r1 & ofv.GenRow("","","","","20","","",t4,"si")

            else


            end if


         end if

        end if

         rs2.movenext
         loop


'''''''''''''''''''''''''''''''''''''''

        end if




       end if




     elseif  rs(3) = "R" then
'''' Carpeta dependiente

       if  rs(6) = "S" then

''----------------

               TSCPT = TSCPT + 1

               ATABSCPT(TSCPT,0) = ""
               ATABSCPT(TSCPT,1) = sCodigo
               ATABSCPT(TSCPT,2) = sTipo



         strsql  = "select * from csbarr where codojasoc = " & sCodigo
         Strsql  = Strsql & " and tipobjasoc = '" & sTipo & "'"
         set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
         if  not rs.eof then
             essCodigo = rs1(1)
             essTipo   = rs1(2)

         else
             essCodigo = 0
             essTipo   = "NO"
         end if
         call ofv.cerrarconsulta(rs1)


         sNIVE   = "no"

         StrSql = "Select archivo, descri, perfil, TIPOCPT From Tipobjeto "
         StrSql = StrSql & "Where Tipobjeto.tipobjeto = '" & essTipo & "'"
         Set rsIcono = OFV.crearconsultaEx(StrSql, cn, 1, "")
         If Not rsIcono.EOF Then

                sIcono  = rsIcono(0)
                sTipoE  = rsIcono(1)
                sPerfil = rsIcono(2)
                sNIVE   = rsIcono(3)

         End If
         Call OFV.cerrarconsulta(rsIcono)

         if sNIVE   = "C"  then

           '  essCodigo = sCodigo
           '  essTipo   = sTipo

'''''''''''''''''''''''''''''''''''''''''

         StrSql = "SELECT DISTINCT OBJETOS.codobjeto, OBJETOS.descri, OBJETOS.tipobjeto, Estados.estado, "
         StrSql = StrSql & "Relaobj.tipobjasoc, OBJPRTY.fecemis, OBJPRTY.fecvigen, OBJPRTY.fecbaja,"
         StrSql = StrSql & " OBJPRTY.autor,OBJPRTY.aprobo, OBJPRTY.elimino, OBJPRTY.fecvto, OBJETOS.OBJUN"
         StrSql = StrSql & " FROM RELAOBJ, OBJETOS, OBJPRTY, Estados "
         StrSql = StrSql & " where "
         StrSql = StrSql & " RELAOBJ.codojasoc = OBJETOS.codobjeto and "
         StrSql = StrSql & " OBJPRTY.estado    = Estados.objid and "
         StrSql = StrSql & " OBJETOS.codobjeto = OBJPRTY.codobjeto  and "
         StrSql = StrSql & " (objetos.objun =  " & session("UN") & " or "
         StrSql = StrSql & "  objetos.objun =  1 )  and "
         StrSql = StrSql & " RELAOBJ.codobjmaster = " & essCodigo
         StrSql = StrSql & " and relaobj.tipobjmaster = '" & essTipo & "' " & sVerEstado
         StrSql = StrSql & " order by OBJETOS.tipobjeto, OBJETOS.descri, Estados.estado "
     ''   response.write "<br>" & strsql & "<br>"
         sVino = "N"
         Set rs2 = OFV.crearconsultaEx(StrSql, cn, 1, "")
         If  not rs2.EOF Then

        Do Until rs2.EOF

             if cstr(rs2(0) & "")  <> cstr(scodigo & "") then

             sVinook = "N"
             if  cstr(rs2(12) & "") = cstr(session("UN") & "") then
   '' response.write "<br><br>entro"
                 sVino   = "S"
                 sVinook = "S"
             end if

          if SESSION("EMPMUN") = "S" THEN
             StrSql = "Select OBJMEMPALC.CODEMPRE  from OBJMEMPALC "
             StrSql = StrSql & " where "
             StrSql = StrSql & " OBJMEMPALC.codobjeto = " & rs2(0) & " and "
             StrSql = StrSql & " OBJMEMPALC.CODEMPRE = " & Session("UN") & " "
             Set rsmas2 = OFV.crearconsultaEx(StrSql, cn, 1, "")
             If  Not rsmas2.EOF Then
                 sVino   = "S"
                 sVinook = "S"
             end if
          end if

          sNIVE   = "D"

        if  sVinook = "S" then

            stitem = rs2(4)
            StrSql = "Select archivo, descri, perfil, TIPOCPT From Tipobjeto "
            StrSql = StrSql & "Where Tipobjeto.tipobjeto = '" & rs2(4) & "'"
            Set rsIcono = OFV.crearconsultaEx(StrSql, cn, 1, "")
            If Not rsIcono.EOF Then
                sIcono  = rsIcono(0)
                sTipoE  = rsIcono(1)
                sPerfil = rsIcono(2)
                sNIVE   = rsIcono(3)

                sitem = sTipoE & ": " & rs2(1)

            End If
            Call OFV.cerrarconsulta(rsIcono)

            if sNIVE   = "D"  then




            else


               TSCPT = TSCPT + 1

               ATABSCPT(TSCPT,0) = ""
               ATABSCPT(TSCPT,1) = rs2(0)
               ATABSCPT(TSCPT,2) = rs2(2)


            end if


         end if

        end if

         rs2.movenext
         loop

        end if



         DO WHILE NOT ISCPT > TSCPT

            IF ISCPT = 0 THEN ISCPT = 1


         FOR ISCPT=ISCPT TO TSCPT



         StrSql = "SELECT DISTINCT OBJETOS.codobjeto, OBJETOS.descri, OBJETOS.tipobjeto, Estados.estado, "
         StrSql = StrSql & "Relaobj.tipobjasoc, OBJPRTY.fecemis, OBJPRTY.fecvigen, OBJPRTY.fecbaja,"
         StrSql = StrSql & " OBJPRTY.autor,OBJPRTY.aprobo, OBJPRTY.elimino, OBJPRTY.fecvto, OBJETOS.OBJUN"
         StrSql = StrSql & " FROM RELAOBJ, OBJETOS, OBJPRTY, Estados "
         StrSql = StrSql & " where "
         StrSql = StrSql & " RELAOBJ.codojasoc = OBJETOS.codobjeto and "
         StrSql = StrSql & " OBJPRTY.estado    = Estados.objid and "
         StrSql = StrSql & " OBJETOS.codobjeto = OBJPRTY.codobjeto  and "
         StrSql = StrSql & " (objetos.objun =  " & session("UN") & " or "
         StrSql = StrSql & "  objetos.objun =  1 )  and "
         StrSql = StrSql & " RELAOBJ.codobjmaster = 0" & ATABSCPT(ISCPT,1)
         StrSql = StrSql & " and relaobj.tipobjmaster = '" & ATABSCPT(ISCPT,2) & "' " & sVerEstado
         StrSql = StrSql & " order by OBJETOS.tipobjeto, OBJETOS.descri, Estados.estado "
    ''    response.write "<br>" & strsql & "<br>"
         sVino = "N"
         Set rs2 = OFV.crearconsultaEx(StrSql, cn, 1, "")
         If  not rs2.EOF Then

         Do Until rs2.EOF

             if cstr(rs2(0) & "")  <> cstr(scodigo & "") then



             sVinook = "N"
             if  cstr(rs2(12) & "") = cstr(session("UN") & "") then
   '' response.write "<br><br>entro"
                 sVino   = "S"
                 sVinook = "S"
             end if

          if SESSION("EMPMUN") = "S" THEN
             StrSql = "Select OBJMEMPALC.CODEMPRE  from OBJMEMPALC "
             StrSql = StrSql & " where "
             StrSql = StrSql & " OBJMEMPALC.codobjeto = " & rs2(0) & " and "
             StrSql = StrSql & " OBJMEMPALC.CODEMPRE = " & Session("UN") & " "
             Set rsmas2 = OFV.crearconsultaEx(StrSql, cn, 1, "")
             If  Not rsmas2.EOF Then
                 sVino   = "S"
                 sVinook = "S"
             end if
          end if

          sNIVE   = "D"

        if  sVinook = "S" then

            stitem = rs2(4)
            StrSql = "Select archivo, descri, perfil, TIPOCPT From Tipobjeto "
            StrSql = StrSql & "Where Tipobjeto.tipobjeto = '" & rs2(4) & "'"
            Set rsIcono = OFV.crearconsultaEx(StrSql, cn, 1, "")
            If Not rsIcono.EOF Then
                sIcono  = rsIcono(0)
                sTipoE  = rsIcono(1)
                sPerfil = rsIcono(2)
                sNIVE   = rsIcono(3)

                sitem = sTipoE & ": " & rs2(1)

            End If
            Call OFV.cerrarconsulta(rsIcono)

            if sNIVE   = "D" then


               sSolapa =  sitem

               If ofv.ModTitCls = "1" Or ofv.ModTitCls = "0" Then
                  sInput = replace(Ucase(sSolapa),"*"," ")
               else
                  sInput = "&nbsp;" & replace(Ucase(sSolapa),"*","&nbsp;") & "&nbsp;"
               End If

               sFuncionDSP = REPLACE(sFuncion,"NNNN",rs2(0))
               sFuncionDSP = REPLACE(sFuncionDSP,"TTTT",rs2(2))

               t3 = xsimagend & "<a  href='javascript:void(0);' " & sFuncionDSP &  " >"
               t3 = t3 & sInput
               t3 = t3 & "</A>"

               t4 = ofv.GenCelda("",xsclase,"","","","","",t3,"si")

               r1 = r1 & ofv.GenRow("","","","","20","","",t4,"si")

            else



               TSCPT = TSCPT + 1

               ATABSCPT(TSCPT,0) = ""
               ATABSCPT(TSCPT,1) = rs2(0)
               ATABSCPT(TSCPT,2) = rs2(2)


            end if


         end if


         end if

         rs2.movenext
         loop

         end if


         NEXT

         LOOP



       end if



       end if


    else

''''' documentos o subcarpetas independientes


     end if


''''''''''''''''''''''''''''


'''''''''''''''''''''''''''''



 rs.movenext
 loop

end if

''''''''''''''''''''''''''''''''''''''''''''


''-----------------------------------------------------------------------------------------------------


END IF

           sHTML = sHTML & "<div id='MDBEDL' style='display:none;height:240px;overflow:auto;" & WStyl & "' >"

        IF  Session("FULLGUI") = "N" THEN
            styl = "background-color:#888888;"
        else
            styl = XspEx("#888888","#cccccc",0) & ";"
        end if

            styl = ""

        If ofv.ModTitCls = "1" Then

           sHTML = sHTML & OFV.GenTabla("DBEDL", "aaa style='border-style:solid;border-color:#ffffff;border-width:1;" & styl & "' ", "", "100%", "", "0", "", "-2", "-2", "", "", r1, "si")

       ''    sHTML = sHTML & ofv.GenTabla("DBED","","","90%","",0,0,-2,-2,"","",r1,"si")

        else

           sHTML = sHTML & OFV.GenTabla("DBEDL", "aaa style='border-style:solid;border-color:#ffffff;border-width:1;" & styl & "' ", "", "100%", "", "0", "", "-2", "-2", "", "", r1, "si")
        ''   sHTML = sHTML & ofv.GenTabla("DBED","","","90%","",0,0,2,2,"","",r1,"si")

        end if

           sHTML = sHTML & "</div>"

  sHTML = sHTML & sHTML1

  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"


end if

Response.Write sHTML

%>
