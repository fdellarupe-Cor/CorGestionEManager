<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/CorInit.inc" -->
<!-- #INCLUDE FILE="IncFile/CorSDK.inc" -->
<%

  set oIni = server.createobject("CorSecurityTCDC.Inicio")

  Set oIni.Request = Request
  Set oIni.Server = Server
  Set oIni.Response = Response
  Set oIni.Session = Session
  Set oIni.Ofv = Ofv
      oIni.Modal = Session("Modal")

  session("noved") = "NO"

  Opcion = Request.querystring("Opcion")

  SESSION("VIEWOPC") = Opcion

  xcinit = corinit() 

  SESSION("CLIEMP") = 0

  SESSION("USREXPIR") = "D"
   

  sHTML = DIRECTLOG()

  Conset(session("CONSET"))
  ofv.wstyles = CorStyler()
  session("TINI") = "I"

  SHTML = "0"

  session("TINI") = "Z"




'''''''''''''' CONFIGURA EMPRESA DEL USUARIO
      IF SESSION("CLIEMP") = 0 THEN
           SESSION("SETEMP") = "N"
           set cn = ofv.conectar(ofv.strconn2)

           StrSql = "Select SITE From USUARIOS WHERE ID = '" & SESSION("USUARIO") & "' "
           Set rs = OFV.crearconsultaEx(StrSql, cn, 1, "")
           IF NOT RS.EOF THEN
              SESSION("CLIEMP") = clng(RS(0))
           ELSE
              SESSION("CLIEMP") = 1
           END IF

           StrSql = "Select * From Clienteplan WHERE ID = " & SESSION("CLIEMP")
           Set rs = OFV.crearconsultaEx(StrSql, cn, 1, "")
           IF NOT RS.EOF THEN
              Session("Cliente") = rs(1)
              Session("CliLogo") = "../COREss20/Empresa/" & rs(5)
              if (rs(11) & "" ) <> "" then Session("CliLogoP") = "../COREss20/Empresa/" & rs(11)
              Session("USRSET")  = RS(10)
           END IF



           call ofv.cerrarconn(cn)
      END IF
''''''''''''''
   WKINITASP = "../COREMANAGER/CorINDEXLISTA.asp"
   session("SMPLMENU") = "S"
   Session("CorApli") = "COREM20"

          
      Response.redirect WKINITASP
          
          
          


FUNCTION DIRECTLOG() 

        vu = OFV.ValidarUsuario()

        vu = "OK"

        'usuario=UserLogon()

        usuario="AOYM4"
        call getInfoUser(Usuario)

        Session("Usuario")   = session("SGU_userid")
        Session("UsrOpcion") = session("SGU_nombre")
        Session("SkrDefT")   = "0"
        Session("SkrDefE")   = "0"
        Session("SkrDefI")   = "0"
        Session("UsrPmin")   = "0"
        Session("UsrPmax")   = "0"
        Session("UsrHist")   = "0"
        Session("SegUsuario") = session("SGU_nombre")
        Session("SegNivel")  = cdbl("8")
        Session("SegGrp")    = "DEF"
        Session("SegPerfil") = Session("Usuario")
        Session("UsuarNT")   = Session("Usuario")

        Session("UsrSector") = session("SGU_idcostcenter")
        Session("UsrSecDsc") = session("SGU_descrisector")
        Session("UsrPuesto") = session("SGU_codigopuesto")
        Session("UsrPstDsc") = session("SGU_descripuesto")
        Session("UsrMail")   = session("SGU_mail")

        call getGroupsUser(Usuario)

        Session("SegGrpOpt")    = "S"

        Session("SegGrpLst")    = session("SGU_grupos")




END FUNCTION






%>
