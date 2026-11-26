<%


DIM strBase
DIM strBase2
DIM userN
DIM pass
DIM ZZ
Dim Srbac
dim sMailLdap

Srbac=session("Srbac")
sMailLdap=session("sMailLdap")
'response.write REQUEST.servervariables("LOGON_USER") & "<br>"
PageSizead=100
DOMC= REQUEST.servervariables("LOGON_USER")
DOMC=SPLIT(DOMC,"\")

if ubound(DOMC) >=0 THEN
SESSION("DOMINIOX")=DOMC(0)
ELSE
SESSION("DOMINIOX")=""
END IF

'SESSION("DOMINIOX")="DOMINIO SF"

'SESSION("DOMINIOX")="DOMINIO ER"
'SESSION("DOMINIOX")="DOMINIO SJ"
'SESSION("DOMINIOX")="DOMINIO SC"


'response.write session("dominioX") & "222"
   XXX=ValDominio(session("dominioX"))

   strBase =  session("ldap")
   strBase2 =REPLACE(strBase & "" ,"<","")
   strBase2=REPLACE(strBase2 & "" ,">","")

   userN = session("usrldap")
   pass = session("pwdldap")



FUNCTION getUsrIngroupAD(grupo)
   DIM ADS_SCOPE_BASE
   ADS_SCOPE_BASE = 2 'subtree



   SET cnAD = ofv.conectar(ofv.strconn2)
   Strsql = "Select MapGrp From grupos where id='" & grupo & "'"
   SET rsADX = ofv.crearconsultaEx(StrSql, cnAd, 1, parametros)
      IF NOT rsADX.EOF THEN
          GrupoADX=rsADX(0)
         GrupoADX = FullNameGrupo(rsADX(0))
      ELSE
         GrupoADX = ""
      END IF

   CALL ofv.cerrarconsulta(rsADX)


   Strsqlq = "Select * From dominios WHERE RAIZLDAP IS NOT NULL OR DESCRIPCION <> 'DEFAULT'"
   SET rsdom = ofv.crearconsultaEx(StrSqlq, cnAD, 1, parametros)
   session("SGU_UIGuserid") = ""
   session("SGU_UIGdescri") = ""

  do while (not rsdom.eof )
    sldap =rsdom("raizldap")
    usernx =  rsdom("usrldap")
    passx = cstr(ofv.getCript("D",rsdom("pwdldap")&""))

    strBase =  rsdom("raizldap")
    strBase2 =REPLACE(strBase ,"<","")
    strBase2=REPLACE(strBase2 ,">","")

   SET objConnection = server.CREATEOBJECT("ADODB.Connection")
   SET objCommand = server.CREATEOBJECT("ADODB.Command")
   objConnection.Provider = "ADsDSOObject"
   objConnection.Properties("User ID") = usernx
   objConnection.Properties("Password") = passx
   objConnection.Properties("Encrypt Password") = TRUE
   objConnection.OPEN "Active Directory Provider"
   SET objCommand.ActiveConnection = objConnection

   objCommand.Properties("Page Size") = PageSizead
   objCommand.Properties("Searchscope") = ADS_SCOPE_BASE

   objCommand.CommandText = "SELECT sAMAccountName,name,cn,title,mail FROM '"& strBase2 &"' WHERE " & "objectCategory='user' and memberof='"& GrupoADX &"'"
   SET objRecordSet = objCommand.EXECUTE


   DO UNTIL objRecordSet.EOF

      session("SGU_UIGuserid") = session("SGU_UIGuserid") & xgusep & objRecordSet("sAMAccountName")
      session("SGU_UIGdescri") = session("SGU_UIGdescri") & xgusep & objRecordSet("name")
      xgusep = CHR(9)

      objRecordSet.MoveNext

   LOOP


   objRecordSet.CLOSE
   objRecordSet = NOTHING
   objConnection.CLOSE
   SET objCommand = NOTHING
   SET objConnectionn = NOTHING
   rsdom.movenext
loop
   call ofv.cerrarconsulta(rsdom)
   CALL ofv.cerrarconn(cnAD)

END FUNCTION

'--------------------------------------------------------------------------------------------------
FUNCTION getUsrIngroupADSV(grupo)
    DIM ADS_SCOPE_BASE
     ADS_SCOPE_BASE = 2 'subtree



     SET cnAD = ofv.conectar(ofv.strconn2)
     Strsql = "Select MapGrp From grupos where id='" & grupo & "'"
     SET rsADX = ofv.crearconsultaEx(StrSql, cnAd, 1, parametros)
        IF NOT rsADX.EOF THEN
            GrupoADX=rsADX(0)
            grupoadx2=  rsADX(0)
            GrupoADX = FullNameGrupo(rsADX(0),strbase)
        ELSE
           GrupoADX = ""
        END IF

     CALL ofv.cerrarconsulta(rsADX)


     Strsqlq = "Select * From dominios WHERE RAIZLDAP IS NOT NULL and DESCRIPCION <> 'DEFAULT'"
     SET rsdom = ofv.crearconsultaEx(StrSqlq, cnAD, 1, parametros)
     session("SGU_UIGuserid") = ""
     session("SGU_UIGdescri") = ""

    do while (not rsdom.eof )

     sldap =rsdom("raizldap")
     usernx =  rsdom("usrldap")
     passx = cstr(ofv.getCript("D",rsdom("pwdldap")&"")) 'rsdom("pwdldap")

     strBase =  rsdom("raizldap")
     strBase2 =REPLACE(strBase ,"<","")
     strBase2=REPLACE(strBase2 ,">","")
     GrupoADX =  FullNameGrupo(GrupoADX2,strBase )
     SET objConnection = server.CREATEOBJECT("ADODB.Connection")
     SET objCommand = server.CREATEOBJECT("ADODB.Command")
     objConnection.Provider = "ADsDSOObject"
     objConnection.Properties("User ID") = usernx
     objConnection.Properties("Password") = passx
     objConnection.Properties("Encrypt Password") = TRUE
     objConnection.OPEN "Active Directory Provider"
     SET objCommand.ActiveConnection = objConnection

     objCommand.Properties("Page Size") = PageSizead
     objCommand.Properties("Searchscope") = ADS_SCOPE_BASE

     objCommand.CommandText = "SELECT sAMAccountName,name,cn,title,mail FROM '"& strBase2 &"' WHERE " & "objectCategory='user' and memberof='"& GrupoADX &"'  and userAccountControl  <> 66050 and userAccountControl  <> 546"

     SET objRecordSet = objCommand.EXECUTE


     DO UNTIL objRecordSet.EOF

        session("SGU_UIGuserid") = session("SGU_UIGuserid") & xgusep & objRecordSet("sAMAccountName")
        session("SGU_UIGdescri") = session("SGU_UIGdescri") & xgusep & objRecordSet("name")
        xgusep = CHR(9)

        objRecordSet.MoveNext

     LOOP


     objRecordSet.CLOSE
'     objRecordSet = NOTHING
     objConnection.CLOSE
     SET objCommand = NOTHING
     SET objConnectionn = NOTHING
     rsdom.movenext
  loop
     call ofv.cerrarconsulta(rsdom)
     CALL ofv.cerrarconn(cnAD)

END FUNCTION


FUNCTION getUsrIngroupADLacaja(grupo)
    DIM ADS_SCOPE_BASE
     ADS_SCOPE_BASE = 2 'subtree



     SET cnAD = ofv.conectar(ofv.strconn2)
     Strsql = "Select MapGrp From grupos where id='" & grupo & "'"
     SET rsADX = ofv.crearconsultaEx(StrSql, cnAd, 1, parametros)
        IF NOT rsADX.EOF THEN
            GrupoADX=rsADX(0)
            grupoadx2=  rsADX(0)
            GrupoADX = FullNameGrupo(rsADX(0),strbase)
        ELSE
           GrupoADX = ""
        END IF

     CALL ofv.cerrarconsulta(rsADX)


     Strsqlq = "Select * From dominios WHERE RAIZLDAP IS NOT NULL and DESCRIPCION <> 'DEFAULT'"
     SET rsdom = ofv.crearconsultaEx(StrSqlq, cnAD, 1, parametros)
     session("SGU_UIGuserid") = ""
     session("SGU_UIGdescri") = ""

    do while (not rsdom.eof )

     sldap =rsdom("raizldap")
     usernx =  rsdom("usrldap")
     passx =cstr(ofv.getCript("D",rsdom("pwdldap")&""))' rsdom("pwdldap")

     strBase =  rsdom("raizldap")
     strBase2 =REPLACE(strBase ,"<","")
     strBase2=REPLACE(strBase2 ,">","")
     GrupoADX =  FullNameGrupo(GrupoADX2,strBase )
     SET objConnection = server.CREATEOBJECT("ADODB.Connection")
     SET objCommand = server.CREATEOBJECT("ADODB.Command")
     objConnection.Provider = "ADsDSOObject"
     objConnection.Properties("User ID") = usernx
     objConnection.Properties("Password") = passx
     objConnection.Properties("Encrypt Password") = TRUE
     objConnection.OPEN "Active Directory Provider"
     SET objCommand.ActiveConnection = objConnection

     objCommand.Properties("Page Size") = PageSizead
     objCommand.Properties("Searchscope") = ADS_SCOPE_BASE

     objCommand.CommandText = "SELECT sAMAccountName,name,cn,title,mail FROM '"& strBase2 &"' WHERE " & "objectCategory='user' and memberof='"& GrupoADX &"'"

     SET objRecordSet = objCommand.EXECUTE


     DO UNTIL objRecordSet.EOF

        session("SGU_UIGuserid") = session("SGU_UIGuserid") & xgusep & objRecordSet("sAMAccountName")
        session("SGU_UIGdescri") = session("SGU_UIGdescri") & xgusep & objRecordSet("name")
        xgusep = CHR(9)

        objRecordSet.MoveNext

     LOOP


     objRecordSet.CLOSE
'     objRecordSet = NOTHING
     objConnection.CLOSE
     SET objCommand = NOTHING
     SET objConnectionn = NOTHING
     rsdom.movenext
  loop
     call ofv.cerrarconsulta(rsdom)
     CALL ofv.cerrarconn(cnAD)

END FUNCTION


FUNCTION getUsrIngroupADBANELCO(grupo)
    DIM ADS_SCOPE_BASE
     ADS_SCOPE_BASE = 2 'subtree



     SET cnAD = ofv.conectar(ofv.strconn2)
     Strsql = "Select MapGrp From grupos where id='" & grupo & "'"
     SET rsADX = ofv.crearconsultaEx(StrSql, cnAd, 1, parametros)
        IF NOT rsADX.EOF THEN
            GrupoADX=rsADX(0)
            grupoadx2=  rsADX(0)
            GrupoADX = FullNameGrupo(rsADX(0),strbase)
        ELSE
           GrupoADX = ""
        END IF

     CALL ofv.cerrarconsulta(rsADX)


     Strsqlq = "Select * From dominios WHERE RAIZLDAP IS NOT NULL and DESCRIPCION <> 'DEFAULT'"
     SET rsdom = ofv.crearconsultaEx(StrSqlq, cnAD, 1, parametros)
     session("SGU_UIGuserid") = ""
     session("SGU_UIGdescri") = ""

    do while (not rsdom.eof )

     sldap =rsdom("raizldap")
     usernx =  rsdom("usrldap")
     passx =cstr(ofv.getCript("D",rsdom("pwdldap")&""))' rsdom("pwdldap")

     strBase =  rsdom("raizldap")
     strBase2 =REPLACE(strBase ,"<","")
     strBase2=REPLACE(strBase2 ,">","")
     GrupoADX =  FullNameGrupo(GrupoADX2,strBase )
     SET objConnection = server.CREATEOBJECT("ADODB.Connection")
     SET objCommand = server.CREATEOBJECT("ADODB.Command")
     objConnection.Provider = "ADsDSOObject"
     objConnection.Properties("User ID") = usernx
     objConnection.Properties("Password") = passx
     objConnection.Properties("Encrypt Password") = TRUE
     objConnection.OPEN "Active Directory Provider"
     SET objCommand.ActiveConnection = objConnection

     objCommand.Properties("Page Size") = PageSizead
     objCommand.Properties("Searchscope") = ADS_SCOPE_BASE

     objCommand.CommandText = "SELECT sAMAccountName,name,cn,title,mail FROM '"& strBase2 &"' WHERE " & "objectCategory='user' and memberof='"& GrupoADX &"'"

     SET objRecordSet = objCommand.EXECUTE


     DO UNTIL objRecordSet.EOF

        session("SGU_UIGuserid") = session("SGU_UIGuserid") & xgusep & objRecordSet("sAMAccountName")
        session("SGU_UIGdescri") = session("SGU_UIGdescri") & xgusep & objRecordSet("name")
        xgusep = CHR(9)

        objRecordSet.MoveNext

     LOOP


     objRecordSet.CLOSE
'     objRecordSet = NOTHING
     objConnection.CLOSE
     SET objCommand = NOTHING
     SET objConnectionn = NOTHING
     rsdom.movenext
  loop
     call ofv.cerrarconsulta(rsdom)
     CALL ofv.cerrarconn(cnAD)

END FUNCTION



FUNCTION getUsrIngroupAD_FINANSUR(grupo)
    DIM ADS_SCOPE_BASE
     ADS_SCOPE_BASE = 2 'subtree



     SET cnAD = ofv.conectar(ofv.strconn2)
     Strsql = "Select MapGrp From grupos where id='" & grupo & "'"
     SET rsADX = ofv.crearconsultaEx(StrSql, cnAd, 1, parametros)
        IF NOT rsADX.EOF THEN
            GrupoADX=rsADX(0)
            grupoadx2=  rsADX(0)
            GrupoADX = FullNameGrupo(rsADX(0),strbase)
        ELSE
           GrupoADX = ""
        END IF

     CALL ofv.cerrarconsulta(rsADX)


     Strsqlq = "Select * From dominios WHERE RAIZLDAP IS NOT NULL and DESCRIPCION <> 'DEFAULT'"
     SET rsdom = ofv.crearconsultaEx(StrSqlq, cnAD, 1, parametros)
     session("SGU_UIGuserid") = ""
     session("SGU_UIGdescri") = ""

    do while (not rsdom.eof )

     sldap =rsdom("raizldap")
     usernx =  rsdom("usrldap")
     passx =cstr(ofv.getCript("D",rsdom("pwdldap")&""))' rsdom("pwdldap")

     strBase =  rsdom("raizldap")
     strBase2 =REPLACE(strBase ,"<","")
     strBase2=REPLACE(strBase2 ,">","")
     GrupoADX =  FullNameGrupo(GrupoADX2,strBase )
     SET objConnection = server.CREATEOBJECT("ADODB.Connection")
     SET objCommand = server.CREATEOBJECT("ADODB.Command")
     objConnection.Provider = "ADsDSOObject"
     objConnection.Properties("User ID") = usernx
     objConnection.Properties("Password") = passx
     objConnection.Properties("Encrypt Password") = TRUE
     objConnection.OPEN "Active Directory Provider"
     SET objCommand.ActiveConnection = objConnection

     objCommand.Properties("Page Size") = PageSizead
     objCommand.Properties("Searchscope") = ADS_SCOPE_BASE

     objCommand.CommandText = "SELECT sAMAccountName,name,cn,title,mail FROM '"& strBase2 &"' WHERE " & "objectCategory='user' and memberof='"& GrupoADX &"'"

     SET objRecordSet = objCommand.EXECUTE


     DO UNTIL objRecordSet.EOF

        session("SGU_UIGuserid") = session("SGU_UIGuserid") & xgusep & objRecordSet("sAMAccountName")
        session("SGU_UIGdescri") = session("SGU_UIGdescri") & xgusep & objRecordSet("name")
        xgusep = CHR(9)

        objRecordSet.MoveNext

     LOOP


     objRecordSet.CLOSE
'     objRecordSet = NOTHING
     objConnection.CLOSE
     SET objCommand = NOTHING
     SET objConnectionn = NOTHING
     rsdom.movenext
  loop
     call ofv.cerrarconsulta(rsdom)
     CALL ofv.cerrarconn(cnAD)

END FUNCTION


FUNCTION getUsrIngroupAD_BPN(grupo)
    DIM ADS_SCOPE_BASE
     ADS_SCOPE_BASE = 2 'subtree



     SET cnAD = ofv.conectar(ofv.strconn2)
     Strsql = "Select MapGrp From grupos where id='" & grupo & "'"
     SET rsADX = ofv.crearconsultaEx(StrSql, cnAd, 1, parametros)
        IF NOT rsADX.EOF THEN
            GrupoADX=rsADX(0)
            grupoadx2=  rsADX(0)
            GrupoADX = FullNameGrupo(rsADX(0),strbase)
        ELSE
           GrupoADX = ""
        END IF

     CALL ofv.cerrarconsulta(rsADX)


     Strsqlq = "Select * From dominios WHERE RAIZLDAP IS NOT NULL and DESCRIPCION <> 'DEFAULT'"
     SET rsdom = ofv.crearconsultaEx(StrSqlq, cnAD, 1, parametros)
     session("SGU_UIGuserid") = ""
     session("SGU_UIGdescri") = ""

    do while (not rsdom.eof )

     sldap =rsdom("raizldap")
     usernx =  rsdom("usrldap")
     passx =cstr(ofv.getCript("D",rsdom("pwdldap")&""))' rsdom("pwdldap")

     strBase =  rsdom("raizldap")
     strBase2 =REPLACE(strBase ,"<","")
     strBase2=REPLACE(strBase2 ,">","")
     GrupoADX =  FullNameGrupo(GrupoADX2,strBase )
     SET objConnection = server.CREATEOBJECT("ADODB.Connection")
     SET objCommand = server.CREATEOBJECT("ADODB.Command")
     objConnection.Provider = "ADsDSOObject"
     objConnection.Properties("User ID") = usernx
     objConnection.Properties("Password") = passx
     objConnection.Properties("Encrypt Password") = TRUE
     objConnection.OPEN "Active Directory Provider"
     SET objCommand.ActiveConnection = objConnection

     objCommand.Properties("Page Size") = PageSizead
     objCommand.Properties("Searchscope") = ADS_SCOPE_BASE

     objCommand.CommandText = "SELECT sAMAccountName,name,cn,title,mail FROM '"& strBase2 &"' WHERE " & "objectCategory='user' and memberof='"& GrupoADX &"'"

     SET objRecordSet = objCommand.EXECUTE


     DO UNTIL objRecordSet.EOF

        session("SGU_UIGuserid") = session("SGU_UIGuserid") & xgusep & objRecordSet("sAMAccountName")
        session("SGU_UIGdescri") = session("SGU_UIGdescri") & xgusep & objRecordSet("name")
        xgusep = CHR(9)

        objRecordSet.MoveNext

     LOOP


     objRecordSet.CLOSE
'     objRecordSet = NOTHING
     objConnection.CLOSE
     SET objCommand = NOTHING
     SET objConnectionn = NOTHING
     rsdom.movenext
  loop
     call ofv.cerrarconsulta(rsdom)
     CALL ofv.cerrarconn(cnAD)

END FUNCTION

FUNCTION getUsrIngroupAD_GP(grupo)
    DIM ADS_SCOPE_BASE
     ADS_SCOPE_BASE = 2 'subtree

   ' RESPONSE.WRITE ofv.strconn2

     SET cnAD = ofv.conectar(ofv.strconn2)
     Strsql = "Select MapGrp From grupos where id='" & grupo & "'"
   '  RESPONSE.WRITE Strsql
     SET rsADX = ofv.crearconsultaEx(StrSql, cnAd, 1, parametros)
        IF NOT rsADX.EOF THEN
            GrupoADX=rsADX(0)
            grupoadx2=  rsADX(0)
            'GrupoADX = FullNameGrupo(rsADX(0),strbase)
        ELSE
           GrupoADX = ""
        END IF

     CALL ofv.cerrarconsulta(rsADX)

     if Session("EMPMUN")  = "S" then
     Strsqlq = "Select * From dominios WHERE cliente=" & session("SGU_site")
     else
     Strsqlq = "Select * From dominios WHERE RAIZLDAP IS NOT NULL and DESCRIPCION <> 'DEFAULT'"
     end if
'response.write Strsqlq
     SET rsdom = ofv.crearconsultaEx(StrSqlq, cnAD, 1, parametros)
     session("SGU_UIGuserid") = ""
     session("SGU_UIGdescri") = ""

    do while (not rsdom.eof )

     sldap =rsdom("raizldap")
     usernx =  rsdom("usrldap")
     passx =cstr(ofv.getCript("D",rsdom("pwdldap")&""))' rsdom("pwdldap")

     strBase =  rsdom("raizldap")
     strBase2 =REPLACE(strBase ,"<","")
     strBase2=REPLACE(strBase2 ,">","")
     GrupoADX =  FullNameGrupo(GrupoADX2,strBase )
     SET objConnection = server.CREATEOBJECT("ADODB.Connection")
     SET objCommand = server.CREATEOBJECT("ADODB.Command")
     objConnection.Provider = "ADsDSOObject"
     objConnection.Properties("User ID") = usernx
     objConnection.Properties("Password") = passx
     objConnection.Properties("Encrypt Password") = TRUE
     objConnection.OPEN "Active Directory Provider"
     SET objCommand.ActiveConnection = objConnection

     objCommand.Properties("Page Size") = PageSizead
     objCommand.Properties("Searchscope") = ADS_SCOPE_BASE

     objCommand.CommandText = "SELECT sAMAccountName,name,cn,title,mail FROM '"& strBase2 &"' WHERE " & "objectCategory='user' and memberof='"& GrupoADX &"'"
    ' RESPONSE.WRITE "SELECT sAMAccountName,name,cn,title,mail FROM '"& strBase2 &"' WHERE " & "objectCategory='user' and memberof='"& GrupoADX &"'"
     SET objRecordSet = objCommand.EXECUTE


     DO UNTIL objRecordSet.EOF

        session("SGU_UIGuserid") = session("SGU_UIGuserid") & xgusep & objRecordSet("sAMAccountName")
        session("SGU_UIGdescri") = session("SGU_UIGdescri") & xgusep & objRecordSet("name")
        xgusep = CHR(9)

        objRecordSet.MoveNext

     LOOP


     objRecordSet.CLOSE
'     objRecordSet = NOTHING
     objConnection.CLOSE
     SET objCommand = NOTHING
     SET objConnectionn = NOTHING
     rsdom.movenext
  loop
     call ofv.cerrarconsulta(rsdom)
     CALL ofv.cerrarconn(cnAD)

END FUNCTION


FUNCTION getUsrIngroupAD_BCR(grupo)
    DIM ADS_SCOPE_BASE
     ADS_SCOPE_BASE = 2 'subtree



     SET cnAD = ofv.conectar(ofv.strconn2)
     Strsql = "Select MapGrp From grupos where id='" & grupo & "'"
     SET rsADX = ofv.crearconsultaEx(StrSql, cnAd, 1, parametros)
        IF NOT rsADX.EOF THEN
            GrupoADX=rsADX(0)
            grupoadx2=  rsADX(0)
            GrupoADX = FullNameGrupo(rsADX(0),strbase)
        ELSE
           GrupoADX = ""
        END IF

     CALL ofv.cerrarconsulta(rsADX)

  IF TRIM(GrupoADX & "")<> "" THEN
     Strsqlq = "Select * From dominios WHERE RAIZLDAP IS NOT NULL and DESCRIPCION <> 'DEFAULT'"
     SET rsdom = ofv.crearconsultaEx(StrSqlq, cnAD, 1, parametros)
     session("SGU_UIGuserid") = ""
     session("SGU_UIGdescri") = ""

    do while (not rsdom.eof )

     sldap =rsdom("raizldap")
     usernx =  rsdom("usrldap")
     passx =cstr(ofv.getCript("D",rsdom("pwdldap")&""))' rsdom("pwdldap")

     strBase =  rsdom("raizldap")
     strBase2 =REPLACE(strBase ,"<","")
     strBase2=REPLACE(strBase2 ,">","")
     GrupoADX =  FullNameGrupo(GrupoADX2,strBase )
     SET objConnection = server.CREATEOBJECT("ADODB.Connection")
     SET objCommand = server.CREATEOBJECT("ADODB.Command")
     objConnection.Provider = "ADsDSOObject"
     objConnection.Properties("User ID") = usernx
     objConnection.Properties("Password") = passx
     objConnection.Properties("Encrypt Password") = TRUE
     objConnection.OPEN "Active Directory Provider"
     SET objCommand.ActiveConnection = objConnection

     objCommand.Properties("Page Size") = PageSizead
     objCommand.Properties("Searchscope") = ADS_SCOPE_BASE

     objCommand.CommandText = "SELECT sAMAccountName,name,cn,title,mail,displayName FROM '"& strBase2 &"' WHERE " & "objectCategory='user' and memberof='"& GrupoADX &"'"

     SET objRecordSet = objCommand.EXECUTE


     DO UNTIL objRecordSet.EOF

        session("SGU_UIGuserid") = session("SGU_UIGuserid") & xgusep & objRecordSet("sAMAccountName")
        session("SGU_UIGdescri") = session("SGU_UIGdescri") & xgusep & objRecordSet("displayName")
        xgusep = CHR(9)

        objRecordSet.MoveNext

     LOOP


     objRecordSet.CLOSE
'     objRecordSet = NOTHING
     objConnection.CLOSE
     SET objCommand = NOTHING
     SET objConnectionn = NOTHING
     rsdom.movenext
  loop
     call ofv.cerrarconsulta(rsdom)
  end if
     CALL ofv.cerrarconn(cnAD)

END FUNCTION


Function GetUserInsectAd(XGU_SECT)

session("SGU_UIGuserid")=""
session("SGU_UIGdescri")=""
IF ISNUMERIC(XGU_SECT) THEN
	   WXGU_SECT = XGU_SECT
	ELSE
	   WXGU_SECT = 1
	END IF

set xcn = ofv.conectar(ofv.strconn2)

xStrSql = "Select codinter From eco WHERE eco = 0" & WXGU_SECT
Set xrs = OFV.crearconsultaEx(xStrSql, xcn, 1, "")

if not xrs.eof then WXGU_SECT= xrs(0) &"*"
'session("xxx")=WXGU_SECT
call ofv.cerrarconsulta(xrs)
call ofv.cerrarconn(xcn)

SET connaXz = server.CREATEOBJECT("ADODB.Connection")


  connaXz.provider = "ADSDSOObject"
  connaXz.Properties("User ID") = usern
  connaXz.Properties("Password") = pass
  connaXz.Properties("Encrypt Password") = TRUE

  connaXz.OPEN

  set oCommand = server.CreateObject("ADODB.Command")
  set oCommand.ActiveConnection = connaXz
  oCommand.Properties("Searchscope")=2


  oCommand.CommandText = "SELECT postOfficeBox,wwwHomePage,sAMAccountName,name FROM '"& strBase2 &"' WHERE "  & "objectCategory='person' and objectclass='user' and wwwHomePage='"& WXGU_SECT  &"' order by postOfficeBox"

  'response.write oCommand.CommandText
  set RSAXz = oCommand.Execute


  'strAttributes = "wwwHomePage,sAMAccountName,name,title,mail,memberOf"

 ' strFilter = "(&(objectCategory=person)(objectClass=user)" & "(wwwHomePage=" & WXGU_SECT &"))" '& "(samAccountName=" & userid &"))"
 ' strQuery = strBase & ";" & strFilter & ";" & strAttributes & ";subtree"
'  response.write strQuery

'  SET RSAXz = connaXz.EXECUTE(strQuery)



  DO UNTIL RSAXz.EOF
	 session("SGU_UIGuserid") = session("SGU_UIGuserid") & xgusep & RSAXz("sAMAccountName")
     session("SGU_UIGdescri") = session("SGU_UIGdescri") & xgusep & RSAXz("name")
	 xgusep = CHR(9)
	 RSAXz.movenext
  LOOP

  RSAXz.CLOSE
  SET RSAXz = NOTHING
  connaXz.CLOSE
  SET connaXz = NOTHING

end function



Function getUserInSectAdBANELCO(XGU_SECT)

session("SGU_UIGuserid")=""
session("SGU_UIGdescri")=""
IF ISNUMERIC(XGU_SECT) THEN
	   WXGU_SECT = XGU_SECT
	ELSE
	   WXGU_SECT = 1
	END IF

set xcn = ofv.conectar(ofv.strconn2)

xStrSql = "Select codinter From eco WHERE eco = 0" & WXGU_SECT
Set xrs = OFV.crearconsultaEx(xStrSql, xcn, 1, "")

if not xrs.eof then WXGU_SECT= xrs(0) &"*"
'session("xxx")=WXGU_SECT
call ofv.cerrarconsulta(xrs)
call ofv.cerrarconn(xcn)

SET connaXz = server.CREATEOBJECT("ADODB.Connection")


  connaXz.provider = "ADSDSOObject"
  connaXz.Properties("User ID") = usern
  connaXz.Properties("Password") = pass
  connaXz.Properties("Encrypt Password") = TRUE

  connaXz.OPEN

  set oCommand = server.CreateObject("ADODB.Command")
  set oCommand.ActiveConnection = connaXz
  oCommand.Properties("Searchscope")=2


  oCommand.CommandText = "SELECT postOfficeBox,wwwHomePage,sAMAccountName,name FROM '"& strBase2 &"' WHERE "  & "objectCategory='person' and objectclass='user' and wwwHomePage='"& WXGU_SECT  &"' order by postOfficeBox"

  'response.write oCommand.CommandText
  set RSAXz = oCommand.Execute


  'strAttributes = "wwwHomePage,sAMAccountName,name,title,mail,memberOf"

 ' strFilter = "(&(objectCategory=person)(objectClass=user)" & "(wwwHomePage=" & WXGU_SECT &"))" '& "(samAccountName=" & userid &"))"
 ' strQuery = strBase & ";" & strFilter & ";" & strAttributes & ";subtree"
'  response.write strQuery

'  SET RSAXz = connaXz.EXECUTE(strQuery)



  DO UNTIL RSAXz.EOF
	 session("SGU_UIGuserid") = session("SGU_UIGuserid") & xgusep & RSAXz("sAMAccountName")
     session("SGU_UIGdescri") = session("SGU_UIGdescri") & xgusep & RSAXz("name")
	 xgusep = CHR(9)
	 RSAXz.movenext
  LOOP

  RSAXz.CLOSE
  SET RSAXz = NOTHING
  connaXz.CLOSE
  SET connaXz = NOTHING

end function




Function getUserInSectAd_FINANSUR(XGU_SECT)

session("SGU_UIGuserid")=""
session("SGU_UIGdescri")=""
IF ISNUMERIC(XGU_SECT) THEN
	   WXGU_SECT = XGU_SECT
	ELSE
	   WXGU_SECT = 1
	END IF

set xcn = ofv.conectar(ofv.strconn2)

xStrSql = "Select codinter From eco WHERE eco = 0" & WXGU_SECT
Set xrs = OFV.crearconsultaEx(xStrSql, xcn, 1, "")

if not xrs.eof then WXGU_SECT= xrs(0) &"*"
'session("xxx")=WXGU_SECT
call ofv.cerrarconsulta(xrs)
call ofv.cerrarconn(xcn)

SET connaXz = server.CREATEOBJECT("ADODB.Connection")


  connaXz.provider = "ADSDSOObject"
  connaXz.Properties("User ID") = usern
  connaXz.Properties("Password") = pass
  connaXz.Properties("Encrypt Password") = TRUE

  connaXz.OPEN

  set oCommand = server.CreateObject("ADODB.Command")
  set oCommand.ActiveConnection = connaXz
  oCommand.Properties("Searchscope")=2


  oCommand.CommandText = "SELECT postOfficeBox,wwwHomePage,sAMAccountName,name FROM '"& strBase2 &"' WHERE "  & "objectCategory='person' and objectclass='user' and wwwHomePage='"& WXGU_SECT  &"' order by postOfficeBox"

  'response.write oCommand.CommandText
  set RSAXz = oCommand.Execute


  'strAttributes = "wwwHomePage,sAMAccountName,name,title,mail,memberOf"

 ' strFilter = "(&(objectCategory=person)(objectClass=user)" & "(wwwHomePage=" & WXGU_SECT &"))" '& "(samAccountName=" & userid &"))"
 ' strQuery = strBase & ";" & strFilter & ";" & strAttributes & ";subtree"
'  response.write strQuery

'  SET RSAXz = connaXz.EXECUTE(strQuery)



  DO UNTIL RSAXz.EOF
	 session("SGU_UIGuserid") = session("SGU_UIGuserid") & xgusep & RSAXz("sAMAccountName")
     session("SGU_UIGdescri") = session("SGU_UIGdescri") & xgusep & RSAXz("name")
	 xgusep = CHR(9)
	 RSAXz.movenext
  LOOP

  RSAXz.CLOSE
  SET RSAXz = NOTHING
  connaXz.CLOSE
  SET connaXz = NOTHING

end function
'---------------------------------------------------------------------------------------------------------------------------------------------



Function getUserInSectAd_BPN(XGU_SECT)

session("SGU_UIGuserid")=""
session("SGU_UIGdescri")=""
IF ISNUMERIC(XGU_SECT) THEN
	   WXGU_SECT = XGU_SECT
	ELSE
	   WXGU_SECT = 1
	END IF

set xcn = ofv.conectar(ofv.strconn2)

xStrSql = "Select codinter From eco WHERE eco = 0" & WXGU_SECT
Set xrs = OFV.crearconsultaEx(xStrSql, xcn, 1, "")

if not xrs.eof then WXGU_SECT= xrs(0) &"*"
'session("xxx")=WXGU_SECT
call ofv.cerrarconsulta(xrs)
call ofv.cerrarconn(xcn)

SET connaXz = server.CREATEOBJECT("ADODB.Connection")


  connaXz.provider = "ADSDSOObject"
  connaXz.Properties("User ID") = usern
  connaXz.Properties("Password") = pass
  connaXz.Properties("Encrypt Password") = TRUE

  connaXz.OPEN

  set oCommand = server.CreateObject("ADODB.Command")
  set oCommand.ActiveConnection = connaXz
  oCommand.Properties("Searchscope")=2


  oCommand.CommandText = "SELECT postOfficeBox,wwwHomePage,sAMAccountName,name FROM '"& strBase2 &"' WHERE "  & "objectCategory='person' and objectclass='user' and wwwHomePage='"& WXGU_SECT  &"' order by postOfficeBox"

  'response.write oCommand.CommandText
  set RSAXz = oCommand.Execute


  'strAttributes = "wwwHomePage,sAMAccountName,name,title,mail,memberOf"

 ' strFilter = "(&(objectCategory=person)(objectClass=user)" & "(wwwHomePage=" & WXGU_SECT &"))" '& "(samAccountName=" & userid &"))"
 ' strQuery = strBase & ";" & strFilter & ";" & strAttributes & ";subtree"
'  response.write strQuery

'  SET RSAXz = connaXz.EXECUTE(strQuery)



  DO UNTIL RSAXz.EOF
	 session("SGU_UIGuserid") = session("SGU_UIGuserid") & xgusep & RSAXz("sAMAccountName")
     session("SGU_UIGdescri") = session("SGU_UIGdescri") & xgusep & RSAXz("name")
	 xgusep = CHR(9)
	 RSAXz.movenext
  LOOP

  RSAXz.CLOSE
  SET RSAXz = NOTHING
  connaXz.CLOSE
  SET connaXz = NOTHING

end function


Function getUserInSectAd_GP(XGU_SECT)

session("SGU_UIGuserid")=""
session("SGU_UIGdescri")=""
IF ISNUMERIC(XGU_SECT) THEN
	   WXGU_SECT = XGU_SECT
	ELSE
	   WXGU_SECT = 1
	END IF

set xcn = ofv.conectar(ofv.strconn2)

xStrSql = "Select codinter From eco WHERE eco = 0" & WXGU_SECT
Set xrs = OFV.crearconsultaEx(xStrSql, xcn, 1, "")

if not xrs.eof then WXGU_SECT= xrs(0) &"*"
'session("xxx")=WXGU_SECT
call ofv.cerrarconsulta(xrs)
call ofv.cerrarconn(xcn)

SET connaXz = server.CREATEOBJECT("ADODB.Connection")


  connaXz.provider = "ADSDSOObject"
  connaXz.Properties("User ID") = usern
  connaXz.Properties("Password") = pass
  connaXz.Properties("Encrypt Password") = TRUE

  connaXz.OPEN

  set oCommand = server.CreateObject("ADODB.Command")
  set oCommand.ActiveConnection = connaXz
  oCommand.Properties("Searchscope")=2


  oCommand.CommandText = "SELECT postOfficeBox,wwwHomePage,sAMAccountName,name FROM '"& strBase2 &"' WHERE "  & "objectCategory='person' and objectclass='user' and wwwHomePage='"& WXGU_SECT  &"' order by postOfficeBox"

  'response.write oCommand.CommandText
  set RSAXz = oCommand.Execute


  'strAttributes = "wwwHomePage,sAMAccountName,name,title,mail,memberOf"

 ' strFilter = "(&(objectCategory=person)(objectClass=user)" & "(wwwHomePage=" & WXGU_SECT &"))" '& "(samAccountName=" & userid &"))"
 ' strQuery = strBase & ";" & strFilter & ";" & strAttributes & ";subtree"
'  response.write strQuery

'  SET RSAXz = connaXz.EXECUTE(strQuery)



  DO UNTIL RSAXz.EOF
	 session("SGU_UIGuserid") = session("SGU_UIGuserid") & xgusep & RSAXz("sAMAccountName")
     session("SGU_UIGdescri") = session("SGU_UIGdescri") & xgusep & RSAXz("name")
	 xgusep = CHR(9)
	 RSAXz.movenext
  LOOP

  RSAXz.CLOSE
  SET RSAXz = NOTHING
  connaXz.CLOSE
  SET connaXz = NOTHING

end function

Function getUserInSectAd_BCR(XGU_SECT)

session("SGU_UIGuserid")=""
session("SGU_UIGdescri")=""
IF ISNUMERIC(XGU_SECT) THEN
	   WXGU_SECT = XGU_SECT
	ELSE
	   WXGU_SECT = 1
	END IF

set xcn = ofv.conectar(ofv.strconn2)

xStrSql = "Select codinter From eco WHERE eco = 0" & WXGU_SECT
Set xrs = OFV.crearconsultaEx(xStrSql, xcn, 1, "")

if not xrs.eof then WXGU_SECT= xrs(0) &"*"
'session("xxx")=WXGU_SECT
call ofv.cerrarconsulta(xrs)
call ofv.cerrarconn(xcn)

SET connaXz = server.CREATEOBJECT("ADODB.Connection")


  connaXz.provider = "ADSDSOObject"
  connaXz.Properties("User ID") = usern
  connaXz.Properties("Password") = pass
  connaXz.Properties("Encrypt Password") = TRUE

  connaXz.OPEN

  set oCommand = server.CreateObject("ADODB.Command")
  set oCommand.ActiveConnection = connaXz
  oCommand.Properties("Searchscope")=2


  oCommand.CommandText = "SELECT postOfficeBox,wwwHomePage,sAMAccountName,name FROM '"& strBase2 &"' WHERE "  & "objectCategory='person' and objectclass='user' and wwwHomePage='"& WXGU_SECT  &"' order by postOfficeBox"

  'response.write oCommand.CommandText
  set RSAXz = oCommand.Execute


  'strAttributes = "wwwHomePage,sAMAccountName,name,title,mail,memberOf"

 ' strFilter = "(&(objectCategory=person)(objectClass=user)" & "(wwwHomePage=" & WXGU_SECT &"))" '& "(samAccountName=" & userid &"))"
 ' strQuery = strBase & ";" & strFilter & ";" & strAttributes & ";subtree"
'  response.write strQuery

'  SET RSAXz = connaXz.EXECUTE(strQuery)



  DO UNTIL RSAXz.EOF
	 session("SGU_UIGuserid") = session("SGU_UIGuserid") & xgusep & RSAXz("sAMAccountName")
     session("SGU_UIGdescri") = session("SGU_UIGdescri") & xgusep & RSAXz("name")
	 xgusep = CHR(9)
	 RSAXz.movenext
  LOOP

  RSAXz.CLOSE
  SET RSAXz = NOTHING
  connaXz.CLOSE
  SET connaXz = NOTHING

end function
'---------------------------------------------------------------------------------------------------------------------------------------------
FUNCTION getUsrIngroupADBCOLU(grupo)
    DIM ADS_SCOPE_BASE
     ADS_SCOPE_BASE = 2 'subtree



     SET cnAD = ofv.conectar(ofv.strconn2)
     Strsql = "Select MapGrp From grupos where id='" & grupo & "'"
     SET rsADX = ofv.crearconsultaEx(StrSql, cnAd, 1, parametros)
        IF NOT rsADX.EOF THEN
            GrupoADX=rsADX(0)
            grupoadx2=  rsADX(0)
            GrupoADX = FullNameGrupo(rsADX(0),strbase)
        ELSE
           GrupoADX = ""
        END IF

     CALL ofv.cerrarconsulta(rsADX)


     Strsqlq = "Select * From dominios WHERE RAIZLDAP IS NOT NULL and DESCRIPCION <> 'DEFAULT'"
     SET rsdom = ofv.crearconsultaEx(StrSqlq, cnAD, 1, parametros)
     session("SGU_UIGuserid") = ""
     session("SGU_UIGdescri") = ""

    do while (not rsdom.eof )

     sldap =rsdom("raizldap")
     usernx =  rsdom("usrldap")
     passx = cstr(ofv.getCript("D",rsdom("pwdldap")&"")) 'rsdom("pwdldap")

     strBase =  rsdom("raizldap")
     strBase2 =REPLACE(strBase ,"<","")
     strBase2=REPLACE(strBase2 ,">","")
     GrupoADX =  FullNameGrupo(GrupoADX2,strBase )
     SET objConnection = server.CREATEOBJECT("ADODB.Connection")
     SET objCommand = server.CREATEOBJECT("ADODB.Command")
     objConnection.Provider = "ADsDSOObject"
     objConnection.Properties("User ID") = usernx
     objConnection.Properties("Password") = passx
     objConnection.Properties("Encrypt Password") = TRUE
     objConnection.OPEN "Active Directory Provider"
     SET objCommand.ActiveConnection = objConnection

     objCommand.Properties("Page Size") = PageSizead
     objCommand.Properties("Searchscope") = ADS_SCOPE_BASE

     objCommand.CommandText = "SELECT sAMAccountName,name,cn,title,mail FROM '"& strBase2 &"' WHERE " & "objectCategory='user' and memberof='"& GrupoADX &"'"

     SET objRecordSet = objCommand.EXECUTE


     DO UNTIL objRecordSet.EOF

        session("SGU_UIGuserid") = session("SGU_UIGuserid") & xgusep & objRecordSet("sAMAccountName")
        session("SGU_UIGdescri") = session("SGU_UIGdescri") & xgusep & objRecordSet("name")
        xgusep = CHR(9)

        objRecordSet.MoveNext

     LOOP


     objRecordSet.CLOSE
'     objRecordSet = NOTHING
     objConnection.CLOSE
     SET objCommand = NOTHING
     SET objConnectionn = NOTHING
     rsdom.movenext
  loop
     call ofv.cerrarconsulta(rsdom)
     CALL ofv.cerrarconn(cnAD)

END FUNCTION







'-----------------------------------------------------------------------------------







FUNCTION getGroupsUserAD(legajo)

   SET connaX = server.CREATEOBJECT("ADODB.Connection")


   connaX.provider = "ADSDSOObject"
   connaX.Properties("User ID") = userN
   connaX.Properties("Password") = pass
   connaX.Properties("Encrypt Password") = TRUE

   connaX.OPEN

   strAttributes = "sAMAccountName,cn,title,mail,memberOf,name"

   strFilter = "(&(objectCategory=person)(objectClass=user)" & "(samAccountName=" & legajo &"))"


   strQuery = strBase & ";" & strFilter & ";" & strAttributes & ";subtree"
'response.write strsql
   SET RSAX = connaX.EXECUTE(strQuery)

'   session("SGU_grupos") = "DEF" & CHR(9)

   session("SGU_grupos")=""
   IF NOT RSAX.EOF THEN

      arrGroups = RSAX.Fields("memberOf")

      IF NOT ISNULL(arrGroups) THEN
         xgusep = ""
         FOR EACH strGroup In arrGroups

            strGroup = UCASE(strGroup &"")
            ' strGroup=replace(strGroup,"CN=","")
            ' xgrupo=split(strGroup,",")

            session("SGU_grupos") = session("SGU_grupos") & xgusep  & strGroup
            xgusep = CHR(9)


         NEXT
      END IF



   END IF

   RSAX.CLOSE
   SET RSAX = NOTHING
   connaX.CLOSE
   SET connaX = NOTHING

END FUNCTION

'--------------------------------------------------------------------------------------------------
FUNCTION getGroupsUserADBCOLU(legajo)

   SET connaX = server.CREATEOBJECT("ADODB.Connection")


   connaX.provider = "ADSDSOObject"
   connaX.Properties("User ID") = userN
   connaX.Properties("Password") = pass
   connaX.Properties("Encrypt Password") = TRUE

   connaX.OPEN

   strAttributes = "sAMAccountName,cn,title,mail,memberOf,name"

   strFilter = "(&(objectCategory=person)(objectClass=user)" & "(samAccountName=" & legajo &"))"


   strQuery = strBase & ";" & strFilter & ";" & strAttributes & ";subtree"

   SET RSAX = connaX.EXECUTE(strQuery)

'   session("SGU_grupos") = "DEF" & CHR(9)


   IF NOT RSAX.EOF THEN

      arrGroups = RSAX.Fields("memberOf")

      IF NOT ISNULL(arrGroups) THEN
         xgusep = ""
         FOR EACH strGroup In arrGroups

            strGroup = UCASE(strGroup &"")
            ' strGroup=replace(strGroup,"CN=","")
             xgrupo=split(strGroup,",")
             sxgrupo=replace(xgrupo(0),"CN=","")
             session("SGU_grupos") = session("SGU_grupos") & xgusep  & sxgrupo
             xgusep = CHR(9)
         NEXT
      END IF



   END IF

   RSAX.CLOSE
   SET RSAX = NOTHING
   connaX.CLOSE
   SET connaX = NOTHING

END FUNCTION


FUNCTION getGroupsUserADLACAJA(legajo)

   SET connaX = server.CREATEOBJECT("ADODB.Connection")


   connaX.provider = "ADSDSOObject"
   connaX.Properties("User ID") = userN
   connaX.Properties("Password") = pass
   connaX.Properties("Encrypt Password") = TRUE

   connaX.OPEN

   strAttributes = "sAMAccountName,cn,title,mail,memberOf,name"

   strFilter = "(&(objectCategory=person)(objectClass=user)" & "(samAccountName=" & legajo &"))"


   strQuery = strBase & ";" & strFilter & ";" & strAttributes & ";subtree"
'response.write strQuery

   SET RSAX = connaX.EXECUTE(strQuery)

'   session("SGU_grupos") = "DEF" & CHR(9)


   IF NOT RSAX.EOF THEN

      arrGroups = RSAX.Fields("memberOf")

      IF NOT ISNULL(arrGroups) THEN
         xgusep = ""
         FOR EACH strGroup In arrGroups

            strGroup = UCASE(strGroup &"")
            ' strGroup=replace(strGroup,"CN=","")
             xgrupo=split(strGroup,",")
             sxgrupo=replace(xgrupo(0),"CN=","")
             session("SGU_grupos") = session("SGU_grupos") & xgusep  & sxgrupo
             xgusep = CHR(9)
         NEXT
      END IF



   END IF

   RSAX.CLOSE
   SET RSAX = NOTHING
   connaX.CLOSE
   SET connaX = NOTHING

END FUNCTION


FUNCTION getGroupsUserADBANELCO(legajo)

   SET connaX = server.CREATEOBJECT("ADODB.Connection")


   connaX.provider = "ADSDSOObject"
   connaX.Properties("User ID") = userN
   connaX.Properties("Password") = pass
   connaX.Properties("Encrypt Password") = TRUE

   connaX.OPEN

   strAttributes = "sAMAccountName,cn,title,mail,memberOf,name"

   strFilter = "(&(objectCategory=person)(objectClass=user)" & "(samAccountName=" & legajo &"))"


   strQuery = strBase & ";" & strFilter & ";" & strAttributes & ";subtree"
'response.write strQuery

   SET RSAX = connaX.EXECUTE(strQuery)

'   session("SGU_grupos") = "DEF" & CHR(9)


   IF NOT RSAX.EOF THEN

      arrGroups = RSAX.Fields("memberOf")

      IF NOT ISNULL(arrGroups) THEN
         xgusep = ""
         FOR EACH strGroup In arrGroups

            strGroup = UCASE(strGroup &"")
            ' strGroup=replace(strGroup,"CN=","")
             xgrupo=split(strGroup,",")
             sxgrupo=replace(xgrupo(0),"CN=","")

             session("SGU_grupos") = session("SGU_grupos") & xgusep  & sxgrupo

             xgusep = CHR(9)
         NEXT
      END IF



   END IF

   RSAX.CLOSE
   SET RSAX = NOTHING
   connaX.CLOSE
   SET connaX = NOTHING

END FUNCTION


FUNCTION getGroupsUserAD_FINANSUR(legajo)

   SET connaX = server.CREATEOBJECT("ADODB.Connection")


   connaX.provider = "ADSDSOObject"
   connaX.Properties("User ID") = userN
   connaX.Properties("Password") = pass
   connaX.Properties("Encrypt Password") = TRUE

   connaX.OPEN

   strAttributes = "sAMAccountName,cn,title,mail,memberOf,name"

   strFilter = "(&(objectCategory=person)(objectClass=user)" & "(samAccountName=" & legajo &"))"


   strQuery = strBase & ";" & strFilter & ";" & strAttributes & ";subtree"
'response.write strQuery

   SET RSAX = connaX.EXECUTE(strQuery)

'   session("SGU_grupos") = "DEF" & CHR(9)


   IF NOT RSAX.EOF THEN

      arrGroups = RSAX.Fields("memberOf")

      IF NOT ISNULL(arrGroups) THEN
         xgusep = ""
         FOR EACH strGroup In arrGroups

            strGroup = UCASE(strGroup &"")
            ' strGroup=replace(strGroup,"CN=","")
             xgrupo=split(strGroup,",")
             sxgrupo=replace(xgrupo(0),"CN=","")

             session("SGU_grupos") = session("SGU_grupos") & xgusep  & sxgrupo

             xgusep = CHR(9)
         NEXT
      END IF



   END IF

   RSAX.CLOSE
   SET RSAX = NOTHING
   connaX.CLOSE
   SET connaX = NOTHING

END FUNCTION

'----------------------------------------------------------------------------------------------------------------------------------------------------------------


FUNCTION getGroupsUserAD_BPN(legajo)

   SET connaX = server.CREATEOBJECT("ADODB.Connection")


   connaX.provider = "ADSDSOObject"
   connaX.Properties("User ID") = userN
   connaX.Properties("Password") = pass
   connaX.Properties("Encrypt Password") = TRUE
'response.write pass
   connaX.OPEN

   strAttributes = "sAMAccountName,cn,title,mail,memberOf,name"

   strFilter = "(&(objectCategory=person)(objectClass=user)" & "(samAccountName=" & legajo &"))"


   strQuery = strBase & ";" & strFilter & ";" & strAttributes & ";subtree"
'response.write strQuery

   SET RSAX = connaX.EXECUTE(strQuery)

'   session("SGU_grupos") = "DEF" & CHR(9)


   IF NOT RSAX.EOF THEN

      arrGroups = RSAX.Fields("memberOf")

      IF NOT ISNULL(arrGroups) THEN
         xgusep = ""
         FOR EACH strGroup In arrGroups

            strGroup = UCASE(strGroup &"")
            ' strGroup=replace(strGroup,"CN=","")
             xgrupo=split(strGroup,",")
             sxgrupo=replace(xgrupo(0),"CN=","")

             session("SGU_grupos") = session("SGU_grupos") & xgusep  & sxgrupo

             xgusep = CHR(9)
         NEXT
      END IF



   END IF

   RSAX.CLOSE
   SET RSAX = NOTHING
   connaX.CLOSE
   SET connaX = NOTHING

END FUNCTION

FUNCTION getGroupsUserAD_GP(legajo)

   SET connaX = server.CREATEOBJECT("ADODB.Connection")


   connaX.provider = "ADSDSOObject"
   connaX.Properties("User ID") = userN
   connaX.Properties("Password") = pass
   connaX.Properties("Encrypt Password") = TRUE
'response.write userN
'EXIT FUNCTION
   connaX.OPEN

   strAttributes = "sAMAccountName,cn,title,mail,memberOf,name"

   strFilter = "(&(objectCategory=person)(objectClass=user)" & "(samAccountName=" & legajo &"))"


   strQuery = strBase & ";" & strFilter & ";" & strAttributes & ";subtree"
'response.write strQuery

   SET RSAX = connaX.EXECUTE(strQuery)

'   session("SGU_grupos") = "DEF" & CHR(9)

session("SGU_grupos")=""
   IF NOT RSAX.EOF THEN

      arrGroups = RSAX.Fields("memberOf")

      IF NOT ISNULL(arrGroups) THEN
         xgusep = ""
         FOR EACH strGroup In arrGroups

           ' strGroup = UCASE(strGroup &"")
            ' strGroup=replace(strGroup,"CN=","")
             xgrupo=split(strGroup,",")
             sxgrupo=replace(xgrupo(0),"CN=","")

             session("SGU_grupos") = session("SGU_grupos") & xgusep  & sxgrupo

             xgusep = CHR(9)
         NEXT
      END IF



   END IF

   RSAX.CLOSE
   SET RSAX = NOTHING
   connaX.CLOSE
   SET connaX = NOTHING
'response.write session("SGU_grupos")
END FUNCTION

'----------------------------------------------------------------------------------------------------------------------------------------------------------------
FUNCTION getGroupsUserADSV(legajo)

  SET connaX = server.CREATEOBJECT("ADODB.Connection")


   connaX.provider = "ADSDSOObject"
   connaX.Properties("User ID") = userN
   connaX.Properties("Password") = pass
   connaX.Properties("Encrypt Password") = TRUE

   connaX.OPEN

   strAttributes = "sAMAccountName,cn,title,mail,memberOf,name"

   strFilter = "(&(objectCategory=person)(objectClass=user)" & "(samAccountName=" & legajo &"))"


   strQuery = strBase & ";" & strFilter & ";" & strAttributes & ";subtree"
'response.write pass

   SET RSAX = connaX.EXECUTE(strQuery)

'   session("SGU_grupos") = "DEF" & CHR(9)


   IF NOT RSAX.EOF THEN

      arrGroups = RSAX.Fields("memberOf")

      IF NOT ISNULL(arrGroups) THEN
         xgusep = ""
         FOR EACH strGroup In arrGroups

            strGroup = UCASE(strGroup &"")
            ' strGroup=replace(strGroup,"CN=","")
             xgrupo=split(strGroup,",")
             sxgrupo=replace(xgrupo(0),"CN=","")
             session("SGU_grupos") = session("SGU_grupos") & xgusep  & sxgrupo
             xgusep = CHR(9)
         NEXT
      END IF



   END IF

   RSAX.CLOSE
   SET RSAX = NOTHING
   connaX.CLOSE
   SET connaX = NOTHING


END FUNCTION


FUNCTION getInfouserWLACAJA(legajo)


   SET cnDOM = ofv.conectar(ofv.strconn0)
   Strsqlq = "Select * From dominios WHERE RAIZLDAP IS NOT NULL and DESCRIPCION <> 'DEFAULT'"
    SET rsdom = ofv.crearconsultaEx(StrSqlq, cnDOM, 1, parametros)

  do while (not rsdom.eof )
   sldap =rsdom("raizldap")
   usernx =  rsdom("usrldap")
   passx = cstr(ofv.getCript("D",rsdom("pwdldap")&""))'rsdom("pwdldap")

   SET connaX = server.CREATEOBJECT("ADODB.Connection")


   connaX.provider = "ADSDSOObject"
   connaX.Properties("User ID") = usernx
   connaX.Properties("Password") = passx
   connaX.Properties("Encrypt Password") = TRUE

   connaX.OPEN

   strAttributes = "sAMAccountName,name,cn,title,mail,memberOf"

   strFilter = "(&(objectCategory=person)(objectClass=user)" & "(samAccountName=" & legajo &"))"


   strQuery = sldap & ";" & strFilter & ";" & strAttributes & ";subtree"

   SET RSAX = connaX.EXECUTE(strQuery)


   IF RSAX.EOF THEN


      session("SGU_userid") = Legajo
      session("SGU_nombre") = "NO DISPONIBLE - " & ERR.NUMBER & "&nbsp;" & ERR.DESCRIPTION
      session("SGU_mail") = "NO@disponible.com"
      session("SGU_descrisector") = "NO disponible"
      session("SGU_idcostcenter") = "0"
      session("SGU_descripuesto") = "NO disponible"
      session("SGU_codigopuesto") = "NO"


   ELSE

      session("SGU_userid") = Legajo
      session("SGU_nombre") = RSAX("name")
      session("SGU_mail") = legajo & "@lacaja.com.ar" 'RSAX("mail")
      session("SGU_descrisector") = "NO"
      session("SGU_idcostcenter") = "0"
      session("SGU_descripuesto") = RSAX("title")
      session("SGU_codigopuesto") = "NO"
      exit do




   END IF

   RSAX.CLOSE
   SET RSAX = NOTHING
   connaX.CLOSE
   SET connaX = NOTHING
   rsdom.MOVENEXT
   LOOP
   CALL ofv.cerrarconsulta(rsdom)
   CALL ofv.cerrarconn(cnDOM)
END FUNCTION



FUNCTION getInfouserWBANELCO(legajo)


   SET cnDOM = ofv.conectar(ofv.strconn0)
   Strsqlq = "Select * From dominios WHERE RAIZLDAP IS NOT NULL and DESCRIPCION <> 'DEFAULT'"
    SET rsdom = ofv.crearconsultaEx(StrSqlq, cnDOM, 1, parametros)

  do while (not rsdom.eof )
   sldap =rsdom("raizldap")
   usernx =  rsdom("usrldap")
   passx = cstr(ofv.getCript("D",rsdom("pwdldap")&""))'rsdom("pwdldap")

   SET connaX = server.CREATEOBJECT("ADODB.Connection")


   connaX.provider = "ADSDSOObject"
   connaX.Properties("User ID") = usernx
   connaX.Properties("Password") = passx
   connaX.Properties("Encrypt Password") = TRUE

   connaX.OPEN

   'strAttributes = "sAMAccountName,name,cn,title,mail,memberOf"
   strAttributes = "sAMAccountName,cn,title,mail,memberOf,name,department,description"

   strFilter = "(&(objectCategory=person)(objectClass=user)" & "(samAccountName=" & legajo &"))"


   strQuery = sldap & ";" & strFilter & ";" & strAttributes & ";subtree"

   SET RSAX = connaX.EXECUTE(strQuery)


   IF RSAX.EOF THEN


      session("SGU_userid") = Legajo
      session("SGU_nombre") = "NO DISPONIBLE - " & ERR.NUMBER & "&nbsp;" & ERR.DESCRIPTION
      session("SGU_mail") = "NO@disponible.com"
      session("SGU_descrisector") = "NO disponible"
      session("SGU_idcostcenter") = "0"
      session("SGU_descripuesto") = "NO disponible"
      session("SGU_codigopuesto") = "NO"
      session("SGU_idnivel")="9"


   ELSE

           session("SGU_userid") = Legajo
		        session("SGU_nombre") = RSAX("name")
		        session("SGU_mail") = RSAX("mail")
		        session("SGU_descrisector") = "NO"
		        session("SGU_idcostcenter") = getsectorExt(rsax("department")&"")
		        session("SGU_descripuesto") = "NO"
		        session("SGU_codigopuesto") = getperfilExtBanelco(rsax("title"))
		        if isarray(rsax("description"))then
		          xnivel=rsax("description")
		          session("SGU_idnivel") =xnivel(0)
		        else
		          xnivel=rsax("description")
		          session("SGU_idnivel") =xnivel
		        end if
                session("SGU_idnivel")=getnivelExtBanelco(ucase(session("SGU_idnivel")))
      exit do




   END IF

   RSAX.CLOSE
   SET RSAX = NOTHING
   connaX.CLOSE
   SET connaX = NOTHING
   rsdom.MOVENEXT
   LOOP
   CALL ofv.cerrarconsulta(rsdom)
   CALL ofv.cerrarconn(cnDOM)
END FUNCTION

FUNCTION getInfouserW_FINANSUR(legajo)


   SET cnDOM = ofv.conectar(ofv.strconn0)
   Strsqlq = "Select * From dominios WHERE RAIZLDAP IS NOT NULL and DESCRIPCION <> 'DEFAULT'"
    SET rsdom = ofv.crearconsultaEx(StrSqlq, cnDOM, 1, parametros)

  do while (not rsdom.eof )
   sldap =rsdom("raizldap")
   usernx =  rsdom("usrldap")
   passx = cstr(ofv.getCript("D",rsdom("pwdldap")&""))'rsdom("pwdldap")

   SET connaX = server.CREATEOBJECT("ADODB.Connection")


   connaX.provider = "ADSDSOObject"
   connaX.Properties("User ID") = usernx
   connaX.Properties("Password") = passx
   connaX.Properties("Encrypt Password") = TRUE

   connaX.OPEN

   'strAttributes = "sAMAccountName,name,cn,title,mail,memberOf"
   strAttributes = "sAMAccountName,cn,title,mail,memberOf,name,department,description"

   strFilter = "(&(objectCategory=person)(objectClass=user)" & "(samAccountName=" & legajo &"))"


   strQuery = sldap & ";" & strFilter & ";" & strAttributes & ";subtree"

   SET RSAX = connaX.EXECUTE(strQuery)


   IF RSAX.EOF THEN


      session("SGU_userid") = Legajo
      session("SGU_nombre") = "NO DISPONIBLE - " & ERR.NUMBER & "&nbsp;" & ERR.DESCRIPTION
      session("SGU_mail") = "NO@disponible.com"
      session("SGU_descrisector") = "NO disponible"
      session("SGU_idcostcenter") = "0"
      session("SGU_descripuesto") = "NO disponible"
      session("SGU_codigopuesto") = "NO"
      session("SGU_idnivel")="9"


   ELSE

           session("SGU_userid") = Legajo
		        session("SGU_nombre") = RSAX("name")
		        session("SGU_mail") = RSAX("mail")
		        session("SGU_descrisector") = "NO"
		        session("SGU_idcostcenter") = "0" 'getsectorExt(rsax("department")&"")
		        session("SGU_descripuesto") = "NO"
		        session("SGU_codigopuesto") ="NO" 'getperfilExtBanelco(rsax("title"))

                session("SGU_idnivel")="9"
      exit do




   END IF

   RSAX.CLOSE
   SET RSAX = NOTHING
   connaX.CLOSE
   SET connaX = NOTHING
   rsdom.MOVENEXT
   LOOP
   CALL ofv.cerrarconsulta(rsdom)
   CALL ofv.cerrarconn(cnDOM)
END FUNCTION


FUNCTION getInfouserW_BPN(legajo)


   SET cnDOM = ofv.conectar(ofv.strconn0)
   Strsqlq = "Select * From dominios WHERE RAIZLDAP IS NOT NULL and DESCRIPCION <> 'DEFAULT'"
    SET rsdom = ofv.crearconsultaEx(StrSqlq, cnDOM, 1, parametros)

  do while (not rsdom.eof )
   sldap =rsdom("raizldap")
   usernx =  rsdom("usrldap")
   passx = cstr(ofv.getCript("D",rsdom("pwdldap")&""))'rsdom("pwdldap")

   SET connaX = server.CREATEOBJECT("ADODB.Connection")


   connaX.provider = "ADSDSOObject"
   connaX.Properties("User ID") = usernx
   connaX.Properties("Password") = passx
   connaX.Properties("Encrypt Password") = TRUE

   connaX.OPEN

   'strAttributes = "sAMAccountName,name,cn,title,mail,memberOf"
   strAttributes = "sAMAccountName,cn,title,mail,memberOf,name,department,description"

   strFilter = "(&(objectCategory=person)(objectClass=user)" & "(samAccountName=" & legajo &"))"


   strQuery = sldap & ";" & strFilter & ";" & strAttributes & ";subtree"

   SET RSAX = connaX.EXECUTE(strQuery)


   IF RSAX.EOF THEN


      session("SGU_userid") = Legajo
      session("SGU_nombre") = "NO DISPONIBLE - " & ERR.NUMBER & "&nbsp;" & ERR.DESCRIPTION
      session("SGU_mail") = "NO@disponible.com"
      session("SGU_descrisector") = "NO disponible"
      session("SGU_idcostcenter") = "0"
      session("SGU_descripuesto") = "NO disponible"
      session("SGU_codigopuesto") = "NO"
      session("SGU_idnivel")="9"


   ELSE

           session("SGU_userid") = Legajo
		        session("SGU_nombre") = RSAX("name")
		        session("SGU_mail") = RSAX("mail")
		        session("SGU_descrisector") = "NO"
		        session("SGU_idcostcenter") = "0" 'getsectorExt(rsax("department")&"")
		        session("SGU_descripuesto") = "NO"
		        session("SGU_codigopuesto") ="NO" 'getperfilExtBanelco(rsax("title"))

                session("SGU_idnivel")="9"
      exit do




   END IF

   RSAX.CLOSE
   SET RSAX = NOTHING
   connaX.CLOSE
   SET connaX = NOTHING
   rsdom.MOVENEXT
   LOOP
   CALL ofv.cerrarconsulta(rsdom)
   CALL ofv.cerrarconn(cnDOM)
END FUNCTION


FUNCTION getInfouserW_GP(legajo)


   SET cnDOM = ofv.conectar(ofv.strconn0)
   Strsqlq = "Select * From dominios WHERE RAIZLDAP IS NOT NULL and DESCRIPCION <> 'DEFAULT'"
    SET rsdom = ofv.crearconsultaEx(StrSqlq, cnDOM, 1, parametros)

  do while (not rsdom.eof )
   sldap =rsdom("raizldap")
   usernx =  rsdom("usrldap")
   passx = cstr(ofv.getCript("D",rsdom("pwdldap")&""))'rsdom("pwdldap")

   SET connaX = server.CREATEOBJECT("ADODB.Connection")


   connaX.provider = "ADSDSOObject"
   connaX.Properties("User ID") = usernx
   connaX.Properties("Password") = passx
   connaX.Properties("Encrypt Password") = TRUE

   connaX.OPEN

   'strAttributes = "sAMAccountName,name,cn,title,mail,memberOf"
   strAttributes = "sAMAccountName,cn,title,mail,memberOf,name,department,description"

   strFilter = "(&(objectCategory=person)(objectClass=user)" & "(samAccountName=" & legajo &"))"


   strQuery = sldap & ";" & strFilter & ";" & strAttributes & ";subtree"

   SET RSAX = connaX.EXECUTE(strQuery)


   IF RSAX.EOF THEN


      session("SGU_userid") = Legajo
      session("SGU_nombre") = "NO DISPONIBLE - " & ERR.NUMBER & "&nbsp;" & ERR.DESCRIPTION
      session("SGU_mail") = "NO@disponible.com"
      session("SGU_descrisector") = "NO disponible"
      session("SGU_idcostcenter") = "0"
      session("SGU_descripuesto") = "NO disponible"
      session("SGU_codigopuesto") = "NO"
      session("SGU_idnivel")="9"
       session("SGU_site")=1

   ELSE

           session("SGU_userid") = Legajo
		        session("SGU_nombre") = RSAX("name")
		        session("SGU_mail") = RSAX("mail")
		        session("SGU_descrisector") = "NO"
		        session("SGU_idcostcenter") = "0" 'getsectorExt(rsax("department")&"")
		        session("SGU_descripuesto") = "NO"
		        session("SGU_codigopuesto") ="NO" 'getperfilExtBanelco(rsax("title"))

                session("SGU_idnivel")="9"
                session("SGU_site")=session("usr_dominio")
      exit do




   END IF

   RSAX.CLOSE
   SET RSAX = NOTHING
   connaX.CLOSE
   SET connaX = NOTHING
   rsdom.MOVENEXT
   LOOP
   CALL ofv.cerrarconsulta(rsdom)
   CALL ofv.cerrarconn(cnDOM)
END FUNCTION


FUNCTION getInfouserw_BCR(legajo)


   SET cnDOM = ofv.conectar(ofv.strconn0)
   Strsqlq = "Select * From dominios WHERE RAIZLDAP IS NOT NULL and DESCRIPCION <> 'DEFAULT'"
    SET rsdom = ofv.crearconsultaEx(StrSqlq, cnDOM, 1, parametros)

  do while (not rsdom.eof )
   sldap =rsdom("raizldap")
   usernx =  rsdom("usrldap")
   passx = cstr(ofv.getCript("D",rsdom("pwdldap")&""))'rsdom("pwdldap")

   SET connaX = server.CREATEOBJECT("ADODB.Connection")


   connaX.provider = "ADSDSOObject"
   connaX.Properties("User ID") = usernx
   connaX.Properties("Password") = passx
   connaX.Properties("Encrypt Password") = TRUE

   connaX.OPEN

   'strAttributes = "sAMAccountName,name,cn,title,mail,memberOf"
   strAttributes = "sAMAccountName,cn,title,mail,memberOf,name,department,description,displayName"

   strFilter = "(&(objectCategory=person)(objectClass=user)" & "(samAccountName=" & legajo &"))"


   strQuery = sldap & ";" & strFilter & ";" & strAttributes & ";subtree"

   SET RSAX = connaX.EXECUTE(strQuery)


   IF RSAX.EOF THEN


      session("SGU_userid") = Legajo
      session("SGU_nombre") = "NO DISPONIBLE - " & ERR.NUMBER & "&nbsp;" & ERR.DESCRIPTION
      session("SGU_mail") = "NO@disponible.com"
      session("SGU_descrisector") = "NO disponible"
      session("SGU_idcostcenter") = "0"
      session("SGU_descripuesto") = "NO disponible"
      session("SGU_codigopuesto") = "NO"
      session("SGU_idnivel")="9"


   ELSE

           session("SGU_userid") = Legajo
		        session("SGU_nombre") = RSAX("displayName")
		        session("SGU_mail") = RSAX("mail")
		        session("SGU_descrisector") = "NO"
		        session("SGU_idcostcenter") = getsectorExtBCR(rsax("department")&"")
		        session("SGU_descripuesto") = "NO"
		        session("SGU_codigopuesto") ="NO" 'getperfilExtBanelco(rsax("title"))

                session("SGU_idnivel")=getnivelExtBCR(rsax("title")&"")
      exit do




   END IF

   RSAX.CLOSE
   SET RSAX = NOTHING
   connaX.CLOSE
   SET connaX = NOTHING
   rsdom.MOVENEXT
   LOOP
   CALL ofv.cerrarconsulta(rsdom)
   CALL ofv.cerrarconn(cnDOM)
END FUNCTION


FUNCTION getInfouserWAD(legajo)

 SET cnDOM = ofv.conectar(ofv.strconn0)
   Strsqlq = "Select * From dominios WHERE RAIZLDAP IS NOT NULL and DESCRIPCION <> 'DEFAULT' order by id"
    SET rsdom = ofv.crearconsultaEx(StrSqlq, cnDOM, 1, parametros)

  do while (not rsdom.eof )
   sldap =rsdom("raizldap")
   usernx =  rsdom("usrldap")
   passx =cstr(ofv.getCript("D",rsdom("pwdldap")&""))' rsdom("pwdldap")

   SET connaX = server.CREATEOBJECT("ADODB.Connection")


   connaX.provider = "ADSDSOObject"
   connaX.Properties("User ID") = usernx
   connaX.Properties("Password") = passx
   connaX.Properties("Encrypt Password") = TRUE

   connaX.OPEN

   strAttributes = "sAMAccountName,name,cn,title,mail,memberOf"

   strFilter = "(&(objectCategory=person)(objectClass=user)" & "(samAccountName=" & legajo &"))"


   strQuery = sldap & ";" & strFilter & ";" & strAttributes & ";subtree"

   SET RSAX = connaX.EXECUTE(strQuery)


   IF RSAX.EOF THEN


      session("SGU_userid") = Legajo
      session("SGU_nombre") = "NO DISPONIBLE - " & ERR.NUMBER & "&nbsp;" & ERR.DESCRIPTION
      session("SGU_mail") = "NO@disponible.com"
      session("SGU_descrisector") = "NO disponible"
      session("SGU_idcostcenter") = "0"
      session("SGU_descripuesto") = "NO disponible"
      session("SGU_codigopuesto") = "NO"


   ELSE

      session("SGU_userid") = Legajo
      session("SGU_nombre") = RSAX("name")
      session("SGU_mail") = RSAX("mail")
      session("SGU_descrisector") = "NO"
      session("SGU_idcostcenter") = "0"
      session("SGU_descripuesto") = RSAX("title")
      session("SGU_codigopuesto") = "NO"
      exit do




   END IF

   RSAX.CLOSE
   SET RSAX = NOTHING
   connaX.CLOSE
   SET connaX = NOTHING
   rsdom.MOVENEXT
   LOOP
   CALL ofv.cerrarconsulta(rsdom)
   CALL ofv.cerrarconn(cnDOM)
END FUNCTION

'------------------------------------------------------------------------------------------------------------------------------------------------
FUNCTION getInfouserWBCOLU(legajo)


   SET cnDOM = ofv.conectar(ofv.strconn0)
   Strsqlq = "Select * From dominios WHERE RAIZLDAP IS NOT NULL and DESCRIPCION <> 'DEFAULT'"
    SET rsdom = ofv.crearconsultaEx(StrSqlq, cnDOM, 1, parametros)

  do while (not rsdom.eof )
   sldap =rsdom("raizldap")
   usernx =  rsdom("usrldap")
   passx = cstr(ofv.getCript("D",rsdom("pwdldap")&""))'rsdom("pwdldap")

   SET connaX = server.CREATEOBJECT("ADODB.Connection")


   connaX.provider = "ADSDSOObject"
   connaX.Properties("User ID") = usernx
   connaX.Properties("Password") = passx
   connaX.Properties("Encrypt Password") = TRUE

   connaX.OPEN

   strAttributes = "sAMAccountName,name,cn,title,mail,memberOf"

   strFilter = "(&(objectCategory=person)(objectClass=user)" & "(samAccountName=" & legajo &"))"


   strQuery = sldap & ";" & strFilter & ";" & strAttributes & ";subtree"

   SET RSAX = connaX.EXECUTE(strQuery)


   IF RSAX.EOF THEN


      session("SGU_userid") = Legajo
      session("SGU_nombre") = "NO DISPONIBLE - " & ERR.NUMBER & "&nbsp;" & ERR.DESCRIPTION
      session("SGU_mail") = "NO@disponible.com"
      session("SGU_descrisector") = "NO disponible"
      session("SGU_idcostcenter") = "0"
      session("SGU_descripuesto") = "NO disponible"
      session("SGU_codigopuesto") = "NO"


   ELSE

      session("SGU_userid") = Legajo
      session("SGU_nombre") = RSAX("name")
      session("SGU_mail") = RSAX("mail")
      session("SGU_descrisector") = "NO"
      session("SGU_idcostcenter") = "0"
      session("SGU_descripuesto") = RSAX("title")
      session("SGU_codigopuesto") = "NO"
      exit do




   END IF

   RSAX.CLOSE
   SET RSAX = NOTHING
   connaX.CLOSE
   SET connaX = NOTHING
   rsdom.MOVENEXT
   LOOP
   CALL ofv.cerrarconsulta(rsdom)
   CALL ofv.cerrarconn(cnDOM)
END FUNCTION
'------------------------------------------------------------------------------------------------------------------------------------------------




FUNCTION getInfouserWSV(legajo)


   SET cnDOM = ofv.conectar(ofv.strconn0)
   Strsqlq = "Select * From dominios WHERE RAIZLDAP IS NOT NULL and DESCRIPCION <> 'DEFAULT'"
    SET rsdom = ofv.crearconsultaEx(StrSqlq, cnDOM, 1, parametros)

  do while (not rsdom.eof )
   sldap =rsdom("raizldap")
   usernx =  rsdom("usrldap")
   passx = cstr(ofv.getCript("D",rsdom("pwdldap")&""))'rsdom("pwdldap")

   SET connaX = server.CREATEOBJECT("ADODB.Connection")


   connaX.provider = "ADSDSOObject"
   connaX.Properties("User ID") = usernx
   connaX.Properties("Password") = passx
   connaX.Properties("Encrypt Password") = TRUE

   connaX.OPEN

   strAttributes = "sAMAccountName,name,cn,title,mail,memberOf"

   strFilter = "(&(objectCategory=person)(objectClass=user)" & "(samAccountName=" & legajo &"))"


   strQuery = sldap & ";" & strFilter & ";" & strAttributes & ";subtree"

   SET RSAX = connaX.EXECUTE(strQuery)


   IF RSAX.EOF THEN


      session("SGU_userid") = Legajo
      session("SGU_nombre") = "NO DISPONIBLE - " & ERR.NUMBER & "&nbsp;" & ERR.DESCRIPTION
      session("SGU_mail") = "NO@disponible.com"
      session("SGU_descrisector") = "NO disponible"
      session("SGU_idcostcenter") = "0"
      session("SGU_descripuesto") = "NO disponible"
      session("SGU_codigopuesto") = "NO"


   ELSE

      session("SGU_userid") = Legajo
      session("SGU_nombre") = RSAX("name")
      session("SGU_mail") = RSAX("mail")
      session("SGU_descrisector") = "NO"
      session("SGU_idcostcenter") = "0"
      session("SGU_descripuesto") = RSAX("title")
      session("SGU_codigopuesto") = "NO"
      exit do




   END IF

   RSAX.CLOSE
   SET RSAX = NOTHING
   connaX.CLOSE
   SET connaX = NOTHING
   rsdom.MOVENEXT
   LOOP
   CALL ofv.cerrarconsulta(rsdom)
   CALL ofv.cerrarconn(cnDOM)
END FUNCTION


FUNCTION getInfouserWBC(legajo)


   SET cnDOM = ofv.conectar(ofv.strconn0)
   Strsqlq = "Select * From dominios WHERE RAIZLDAP IS NOT NULL and DESCRIPCION <> 'DEFAULT'"
    SET rsdom = ofv.crearconsultaEx(StrSqlq, cnDOM, 1, parametros)

  do while (not rsdom.eof )
   sldap =rsdom("raizldap")
   usernx =  rsdom("usrldap")
   passx = cstr(ofv.getCript("D",rsdom("pwdldap")&""))'rsdom("pwdldap")

   SET connaX = server.CREATEOBJECT("ADODB.Connection")


   connaX.provider = "ADSDSOObject"
   connaX.Properties("User ID") = usernx
   connaX.Properties("Password") = passx
   connaX.Properties("Encrypt Password") = TRUE

   connaX.OPEN

   strAttributes = "sAMAccountName,name,cn,title,mail,memberOf"

   strFilter = "(&(objectCategory=person)(objectClass=user)" & "(samAccountName=" & legajo &"))"


   strQuery = sldap & ";" & strFilter & ";" & strAttributes & ";subtree"

   SET RSAX = connaX.EXECUTE(strQuery)


   IF RSAX.EOF THEN


      session("SGU_userid") = Legajo
      session("SGU_nombre") = "NO DISPONIBLE - " & ERR.NUMBER & "&nbsp;" & ERR.DESCRIPTION
      session("SGU_mail") = "NO@disponible.com"
      session("SGU_descrisector") = "NO disponible"
      session("SGU_idcostcenter") = "0"
      session("SGU_descripuesto") = "NO disponible"
      session("SGU_codigopuesto") = "NO"


   ELSE

     set Cnpeople=ofv.conectar(ofv.strconn7)
     legajoZ=mid(legajo,3,len(legajo))
	 strsqlCn= "select ALTER_EMPLID  ,NAME1, DEPTID  ,JOBCODE   from PS_BC_PS_AD_VW1 where ALTER_EMPLID='" & legajoZ &"'"
	 set CnViewPeople=  ofv.crearconsultaEx(strsqlCn ,Cnpeople ,1,volver)

	 if not CnViewPeople.eof then

	  session("SGU_userid") = Legajo
	 session("SGU_nombre") =replace(CnViewPeople("NAME1")& "" ,"'"," ") ' & " " &  CnViewPeople("MIDDLE_NAME") & " " & CnViewPeople("LAST_NAME"),"'"," ")
                session("SGU_mail") = ChangeMail(RSAX("mail"))
	  session("SGU_descrisector") = "NO"
	  SDEPTID = CnViewPeople("DEPTID") & ""
	  SDEPTID=TRIM(SDEPTID)
	  session("SGU_idcostcenter") = getsectorExt(SDEPTID) 'CnViewPeople("DEPTID")
	  session("SGU_descripuesto") = ""
	  SXJOB=CnViewPeople("JOBCODE") & ""
	  SXJOB=TRIM(SXJOB)
	  session("SGU_codigopuesto") = SXJOB
	  session("SGU_idnivel")      =getnivelExt(SXJOB)
	 else

	  session("SGU_userid") = Legajo
	  session("SGU_nombre") = "NO DISPONIBLE - " & ERR.NUMBER & "&nbsp;" & ERR.DESCRIPTION
	  session("SGU_mail") = "NO@disponible.com"
	  session("SGU_descrisector") = "NO disponible"
	  session("SGU_idcostcenter") = "0"
	  session("SGU_descripuesto") = "NO disponible"
	  session("SGU_codigopuesto") = "NO"
	  session("SGU_idnivel")      ="9"
	 end if

	  call ofv.cerrarconsulta(CnViewPeople)
	  CALL ofv.cerrarconn(Cnpeople)
      exit do




   END IF

   RSAX.CLOSE
   SET RSAX = NOTHING
   connaX.CLOSE
   SET connaX = NOTHING
   rsdom.MOVENEXT
   LOOP
   CALL ofv.cerrarconsulta(rsdom)
   CALL ofv.cerrarconn(cnDOM)
END FUNCTION







FUNCTION getUsrInTextAD(texto)

   xgusep = ""
   session("SGU_UIGuserid") = ""
   session("SGU_UIGdescri") = ""
   SET cnDOM = ofv.conectar(ofv.strconn0)
   Strsqlq = "Select * From dominios WHERE RAIZLDAP IS NOT NULL and DESCRIPCION <> 'DEFAULT'"
   SET rsdom = ofv.crearconsultaEx(StrSqlq, cnDOM, 1, parametros)
    TEXTO=REPLACE(TEXTO,"*","")
       IF TEXTO="*" or TEXTO=""   THEN
        TEXTO="NO"
       elseif len(texto)=1 then
       texto= texto & "*"
       else
       texto="*"  & texto & "*"
     end if
   do while (not rsdom.eof )
    strBase =rsdom("raizldap")
    usernx =  rsdom("usrldap")
    passx = cstr(ofv.getCript("D",rsdom("pwdldap")&""))'rsdom("pwdldap")

   SET connaX = server.CREATEOBJECT("ADODB.Connection")


   connaX.provider = "ADSDSOObject"
   connaX.Properties("User ID") = usernx
   connaX.Properties("Password") = passx
   connaX.Properties("Encrypt Password") = TRUE

   connaX.OPEN

   strAttributes = "sAMAccountName,cn,title,mail,name"

 '  strFilter = "(&(objectCategory=person)(objectClass=user)" & "(Name=" & texto &"))"
   strFilter="(&(objectCategory=person)(objectClass=user)(name=" & texto & ")(!userAccountControl:1.2.840.113556.1.4.803:=2))"

   strQuery = strBase & ";" & strFilter & ";" & strAttributes & ";subtree"
  ' response.write strQuery

   SET RSAX = connaX.EXECUTE(strQuery)



   DO UNTIL RSAX.EOF
      session("SGU_UIGuserid") = session("SGU_UIGuserid") & xgusep & RSAX("sAMAccountName")
      session("SGU_UIGdescri") = session("SGU_UIGdescri") & xgusep & RSAX("Name")
      xgusep = CHR(9)
      RSAX.movenext
   LOOP

   RSAX.CLOSE
   SET RSAX = NOTHING
   connaX.CLOSE
   SET connaX = NOTHING
   rsdom.movenext
 loop

END FUNCTION


FUNCTION getUsrInTextADBANELCO(texto)

   xgusep = ""
   session("SGU_UIGuserid") = ""
   session("SGU_UIGdescri") = ""
   SET cnDOM = ofv.conectar(ofv.strconn0)
   Strsqlq = "Select * From dominios WHERE RAIZLDAP IS NOT NULL and DESCRIPCION <> 'DEFAULT'"
   SET rsdom = ofv.crearconsultaEx(StrSqlq, cnDOM, 1, parametros)

   do while (not rsdom.eof )
    strBase =rsdom("raizldap")
    usernx =  rsdom("usrldap")
    passx = cstr(ofv.getCript("D",rsdom("pwdldap")&""))'rsdom("pwdldap")

   SET connaX = server.CREATEOBJECT("ADODB.Connection")


   connaX.provider = "ADSDSOObject"
   connaX.Properties("User ID") = usernx
   connaX.Properties("Password") = passx
   connaX.Properties("Encrypt Password") = TRUE

   connaX.OPEN

   strAttributes = "sAMAccountName,cn,title,mail,name"

 '  strFilter = "(&(objectCategory=person)(objectClass=user)" & "(Name=" & texto &"))"
   strFilter="(&(objectCategory=person)(objectClass=user)(name=" & texto & ")(!userAccountControl:1.2.840.113556.1.4.803:=2))"

   strQuery = strBase & ";" & strFilter & ";" & strAttributes & ";subtree"
  ' response.write strQuery

   SET RSAX = connaX.EXECUTE(strQuery)



   DO UNTIL RSAX.EOF
      session("SGU_UIGuserid") = session("SGU_UIGuserid") & xgusep & RSAX("sAMAccountName")
      session("SGU_UIGdescri") = session("SGU_UIGdescri") & xgusep & RSAX("Name")
      xgusep = CHR(9)
      RSAX.movenext
   LOOP

   RSAX.CLOSE
   SET RSAX = NOTHING
   connaX.CLOSE
   SET connaX = NOTHING
   rsdom.movenext
 loop

END FUNCTION


fUNCTION getUsrInTextAD_FINANSUR(texto)

   xgusep = ""
   session("SGU_UIGuserid") = ""
   session("SGU_UIGdescri") = ""
   SET cnDOM = ofv.conectar(ofv.strconn0)
   Strsqlq = "Select * From dominios WHERE RAIZLDAP IS NOT NULL and DESCRIPCION <> 'DEFAULT'"
   SET rsdom = ofv.crearconsultaEx(StrSqlq, cnDOM, 1, parametros)
    TEXTO=REPLACE(TEXTO,"*","")
     IF TEXTO="*" or TEXTO=""   THEN
      TEXTO="NO"
     else
     if trim(len(texto&""))=1 then
      texto=texto & "*"
     else
      texto="*"  & texto & "*"
     end if
     end if

   do while (not rsdom.eof )
    strBase =rsdom("raizldap")
    usernx =  rsdom("usrldap")
    passx = cstr(ofv.getCript("D",rsdom("pwdldap")&""))'rsdom("pwdldap")

    SET connaX = server.CREATEOBJECT("ADODB.Connection")


   connaX.provider = "ADSDSOObject"
   connaX.Properties("User ID") = usernx
   connaX.Properties("Password") = passx
   connaX.Properties("Encrypt Password") = TRUE

   connaX.OPEN

   strAttributes = "sAMAccountName,cn,title,mail,name,displayName"

 '  strFilter = "(&(objectCategory=person)(objectClass=user)" & "(Name=" & texto &"))"
   strFilter="(&(objectCategory=person)(objectClass=user)(displayName=" & texto & ")(!userAccountControl:1.2.840.113556.1.4.803:=2))"

   strQuery = strBase & ";" & strFilter & ";" & strAttributes & ";subtree"
  ' response.write strQuery

   SET objCommand = server.CREATEOBJECT("ADODB.Command")

   SET objCommand.ActiveConnection = connaX

     objCommand.Properties("Size Limit") =  PageSizead
     objCommand.Properties("Searchscope") = 2

   objCommand.CommandText=strQuery
   ' response.write  strQuery
   SET RSAX = objCommand.EXECUTE '(strQuery)

   totu=0

   DO while not  RSAX.EOF and totu < PageSizead
      session("SGU_UIGuserid") = session("SGU_UIGuserid") & xgusep & RSAX("sAMAccountName")
      session("SGU_UIGdescri") = session("SGU_UIGdescri") & xgusep & RSAX("displayName")
      xgusep = CHR(9)
      RSAX.movenext
      totu=totu+1
   LOOP

   RSAX.CLOSE
   SET RSAX = NOTHING
   set objCommand=NOTHING
   connaX.CLOSE
   SET connaX = NOTHING

   rsdom.movenext
 loop
CALL ofv.cerrarconsulta(rsdom)
CALL ofv.cerrarconn(cnDOM)
END FUNCTION



FUNCTION getUsrInTextAD_BPN(texto)

   xgusep = ""
   session("SGU_UIGuserid") = ""
   session("SGU_UIGdescri") = ""
   SET cnDOM = ofv.conectar(ofv.strconn0)

   Strsqlq = "Select * From dominios WHERE RAIZLDAP IS NOT NULL and DESCRIPCION <> 'DEFAULT'"
   SET rsdom = ofv.crearconsultaEx(StrSqlq, cnDOM, 1, parametros)

    TEXTO=REPLACE(TEXTO,"*","")
     IF TEXTO="*" or TEXTO=""   THEN
      TEXTO="NO"
     elseif len(texto)=1 then
     texto= texto & "*"
     else
     texto="*"  & texto & "*"
     end if

   do while (not rsdom.eof )
    strBase =rsdom("raizldap")
    usernx =  rsdom("usrldap")
    passx = cstr(ofv.getCript("D",rsdom("pwdldap")&""))'rsdom("pwdldap")

    SET connaX = server.CREATEOBJECT("ADODB.Connection")


   connaX.provider = "ADSDSOObject"
   connaX.Properties("User ID") = usernx
   connaX.Properties("Password") = passx
   connaX.Properties("Encrypt Password") = TRUE

   connaX.OPEN

   SET objCommand = server.CREATEOBJECT("ADODB.Command")

   SET objCommand.ActiveConnection = connaX

   objCommand.Properties("Size Limit") = 500
   objCommand.Properties("Searchscope") = 2
  ' objCommand.Properties("Sort On") = "NAME"

   SET objCommand.ActiveConnection = connaX
   strAttributes = "sAMAccountName,cn,title,mail,name"

 '  strFilter = "(&(objectCategory=person)(objectClass=user)" & "(Name=" & texto &"))"
   strFilter="(&(objectCategory=person)(objectClass=user)(name=" & texto & ")(!userAccountControl:1.2.840.113556.1.4.803:=2))"

   strQuery = strBase & ";" & strFilter & ";" & strAttributes & ";subtree"
   'response.write strQuery
   objCommand.CommandText=strQuery
   SET RSAX = objCommand.EXECUTE'(strQuery)


   totu=0
   DO while not  RSAX.EOF and totu < PageSizead
      session("SGU_UIGuserid") = session("SGU_UIGuserid") & xgusep & RSAX("sAMAccountName")
      session("SGU_UIGdescri") = session("SGU_UIGdescri") & xgusep & RSAX("Name")
      xgusep = CHR(9)
      RSAX.movenext
      totu=totu+1
   LOOP

   RSAX.CLOSE
   SET RSAX = NOTHING
  ' objCommand.close
   SET objCommand = NOTHING
   connaX.CLOSE
   SET connaX = NOTHING
   rsdom.movenext
 loop

END FUNCTION


FUNCTION getUsrInTextAD_GP(texto)

   xgusep = ""
   session("SGU_UIGuserid") = ""
   session("SGU_UIGdescri") = ""
   SET cnDOM = ofv.conectar(ofv.strconn0)

  ' Strsqlq = "Select * From dominios WHERE RAIZLDAP IS NOT NULL and DESCRIPCION <> 'DEFAULT'"

  if Session("EMPMUN")  = "S" then
       Strsqlq = "Select * From dominios WHERE cliente=0" & session("SETIDEMPRE")
       else
       Strsqlq = "Select * From dominios WHERE RAIZLDAP IS NOT NULL and DESCRIPCION <> 'DEFAULT'"
  end if
'       response.write Strsqlq
   SET rsdom = ofv.crearconsultaEx(StrSqlq, cnDOM, 1, parametros)

    TEXTO=REPLACE(TEXTO,"*","")
     IF TEXTO="*" or TEXTO=""   THEN
      TEXTO="NO"
     elseif len(texto)=1 then
     texto= texto & "*"
     else
     texto="*"  & texto & "*"
     end if

   do while (not rsdom.eof )
    strBase =rsdom("raizldap")
    usernx =  rsdom("usrldap")
    passx = cstr(ofv.getCript("D",rsdom("pwdldap")&""))'rsdom("pwdldap")

    SET connaX = server.CREATEOBJECT("ADODB.Connection")


   connaX.provider = "ADSDSOObject"
   connaX.Properties("User ID") = usernx
   connaX.Properties("Password") = passx
   connaX.Properties("Encrypt Password") = TRUE

   connaX.OPEN

   SET objCommand = server.CREATEOBJECT("ADODB.Command")

   SET objCommand.ActiveConnection = connaX

   objCommand.Properties("Size Limit") = 500
   objCommand.Properties("Searchscope") = 2
  ' objCommand.Properties("Sort On") = "NAME"

   SET objCommand.ActiveConnection = connaX
   strAttributes = "sAMAccountName,cn,title,mail,name"

 '  strFilter = "(&(objectCategory=person)(objectClass=user)" & "(Name=" & texto &"))"
   strFilter="(&(objectCategory=person)(objectClass=user)(name=" & texto & ")(!userAccountControl:1.2.840.113556.1.4.803:=2))"

   strQuery = strBase & ";" & strFilter & ";" & strAttributes & ";subtree"
   'response.write strQuery
   objCommand.CommandText=strQuery
   SET RSAX = objCommand.EXECUTE'(strQuery)


   totu=0
   DO while not  RSAX.EOF and totu < PageSizead
      session("SGU_UIGuserid") = session("SGU_UIGuserid") & xgusep & RSAX("sAMAccountName")
      session("SGU_UIGdescri") = session("SGU_UIGdescri") & xgusep & RSAX("Name")
      xgusep = CHR(9)
      RSAX.movenext
      totu=totu+1
   LOOP

   RSAX.CLOSE
   SET RSAX = NOTHING
  ' objCommand.close
   SET objCommand = NOTHING
   connaX.CLOSE
   SET connaX = NOTHING
   rsdom.movenext
 loop

END FUNCTION



FUNCTION getGruposMail_BPN(texto)

   xgusep = ""
   session("SGU_UIGuserid") = ""
   session("SGU_UIGdescri") = ""
   SET cnDOM = ofv.conectar(ofv.strconn0)
   Strsqlq = "Select * From dominios WHERE RAIZLDAP IS NOT NULL and DESCRIPCION <> 'DEFAULT'"
   SET rsdom = ofv.crearconsultaEx(StrSqlq, cnDOM, 1, parametros)
    TEXTO=trim(REPLACE(TEXTO,"*",""))
     IF TEXTO="*" or TEXTO=""   THEN
      TEXTO="*"
     else
     texto="*"  & texto & "*"
     end if

   do while (not rsdom.eof )
    strBase =rsdom("raizldap")
    usernx =  rsdom("usrldap")
    passx = cstr(ofv.getCript("D",rsdom("pwdldap")&""))'rsdom("pwdldap")

    SET connaX = server.CREATEOBJECT("ADODB.Connection")


   connaX.provider = "ADSDSOObject"
   connaX.Properties("User ID") = usernx
   connaX.Properties("Password") = passx
   connaX.Properties("Encrypt Password") = TRUE


     connaX.OPEN
    SET objCommand = server.CREATEOBJECT("ADODB.Command")

     SET objCommand.ActiveConnection = connaX

     objCommand.Properties("Size Limit") = PageSizead
     objCommand.Properties("Searchscope") = 2
     objCommand.Properties("Sort On") = "name"
     strAttributes = "name,mail"
  ' strFilter = "(&(objectCategory=group)((&(groupType:1.2.840.113556.1.4.803:=8)(!(groupType:1.2.840.113556.1.4.803:=-2147483648)))))"


   'strFilter = "(&((&(objectCategory=group)(&(groupType:1.2.840.113556.1.4.803:=8)(!(groupType:1.2.840.113556.1.4.803:=-2147483648))))) (name=" & texto &  ") )"

   ' activar siguiente linea en cor porque no hay exchange
   strFilter =  	"(&(objectCategory=group)(name=" &texto&"))"

   strQuery = strBase & ";" & strFilter & ";" & strAttributes & ";subtree"
   objCommand.CommandText=strQuery
 ' response.write  strQuery
   SET RSAX = objCommand.EXECUTE '(strQuery)

  ' response.write RSAX.EOF


   DO while not  RSAX.EOF

     if ValidaMail(rsax("mail")& "") = 2 then
        session("SGU_UIGuserid") = session("SGU_UIGuserid") & xgusep & RSAX("name")
        session("SGU_UIGdescri") = session("SGU_UIGdescri") & xgusep & RSAX("Name")
        xgusep = CHR(9)
     end if
      RSAX.movenext
   LOOP

   RSAX.CLOSE
   SET RSAX = NOTHING
   connaX.CLOSE
   SET connaX = NOTHING
   rsdom.movenext
 loop

END FUNCTION

fUNCTION getUsrInTextAD_BCR(texto)

   xgusep = ""
   session("SGU_UIGuserid") = ""
   session("SGU_UIGdescri") = ""
   SET cnDOM = ofv.conectar(ofv.strconn0)
   Strsqlq = "Select * From dominios WHERE RAIZLDAP IS NOT NULL and DESCRIPCION <> 'DEFAULT'"
   SET rsdom = ofv.crearconsultaEx(StrSqlq, cnDOM, 1, parametros)
    TEXTO=REPLACE(TEXTO,"*","")
     IF TEXTO="*" or TEXTO=""   THEN
      TEXTO="NO"
     else
     if trim(len(texto&""))=1 then
      texto=texto & "*"
     else
      texto="*"  & texto & "*"
     end if
     end if

   do while (not rsdom.eof )
    strBase =rsdom("raizldap")
    usernx =  rsdom("usrldap")
    passx = cstr(ofv.getCript("D",rsdom("pwdldap")&""))'rsdom("pwdldap")

    SET connaX = server.CREATEOBJECT("ADODB.Connection")


   connaX.provider = "ADSDSOObject"
   connaX.Properties("User ID") = usernx
   connaX.Properties("Password") = passx
   connaX.Properties("Encrypt Password") = TRUE

   connaX.OPEN

   strAttributes = "sAMAccountName,cn,title,mail,name,displayName"

 '  strFilter = "(&(objectCategory=person)(objectClass=user)" & "(Name=" & texto &"))"
   strFilter="(&(objectCategory=person)(objectClass=user)(displayName=" & texto & ")(!userAccountControl:1.2.840.113556.1.4.803:=2))"

   strQuery = strBase & ";" & strFilter & ";" & strAttributes & ";subtree"
  ' response.write strQuery

   SET objCommand = server.CREATEOBJECT("ADODB.Command")

   SET objCommand.ActiveConnection = connaX

     objCommand.Properties("Size Limit") =  PageSizead
     objCommand.Properties("Searchscope") = 2

   objCommand.CommandText=strQuery
   ' response.write  strQuery
   SET RSAX = objCommand.EXECUTE '(strQuery)

   totu=0

   DO while not  RSAX.EOF and totu < PageSizead
      session("SGU_UIGuserid") = session("SGU_UIGuserid") & xgusep & RSAX("sAMAccountName")
      session("SGU_UIGdescri") = session("SGU_UIGdescri") & xgusep & RSAX("displayName")
      xgusep = CHR(9)
      RSAX.movenext
      totu=totu+1
   LOOP

   RSAX.CLOSE
   SET RSAX = NOTHING
   set objCommand=NOTHING
   connaX.CLOSE
   SET connaX = NOTHING

   rsdom.movenext
 loop
CALL ofv.cerrarconsulta(rsdom)
CALL ofv.cerrarconn(cnDOM)
END FUNCTION

FUNCTION getMailDelgrupo_BPN(texto)



   SET cnDOM = ofv.conectar(ofv.strconn0)
   Strsqlq = "Select * From dominios WHERE RAIZLDAP IS NOT NULL and DESCRIPCION <> 'DEFAULT'"
   SET rsdom = ofv.crearconsultaEx(StrSqlq, cnDOM, 1, parametros)
    TEXTO=trim(REPLACE(TEXTO,"*",""))
   getMailDelgrupo_BPN=""

   do while (not rsdom.eof )
    strBase =rsdom("raizldap")
    usernx =  rsdom("usrldap")
    passx = cstr(ofv.getCript("D",rsdom("pwdldap")&""))'rsdom("pwdldap")

    SET connaX = server.CREATEOBJECT("ADODB.Connection")


   connaX.provider = "ADSDSOObject"
   connaX.Properties("User ID") = usernx
   connaX.Properties("Password") = passx
   connaX.Properties("Encrypt Password") = TRUE


   connaX.OPEN
    SET objCommand = server.CREATEOBJECT("ADODB.Command")

     SET objCommand.ActiveConnection = connaX

     objCommand.Properties("Size Limit") = 500
     objCommand.Properties("Searchscope") = 2

    strAttributes = "name,mail"

   strFilter =  	"(&(objectCategory=group)(name=" &texto&"))"
   strQuery = strBase & ";" & strFilter & ";" & strAttributes & ";subtree"
   objCommand.CommandText=strQuery
  'response.write  strQuery
   SET RSAX = objCommand.EXECUTE

  ' response.write RSAX.EOF


   if not  RSAX.EOF then

    ' if ValidaMail(rsax("mail")& "") = 2 then
        getMailDelgrupo_BPN=rsax("mail")&""
        exit do
   '  end if
   end if
   RSAX.CLOSE
   SET RSAX = NOTHING

   rsdom.movenext
 loop


  connaX.CLOSE
  SET connaX = NOTHING

END FUNCTION





FUNCTION getinfoUserAD(legajo)

 SET cnDOM = ofv.conectar(ofv.strconn0)

      Strsqlq = "Select * From dominios WHERE RAIZLDAP IS NOT NULL and DESCRIPCION <> 'DEFAULT' order by id"

      SET rsdom = ofv.crearconsultaEx(StrSqlq, cnDOM, 1, parametros)
    ' session("SGU_nombre")=""
     do while (not rsdom.eof) ' OR (session("SGU_nombre") <> "NO")
     sldap =rsdom("raizldap")


      usernx =  rsdom("usrldap")
      passx = cstr(ofv.getCript("D",rsdom("pwdldap")&""))'rsdom("pwdldap") 'ofv.GetCript("D", CSTR(rsdom("pwdldap")))

      SET connaX = server.CREATEOBJECT("ADODB.Connection")


       connaX.provider = "ADSDSOObject"
       connaX.Properties("User ID") = usernx
       connaX.Properties("Password") = passx
       connaX.Properties("Encrypt Password") = TRUE
       ' response.write userN & "----" & pass & "---" & strBase
        connaX.OPEN

         strAttributes = "sAMAccountName,cn,title,mail,memberOf,name"

         strFilter = "(&(objectCategory=person)(objectClass=user)" & "(samAccountName=" & legajo &"))"


         strQuery = sldap & ";" & strFilter & ";" & strAttributes & ";subtree"
   'response.write strQuery
   SET RSAX = connaX.EXECUTE(strQuery)



   IF RSAX.EOF THEN
      session("SGU_userid") = Legajo
      session("SGU_nombre") = "NO"
      session("SGU_mail") = "NO"
      session("SGU_descrisector") = "NO"
      session("SGU_idcostcenter") = "0"
      session("SGU_descripuesto") = "NO"
      session("SGU_codigopuesto") = "NO"

   ELSE

      session("SGU_userid") = Legajo
      session("SGU_nombre") = RSAX("name")
      session("SGU_mail") =  RSAX("mail")
      session("SGU_descrisector") = "NO"
      session("SGU_idcostcenter") = "0"
      session("SGU_descripuesto") = RSAX("title")
      session("SGU_codigopuesto") = "NO"
      exit do
   END IF

   RSAX.CLOSE
   SET RSAX = NOTHING
   connaX.CLOSE
   SET connaX = NOTHING
   rsdom.MOVENEXT
   LOOP
  CALL ofv.cerrarconsulta(rsdom)
  CALL ofv.cerrarconn(cnDOM)


END FUNCTION


FUNCTION getinfoUserADlacaja(legajo)


      SET cnDOM = ofv.conectar(ofv.strconn0)

      Strsqlq = "Select * From dominios WHERE RAIZLDAP IS NOT NULL and DESCRIPCION <> 'DEFAULT'"

      SET rsdom = ofv.crearconsultaEx(StrSqlq, cnDOM, 1, parametros)

     do while (not rsdom.eof)  OR (session("SGU_nombre") <> "NO")
     sldap =rsdom("raizldap")


      usernx =  rsdom("usrldap")
      passx = cstr(ofv.getCript("D",rsdom("pwdldap")&""))'rsdom("pwdldap") 'ofv.GetCript("D", CSTR(rsdom("pwdldap")))

      SET connaX = server.CREATEOBJECT("ADODB.Connection")


       connaX.provider = "ADSDSOObject"
       connaX.Properties("User ID") = usernx
       connaX.Properties("Password") = passx
       connaX.Properties("Encrypt Password") = TRUE
       ' response.write userN & "----" & pass & "---" & strBase
        connaX.OPEN

         strAttributes = "sAMAccountName,cn,title,mail,memberOf,name"

         strFilter = "(&(objectCategory=person)(objectClass=user)" & "(samAccountName=" & legajo &"))"


         strQuery = sldap & ";" & strFilter & ";" & strAttributes & ";subtree"
   'response.write strQuery
   SET RSAX = connaX.EXECUTE(strQuery)



   IF RSAX.EOF THEN


      session("SGU_userid") = Legajo
      session("SGU_nombre") = "NO"
      session("SGU_mail") = "NO"
      session("SGU_descrisector") = "NO"
      session("SGU_idcostcenter") = "0"
      session("SGU_descripuesto") = "NO"
      session("SGU_codigopuesto") = "NO"

   ELSE

      session("SGU_userid") = Legajo
      session("SGU_nombre") = RSAX("name")
      session("SGU_mail") = legajo & "@lacaja.com.ar" ' RSAX("mail")
      session("SGU_descrisector") = "NO"
      session("SGU_idcostcenter") = "0"
      session("SGU_descripuesto") = RSAX("title")
      session("SGU_codigopuesto") = "NO"
      exit do
   END IF

   RSAX.CLOSE
   SET RSAX = NOTHING
   connaX.CLOSE
   SET connaX = NOTHING
   rsdom.MOVENEXT
   LOOP
  CALL ofv.cerrarconsulta(rsdom)
  CALL ofv.cerrarconn(cnDOM)

END FUNCTION



FUNCTION getinfoUserADBanelco(legajo)


      SET cnDOM = ofv.conectar(ofv.strconn0)

      Strsqlq = "Select * From dominios WHERE RAIZLDAP IS NOT NULL and DESCRIPCION <> 'DEFAULT'"

      SET rsdom = ofv.crearconsultaEx(StrSqlq, cnDOM, 1, parametros)

     do while (not rsdom.eof)  OR (session("SGU_nombre") <> "NO")
     sldap =rsdom("raizldap")


      usernx =  rsdom("usrldap")
      passx = cstr(ofv.getCript("D",rsdom("pwdldap")&""))'rsdom("pwdldap") 'ofv.GetCript("D", CSTR(rsdom("pwdldap")))

      SET connaX = server.CREATEOBJECT("ADODB.Connection")


       connaX.provider = "ADSDSOObject"
       connaX.Properties("User ID") = usernx
       connaX.Properties("Password") = passx
       connaX.Properties("Encrypt Password") = TRUE
       ' response.write userN & "----" & pass & "---" & strBase
        connaX.OPEN

         strAttributes = "sAMAccountName,cn,title,mail,memberOf,name,department,description"

         strFilter = "(&(objectCategory=person)(objectClass=user)" & "(samAccountName=" & legajo &"))"


         strQuery = sldap & ";" & strFilter & ";" & strAttributes & ";subtree"
   'response.write strQuery
   SET RSAX = connaX.EXECUTE(strQuery)



   IF RSAX.EOF THEN


      session("SGU_userid") = Legajo
      session("SGU_nombre") = "NO"
      session("SGU_mail") = "NO"
      session("SGU_descrisector") = "NO"
      session("SGU_idcostcenter") = "0"
      session("SGU_descripuesto") = "NO"
      session("SGU_codigopuesto") = "NO"
      session("SGU_idnivel") ="9"

   ELSE

      session("SGU_userid") = Legajo
      session("SGU_nombre") = RSAX("name")
      session("SGU_mail") = RSAX("mail")
      session("SGU_descrisector") = "NO"
      session("SGU_idcostcenter") = getsectorExt(rsax("department")&"")
      session("SGU_descripuesto") = "NO"
      session("SGU_codigopuesto") = getperfilExtBanelco(rsax("title"))
      if isarray(rsax("description"))then
        xnivel=rsax("description")
        session("SGU_idnivel") =xnivel(0)
      else
        xnivel=rsax("description")
        session("SGU_idnivel") =xnivel
      end if
      session("SGU_idnivel")=getnivelExtBanelco(ucase(session("SGU_idnivel")))
      exit do
   END IF

   RSAX.CLOSE
   SET RSAX = NOTHING
   connaX.CLOSE
   SET connaX = NOTHING
   rsdom.MOVENEXT
   LOOP
  CALL ofv.cerrarconsulta(rsdom)
  CALL ofv.cerrarconn(cnDOM)

END FUNCTION


FUNCTION getinfoUserAD_FINANSUR(legajo)


      SET cnDOM = ofv.conectar(ofv.strconn0)

      Strsqlq = "Select * From dominios WHERE RAIZLDAP IS NOT NULL and DESCRIPCION <> 'DEFAULT'"

      SET rsdom = ofv.crearconsultaEx(StrSqlq, cnDOM, 1, parametros)

     do while (not rsdom.eof)  OR (session("SGU_nombre") <> "NO")
     sldap =rsdom("raizldap")


      usernx =  rsdom("usrldap")
      passx = cstr(ofv.getCript("D",rsdom("pwdldap")&""))'rsdom("pwdldap") 'ofv.GetCript("D", CSTR(rsdom("pwdldap")))

      SET connaX = server.CREATEOBJECT("ADODB.Connection")


       connaX.provider = "ADSDSOObject"
       connaX.Properties("User ID") = usernx
       connaX.Properties("Password") = passx
       connaX.Properties("Encrypt Password") = TRUE
       ' response.write userN & "----" & pass & "---" & strBase
        connaX.OPEN

         strAttributes = "sAMAccountName,cn,title,mail,memberOf,name,department,description"

         strFilter = "(&(objectCategory=person)(objectClass=user)" & "(samAccountName=" & legajo &"))"


         strQuery = sldap & ";" & strFilter & ";" & strAttributes & ";subtree"
   'response.write strQuery
   SET RSAX = connaX.EXECUTE(strQuery)



   IF RSAX.EOF THEN


      session("SGU_userid") = Legajo
      session("SGU_nombre") = "NO"
      session("SGU_mail") = "NO"
      session("SGU_descrisector") = "NO"
      session("SGU_idcostcenter") = "0"
      session("SGU_descripuesto") = "NO"
      session("SGU_codigopuesto") = "NO"
      session("SGU_idnivel") ="9"

   ELSE

      session("SGU_userid") = Legajo
      session("SGU_nombre") = RSAX("name")
      session("SGU_mail") = RSAX("mail")
      session("SGU_descrisector") = "NO"
      session("SGU_idcostcenter") = "0" 'getsectorExt(rsax("department")&"")
      session("SGU_descripuesto") = "NO"
      session("SGU_codigopuesto") = "NO " 'getperfilExtBanelco(rsax("title"))
      session("SGU_idnivel")="9" 'getnivelExtBanelco(ucase(session("SGU_idnivel")))
      exit do
   END IF

   RSAX.CLOSE
   SET RSAX = NOTHING
   connaX.CLOSE
   SET connaX = NOTHING
   rsdom.MOVENEXT
   LOOP
  CALL ofv.cerrarconsulta(rsdom)
  CALL ofv.cerrarconn(cnDOM)

END FUNCTION



FUNCTION getinfoUserAD_BPN(legajo)

 IF LEN(TRIM(LEGAJO & "")) = 0 THEN

      session("SGU_userid") = "--" & Legajo & "-"
  	      session("SGU_nombre") = "NO"
  	      session("SGU_mail") = "NO"
  	      session("SGU_descrisector") = "NO"
  	      session("SGU_idcostcenter") = "0"
  	      session("SGU_descripuesto") = "NO"
  	      session("SGU_codigopuesto") = "NO"
        session("SGU_idnivel") ="9"
  ELSE
      SET cnDOM = ofv.conectar(ofv.strconn0)

      Strsqlq = "Select * From dominios WHERE RAIZLDAP IS NOT NULL and DESCRIPCION <> 'DEFAULT'"

      SET rsdom = ofv.crearconsultaEx(StrSqlq, cnDOM, 1, parametros)

     do while (not rsdom.eof)  OR (session("SGU_nombre") <> "NO")
     sldap =rsdom("raizldap")


      usernx =  rsdom("usrldap")
      passx = cstr(ofv.getCript("D",rsdom("pwdldap")&""))'rsdom("pwdldap") 'ofv.GetCript("D", CSTR(rsdom("pwdldap")))

      SET connaX = server.CREATEOBJECT("ADODB.Connection")


       connaX.provider = "ADSDSOObject"
       connaX.Properties("User ID") = usernx
       connaX.Properties("Password") = passx
       connaX.Properties("Encrypt Password") = TRUE
       ' response.write userN & "----" & pass & "---" & strBase
        connaX.OPEN

         strAttributes = "sAMAccountName,cn,title,mail,memberOf,name,department,description"

         strFilter = "(&(objectCategory=person)(objectClass=user)" & "(samAccountName=" & legajo &"))"


         strQuery = sldap & ";" & strFilter & ";" & strAttributes & ";subtree"
   'response.write strQuery
   SET RSAX = connaX.EXECUTE(strQuery)



   IF RSAX.EOF THEN


      session("SGU_userid") = Legajo
      session("SGU_nombre") = "NO"
      session("SGU_mail") = "NO"
      session("SGU_descrisector") = "NO"
      session("SGU_idcostcenter") = "0"
      session("SGU_descripuesto") = "NO"
      session("SGU_codigopuesto") = "NO"
      session("SGU_idnivel") ="9"

   ELSE

      session("SGU_userid") = Legajo
      session("SGU_nombre") = RSAX("name")
      session("SGU_mail") = RSAX("mail")
      session("SGU_descrisector") = "NO"
      session("SGU_idcostcenter") = "0" 'getsectorExt(rsax("department")&"")
      session("SGU_descripuesto") = "NO"
      session("SGU_codigopuesto") = "NO " 'getperfilExtBanelco(rsax("title"))
      session("SGU_idnivel")="9" 'getnivelExtBanelco(ucase(session("SGU_idnivel")))
      exit do
   END IF

   RSAX.CLOSE
   SET RSAX = NOTHING
   connaX.CLOSE
   SET connaX = NOTHING
   rsdom.MOVENEXT
   LOOP
  CALL ofv.cerrarconsulta(rsdom)
  CALL ofv.cerrarconn(cnDOM)
end if
END FUNCTION


FUNCTION getinfoUserAD_GP(legajo)

 IF LEN(TRIM(LEGAJO & "")) = 0 THEN

      session("SGU_userid") = "--" & Legajo & "-"
  	      session("SGU_nombre") = "NO"
  	      session("SGU_mail") = "NO"
  	      session("SGU_descrisector") = "NO"
  	      session("SGU_idcostcenter") = "0"
  	      session("SGU_descripuesto") = "NO"
  	      session("SGU_codigopuesto") = "NO"
          session("SGU_idnivel") ="9"
  ELSE
      SET cnDOM = ofv.conectar(ofv.strconn0)

      Strsqlq = "Select * From dominios WHERE RAIZLDAP IS NOT NULL and DESCRIPCION <> 'DEFAULT'"

      SET rsdom = ofv.crearconsultaEx(StrSqlq, cnDOM, 1, parametros)

     do while (not rsdom.eof)  OR (session("SGU_nombre") <> "NO")
     sldap =rsdom("raizldap")


      usernx =  rsdom("usrldap")
      passx = cstr(ofv.getCript("D",rsdom("pwdldap")&""))'rsdom("pwdldap") 'ofv.GetCript("D", CSTR(rsdom("pwdldap")))
     RESPONSE.WRITE passx
      SET connaX = server.CREATEOBJECT("ADODB.Connection")


       connaX.provider = "ADSDSOObject"
       connaX.Properties("User ID") = usernx
       connaX.Properties("Password") = passx
       connaX.Properties("Encrypt Password") = TRUE
       ' response.write userN & "----" & pass & "---" & strBase
        connaX.OPEN

         strAttributes = "sAMAccountName,cn,title,mail,memberOf,name,department,description"

         strFilter = "(&(objectCategory=person)(objectClass=user)" & "(samAccountName=" & legajo &"))"


         strQuery = sldap & ";" & strFilter & ";" & strAttributes & ";subtree"

   SET RSAX = connaX.EXECUTE(strQuery)



   IF RSAX.EOF THEN


      session("SGU_userid") = Legajo
      session("SGU_nombre") = "NO"
      session("SGU_mail") = "NO"
      session("SGU_descrisector") = "NO"
      session("SGU_idcostcenter") = "0"
      session("SGU_descripuesto") = "NO"
      session("SGU_codigopuesto") = "NO"
      session("SGU_idnivel") ="9"
      session("SGU_site")=1

   ELSE

      session("SGU_userid") = Legajo
      session("SGU_nombre") = RSAX("name")
      session("SGU_mail") = RSAX("mail")
      session("SGU_descrisector") = "NO"
      session("SGU_idcostcenter") = "0" 'getsectorExt(rsax("department")&"")
      session("SGU_descripuesto") = "NO"
      session("SGU_codigopuesto") = "NO " 'getperfilExtBanelco(rsax("title"))
      session("SGU_idnivel")="9" 'getnivelExtBanelco(ucase(session("SGU_idnivel")))
      session("SGU_site")=session("usr_dominio")
    '  RESPONSE.WRITE session("SGU_site")
      exit do
   END IF

   RSAX.CLOSE
   SET RSAX = NOTHING
   connaX.CLOSE
   SET connaX = NOTHING
   rsdom.MOVENEXT
   LOOP
   'response.write strQuery & session("SGU_site")
  CALL ofv.cerrarconsulta(rsdom)
  CALL ofv.cerrarconn(cnDOM)
end if
END FUNCTION

FUNCTION getinfoUserAD_BCR(legajo)

  IF LEN(TRIM(LEGAJO & "")) = 0 THEN

      session("SGU_userid") = "--" & Legajo & "-"
  	      session("SGU_nombre") = "NO"
  	      session("SGU_mail") = "NO"
  	      session("SGU_descrisector") = "NO"
  	      session("SGU_idcostcenter") = "0"
  	      session("SGU_descripuesto") = "NO"
  	      session("SGU_codigopuesto") = "NO"
        session("SGU_idnivel") ="9"
  ELSE
      SET cnDOM = ofv.conectar(ofv.strconn0)

      Strsqlq = "Select * From dominios WHERE RAIZLDAP IS NOT NULL and DESCRIPCION <> 'DEFAULT'"

      SET rsdom = ofv.crearconsultaEx(StrSqlq, cnDOM, 1, parametros)

     do while (not rsdom.eof)  OR (session("SGU_nombre") <> "NO")
     sldap =rsdom("raizldap")


      usernx =  rsdom("usrldap")
      passx = cstr(ofv.getCript("D",rsdom("pwdldap")&""))'rsdom("pwdldap") 'ofv.GetCript("D", CSTR(rsdom("pwdldap")))

      SET connaX = server.CREATEOBJECT("ADODB.Connection")


       connaX.provider = "ADSDSOObject"
       connaX.Properties("User ID") = usernx
       connaX.Properties("Password") = passx
       connaX.Properties("Encrypt Password") = TRUE
       ' response.write userN & "----" & pass & "---" & strBase
        connaX.OPEN

         strAttributes = "sAMAccountName,cn,title,mail,memberOf,name,department,description,displayName"

         strFilter = "(&(objectCategory=person)(objectClass=user)" & "(samAccountName=" & legajo &"))"


         strQuery = sldap & ";" & strFilter & ";" & strAttributes & ";subtree"
   'response.write strQuery
   SET RSAX = connaX.EXECUTE(strQuery)



   IF RSAX.EOF THEN


      session("SGU_userid") = Legajo
      session("SGU_nombre") = "NO"
      session("SGU_mail") = "NO"
      session("SGU_descrisector") = "NO"
      session("SGU_idcostcenter") = "0"
      session("SGU_descripuesto") = "NO"
      session("SGU_codigopuesto") = "NO"
      session("SGU_idnivel") ="9"

   ELSE

      session("SGU_userid") = Legajo
      session("SGU_nombre") = RSAX("displayName")
      session("SGU_mail") = RSAX("mail")
      session("SGU_descrisector") = "NO"
      session("SGU_idcostcenter") = getsectorExtbcr(rsax("department")&"")
      session("SGU_descripuesto") = "NO"
      session("SGU_codigopuesto") ="NO" 'getperfilExtbcr(rsax("title"))
      session("SGU_idnivel")=getnivelExtBCR(rsax("title")&"")
      exit do
   END IF

   RSAX.CLOSE
   SET RSAX = NOTHING
   connaX.CLOSE
   SET connaX = NOTHING
   rsdom.MOVENEXT
   LOOP
  CALL ofv.cerrarconsulta(rsdom)
  CALL ofv.cerrarconn(cnDOM)
 end if


END FUNCTION






FUNCTION getinfoUserADBC(legajo)


      SET cnDOM = ofv.conectar(ofv.strconn0)

      Strsqlq = "Select * From dominios WHERE RAIZLDAP IS NOT NULL and DESCRIPCION <> 'DEFAULT'"

      SET rsdom = ofv.crearconsultaEx(StrSqlq, cnDOM, 1, parametros)

     do while (not rsdom.eof)  OR (session("SGU_nombre") <> "NO")
     sldap =rsdom("raizldap")


      usernx =  rsdom("usrldap")
      passx =cstr(ofv.getCript("D",rsdom("pwdldap")&""))' rsdom("pwdldap") 'ofv.GetCript("D", CSTR(rsdom("pwdldap")))

      SET connaX = server.CREATEOBJECT("ADODB.Connection")


       connaX.provider = "ADSDSOObject"
       connaX.Properties("User ID") = usernx
       connaX.Properties("Password") = passx
       connaX.Properties("Encrypt Password") = TRUE
       ' response.write userN & "----" & pass & "---" & strBase
        connaX.OPEN

         strAttributes = "sAMAccountName,cn,title,mail,memberOf,name"

         strFilter = "(&(objectCategory=person)(objectClass=user)" & "(samAccountName=" & legajo &"))"


         strQuery = sldap & ";" & strFilter & ";" & strAttributes & ";subtree"
   'response.write strQuery
   SET RSAX = connaX.EXECUTE(strQuery)



   IF RSAX.EOF THEN


      session("SGU_userid") = Legajo
      session("SGU_nombre") = "NO"
      session("SGU_mail") = "NO"
      session("SGU_descrisector") = "NO"
      session("SGU_idcostcenter") = "0"
      session("SGU_descripuesto") = "NO"
      session("SGU_codigopuesto") = "NO"

   ELSE

         set Cnpeople =ofv.conectar(ofv.strconn7)
	 legajoZ=mid(legajo,3,len(legajo))
	 strsqlCn = "select ALTER_EMPLID  ,FIRST_NAME,MIDDLE_NAME,LAST_NAME, DEPTID  ,JOBCODE   from PS_BC_EMPL_INGR_VW where ALTER_EMPLID='" & legajoZ &"'"

	 set CnViewPeople=ofv.crearconsultaEx(strsqlCn ,Cnpeople ,1,volver)

	 if not CnViewPeople.eof then

	  session("SGU_userid") = Legajo
	  session("SGU_nombre") =CnViewPeople("FIRST_NAME") & "&nbsp;" &  CnViewPeople("MIDDLE_NAME") & " " & CnViewPeople("LAST_NAME")
	  session("SGU_mail") = ChangeMail(RSAX("mail"))'RSAX("mail")
	  session("SGU_descrisector") = "NO"
	  SDEPTID = CnViewPeople("DEPTID") & ""
	  SDEPTID=TRIM(SDEPTID)
	  session("SGU_idcostcenter") = getsectorExt(SDEPTID) 'CnViewPeople("DEPTID")
	  session("SGU_descripuesto") = ""
	  SXJOB=CnViewPeople("JOBCODE") & ""
	  SXJOB=TRIM(SXJOB)
	  session("SGU_codigopuesto") = SXJOB
	  session("SGU_idnivel")      = getnivelExt(SXJOB)
	 else

	  session("SGU_userid") = Legajo
	  session("SGU_nombre") = "NO DISPONIBLE - " & ERR.NUMBER & "&nbsp;" & ERR.DESCRIPTION '& strsqlCn
	  session("SGU_mail") = "NO@disponible.com"
	  session("SGU_descrisector") = "NO disponible"
	  session("SGU_idcostcenter") = "0"
	  session("SGU_descripuesto") = "NO disponible"
	  session("SGU_codigopuesto") = "NO"
	  session("SGU_idnivel")      ="9"
	 end if

	 call ofv.cerrarconsulta(CnViewPeople)
	 CALL ofv.cerrarconn(Cnpeople)

      exit do
   END IF

   RSAX.CLOSE
   SET RSAX = NOTHING
   connaX.CLOSE
   SET connaX = NOTHING
   rsdom.MOVENEXT
   LOOP
  CALL ofv.cerrarconsulta(rsdom)
  CALL ofv.cerrarconn(cnDOM)

END FUNCTION


Function GetUserInsectBC(XGU_SECT)

session("SGU_UIGuserid")=""
session("SGU_UIGdescri")=""
IF ISNUMERIC(XGU_SECT) THEN
	   WXGU_SECT = XGU_SECT
	ELSE
	   WXGU_SECT = 1
	END IF

set xcn = ofv.conectar(ofv.strconn2)

xStrSql = "Select codinter From eco WHERE eco = 0" & WXGU_SECT
Set xrs = OFV.crearconsultaEx(xStrSql, xcn, 1, "")

if not xrs.eof then WXGU_SECT= xrs(0)

call ofv.cerrarconsulta(xrs)
call ofv.cerrarconn(xcn)

 set Cnpeople=ofv.conectar(ofv.strconn7)
 strsqlCn= "select ALTER_EMPLID ,FIRST_NAME,MIDDLE_NAME,LAST_NAME from PS_BC_EMPL_INGR_VW where DEPTID='" & WXGU_SECT &"'"
 set CnViewPeople=ofv.crearconsultaEx(strsqlCn ,Cnpeople ,1,volver)

  DO WHILE NOT CnViewPeople.EOF

      SALTER_EMPLID=CnViewPeople("ALTER_EMPLID") & ""
	  SALTER_EMPLID=TRIM(SALTER_EMPLID)
          legajo=getInfouserBCEx(SALTER_EMPLID)

	 session("SGU_UIGuserid") =  session("SGU_UIGuserid") & xgusep & legajo
	 session("SGU_UIGdescri") = session("SGU_UIGdescri") & xgusep & CnViewPeople("FIRST_NAME") & " " &  CnViewPeople("MIDDLE_NAME") & " " & CnViewPeople("LAST_NAME")
	 xgusep = CHR(9)
	 CnViewPeople.movenext
  LOOP

  call ofv.cerrarconsulta(CnViewPeople)
  CALL ofv.cerrarconn(Cnpeople)


end function

'------------------------------------------------------------------------------------------------------------

FUNCTION ValDominio(dominio)


   IF session("ldap") = "" THEN
   session("usr_dominio")=1
      SET cnDOM = ofv.conectar(ofv.strconn0)

      Strsql = "Select * From dominios WHERE DESCRIPCION='" & dominio & "'"
'RESPONSE.WRITE STRSQL
      SET rsdom = ofv.crearconsultaEx(StrSql, cnDOM, 1, parametros)
      IF NOT rsdom.EOF THEN
         session("usr_dominio")=rsdom("cliente")
         session("ldap") =rsdom("raizldap")
         session("usrldap") =  rsdom("usrldap")
         if isnumeric(rsdom("pwdldap")&"") then
         session("pwdldap") = ofv.GetCript("D", CSTR(rsdom("pwdldap")&"")) 'rsdom("pwdldap") '
         else
          session("pwdldap") = rsdom("pwdldap")
         end if
      END IF
      CALL ofv.cerrarconsulta(rsdom)
      CALL ofv.cerrarconn(cnDOM)

   END IF
 ' RESPONSE.WRITE  Strsql  & "---" & dominio 'session("ldap")

END FUNCTION

'--------------------------------------------------------------------------------------------------------------

FUNCTION FullNameGrupo(idgrupo,strbase1)
   DIM rsad
   DIM cnad


   SET cnad = server.CREATEOBJECT("ADODB.Connection")


   cnad.provider = "ADSDSOObject"
   cnad.Properties("User ID") = userN
   cnad.Properties("Password") = pass
   cnad.Properties("Encrypt Password") = TRUE

   cnAd.OPEN

   strAttributes = "name,distinguishedName"

   strFilter = "(&(objectCategory=group) " & "(Name=" & idgrupo &"))"


   strQuery = strBase1 & ";" & strFilter & ";" & strAttributes & ";subtree"
     'response.write strQuery
   SET rsad = cnAd.EXECUTE(strQuery)

   IF NOT rsad.EOF THEN
      FullNameGrupo = rsad("distinguishedName")
   ELSE
      FullNameGrupo = idgrupo
   END IF

   rsad.CLOSE
   SET rsad = NOTHING
   cnAd.CLOSE
   SET cnad = NOTHING

END FUNCTION


FUNCTION FullNameGrupoEx(idgrupo,strbase1,userN,pass)
   DIM rsad
   DIM cnad


   SET cnad = server.CREATEOBJECT("ADODB.Connection")


   cnad.provider = "ADSDSOObject"
   cnad.Properties("User ID") = userN
   cnad.Properties("Password") = ofv.GetCript("D", CSTR( pass &""))
   cnad.Properties("Encrypt Password") = TRUE

   cnAd.OPEN

   strAttributes = "name,distinguishedName"

   strFilter = "(&(objectCategory=group) " & "(Name=" & idgrupo &"))"


   strQuery = strBase1 & ";" & strFilter & ";" & strAttributes & ";subtree"
   '  response.write strQuery
   SET rsad = cnAd.EXECUTE(strQuery)

   IF NOT rsad.EOF THEN
      FullNameGrupoEx = rsad("distinguishedName")
   ELSE
      FullNameGrupoEx = idgrupo
   END IF

   rsad.CLOSE
   SET rsad = NOTHING
   cnAd.CLOSE
   SET cnad = NOTHING

END FUNCTION

'----------RBAC------------------------------------------------RBAC-----------------------------------------------------------------


FUNCTION getinfoUserADBCRED(legajo)

 Dirc=  Request.ServerVariables("APPL_PHYSICAL_PATH")

 Dirc=Dirc  & "coressdll\getuser.xml"

 set xmlHttp2=SERVER.createobject("Microsoft.XMLHTTP")

    xml4=leerArc(dirc)
    xml4=REPLACE(XML4,"XUSUARIO",legajo)
    xml4=REPLACE(XML4,"XAPP",session("USRPRV"))


    xmlHttp2.open "POST", sRBAC, FALSE
    xmlHttp2.send (xml4 )

    set xml3 =xmlHTTP2.responsexml

   ' Set xmldoc=xml3.selectSingleNode("//response")
	'Set z1=xmldoc.selectSingleNode("//error")
   ' serror= cbool(z1.text)

    Set xmldoc=xml3.selectSingleNode("//userData")
    if Not (xmldoc Is Nothing)  THEN
            session("SGU_userid") = legajo

            Set z1=xmldoc.selectSingleNode("//dsFirstName")
		    nombre=z1.text
		    Set z2=xmldoc.selectSingleNode("//dsLastName")
            apellido=z2.text

            session("SGU_nombre") = nombre & " " & apellido
            session("SGU_mail") = GETMAIL(legajo)
            session("SGU_descrisector") = "MySect"
            session("SGU_idcostcenter") = "0"
            session("SGU_descripuesto") = "NO"
			session("SGU_codigopuesto") = "NO"
            session("SGU_idnivel")      ="8"
         ELSE
            session("SGU_userid") = Legajo
            session("SGU_nombre") = "NO"
            session("SGU_mail") = "NO"
            session("SGU_descrisector") = "NO"
            session("SGU_idcostcenter") = "0"
            session("SGU_descripuesto") = "NO"
            session("SGU_codigopuesto") = "NO"
            session("SGU_idnivel")      ="8"
         END IF

  set xmlHTTP2=nothing

END FUNCTION

'-------------------------------------------------------------------------------------

FUNCTION getinfoUserWBCRED(legajo)

 Dirc=  Request.ServerVariables("APPL_PHYSICAL_PATH")

 Dirc=Dirc  & "coressdll\getuser.xml"

 set xmlHttp2=SERVER.createobject("Microsoft.XMLHTTP")

    xml4=leerArc(dirc)
    xml4=REPLACE(XML4,"XUSUARIO",legajo)
    xml4=REPLACE(XML4,"XAPP",session("USRPRV"))


    xmlHttp2.open "POST", sRBAC, FALSE
    xmlHttp2.send (xml4 )

    set xml3 =xmlHTTP2.responsexml

   ' Set xmldoc=xml3.selectSingleNode("//response")
	'Set z1=xmldoc.selectSingleNode("//error")
   ' serror= ucase(z1.text&"")

   ' if serror="FALSE" THEN
    Set xmldoc=xml3.selectSingleNode("//userData")
    if Not (xmldoc Is Nothing)  THEN

            session("SGU_userid") = legajo

            Set z1=xmldoc.selectSingleNode("//dsFirstName")
		    nombre=z1.text
		    Set z2=xmldoc.selectSingleNode("//dsLastName")
            apellido=z2.text

            session("SGU_nombre") = nombre & " " & apellido
            session("SGU_mail") = GETMAIL(legajo)
            session("SGU_descrisector") = "MySect"
            session("SGU_idcostcenter") = "0"
            session("SGU_descripuesto") = "NO"
			session("SGU_codigopuesto") = "NO"
            session("SGU_idnivel")      ="8"
         ELSE
            session("SGU_userid") = Legajo
            session("SGU_nombre") = "NO"
            session("SGU_mail") = "NO"
            session("SGU_descrisector") = "NO"
            session("SGU_idcostcenter") = "0"
            session("SGU_descripuesto") = "NO"
            session("SGU_codigopuesto") = "NO"
            session("SGU_idnivel")      ="8"
         END IF

  set xmlHTTP2=nothing

END FUNCTION

'-------------------------------------------------------------------------------------

FUNCTION getGroupsUserBCRED(legajo)

 session("SGU_grupos")=""
 xgusep=""
 Dirc=  Request.ServerVariables("APPL_PHYSICAL_PATH")

 Dirc=Dirc  & "coressdll\getroles.xml"

 set xmlHttp2=SERVER.createobject("Microsoft.XMLHTTP")

    xml4=leerArc(dirc)
    xml4=REPLACE(XML4,"XUSUARIO",legajo)
    xml4=REPLACE(XML4,"XAPP",session("USRPRV"))


    xmlHttp2.open "POST", sRBAC, FALSE
    xmlHttp2.send (xml4)

    set xml3 =xmlHTTP2.responsexml


    Set xmldoc=xml3.selectSingleNode("//return")


     Set nList = xmldoc.selectNodes("roles")

        for Each n In nList
          session("SGU_grupos") = session("SGU_grupos") & xgusep  & ucase(n.text &"")
          xgusep = CHR(9)
        next

 set xmlHTTP2=nothing
 END FUNCTION


 function getInfouserPstBANELCOEX(Legajo,IDPST,tipo)


     session("SGU_userid")      = Legajo
 	session("SGU_nombre")      = "NO"
 	session("SGU_mail")	   = "NO"
 	session("SGU_descrisector")= "NO"
 	session("SGU_idcostcenter")= "0"
 	session("SGU_descripuesto")= "NO"
 	session("SGU_codigopuesto")= "NO"
 	session("SGU_direccion")   ="NO"
 	session("SGU_telefono")   ="NO"
 	session("SGU_DEPTO")   =""
         session("SGU_IMG")=""


         select case tipo


          case "USUARIO"



 	           set xcnPSB = ofv.conectar(ofv.strconn2)

 	           session("SGU_descrisector")="NO"
 	           session("SGU_descripuesto")="NO"

 	           XSQL="    Select C.DESCRI,C1.DESCRI From PUESTOREL "
 	   	       XSQL=XSQL & "  INNER JOIN LISTADEPUESTOS C ON(C.OBJID=PUESTOREL.IDPST) "
 	   	       XSQL=XSQL & "   	   LEFT JOIN ECO C1 ON (C1.ECO=C.SECTOR) "
 	           XSQL=XSQL & " WHERE PUESTOREL.idpst =" &  IDPST

 	              Set xrs2 = OFV.crearconsultaEx(XSQL, xcnPSB, 1, "")

 	        	 xx=getinfouser(legajo)

                 session("SGU_codigopuesto")= IDPST
                 session("SGU_DEPTO")=xrs2(1)




 	              call ofv.cerrarconsulta(xrs2)

            call ofv.cerrarconn(xcnPSB)






          CASE "PUESTO"

            xgusep                   = ""
 	   session("SGU_UIGuserid") = ""
 	   session("SGU_UIGdescri") = ""

 	   select case GTUcase

 	        case "I"

 	            set xcn = ofv.conectar(ofv.strconn2)

 	            xStrSql = "Select cod From relapuestos WHERE idpst = '" & Legajo
 	            Set xrs = OFV.crearconsultaEx(xStrSql, xcn, 1, "")
 	            IF NOT xRS.EOF THEN

 	               DO WHILE NOT XRS.EOF

 	                 session("SGU_UIGuserid") = session("SGU_UIGuserid") & xgusep & xrs(0)
 	                 xx=getinfouser(xrs(0) & "")
 	                 session("SGU_UIGdescri") = session("SGU_UIGdescri") & xgusep & session("SGU_nombre")

 	                 xgusep                   = chr(9)

 	               XRS.MOVENEXT
 	               LOOP

 	            ELSE

 	            END IF

 	            call ofv.cerrarconsulta(xrs)
 	            call ofv.cerrarconn(xcn)
             END SELECT

          END SELECT
 end function
'---------------------------------------------------------------------------


'---------------------------------------------------------------------------

function getInfouserPstBcEX(Legajo,IDPST,tipo)


        session("SGU_userid")      = Legajo
	session("SGU_nombre")      = "NO"
	session("SGU_mail")	   = "NO"
	session("SGU_descrisector")= "NO"
	session("SGU_idcostcenter")= "0"
	session("SGU_descripuesto")= "NO"
	session("SGU_codigopuesto")= "NO"
	session("SGU_direccion")   ="NO"
	session("SGU_telefono")   ="NO"
	session("SGU_DEPTO")   =""
        session("SGU_IMG")=""


        select case tipo


         case "USUARIO"

           set xcnPS = ofv.conectar(ofv.strconn7)
           set xcnPSB = ofv.conectar(ofv.strconn2)



            xStrSql="select * from PS_BC_PS_AD_VW where	RTRIM(UPPER(ALTER_EMPLID))='" & UCASE(Legajo & "") &"'"

         ' response.write xStrSql

           session("SGU_descrisector")="NO"
           session("SGU_descripuesto")="NO"
           Set xrs = OFV.crearconsultaEx(xStrSql, xcnPS, 1, "")
           IF NOT xRS.EOF THEN

           XSQL="    Select C.DESCRI,C1.DESCRI From PUESTOREL "
	   XSQL=XSQL & "  INNER JOIN LISTADEPUESTOS C ON(C.OBJID=PUESTOREL.IDPST) "
	   XSQL=XSQL & "   	   LEFT JOIN ECO C1 ON (C1.ECO=C.SECTOR) "
           XSQL=XSQL & " WHERE PUESTOREL.idpst =" &  IDPST
           'RESPONSE.WRITE XSQL
           Set xrs2 = OFV.crearconsultaEx(XSQL, xcnPSB, 1, "")

     		session("SGU_userid")      = trim(xrs("ALTER_EMPLID")&"")
  		session("SGU_nombre")      = replace(trim(xrs("NAME1")&""),"'"," ")
  		session("SGU_mail")	   = trim(xrs("EMAILID") &"")
  		'session("SGU_descrisector")= trim(xrs2(1)&"")&"&nbsp;"

  		IF NOT XRS2.EOF THEN
		  		session("SGU_descrisector")= trim(xrs2(1)&"")&"&nbsp;"
		  		session("SGU_descripuesto")= trim(xrs2(0) &"")&"&nbsp;"
		  		ELSE
		  		session("SGU_descrisector")="NO"
		  		session("SGU_descripuesto")="NO"


  		END IF
  		session("SGU_idcostcenter")= xrs(3)
  		'session("SGU_descripuesto")= trim(xrs2(0) &"")&"&nbsp;"
  		session("SGU_codigopuesto")= IDPST
  		session("SGU_direccion")   =trim(xrs("DESCR200")&"")
  		session("SGU_telefono")   =trim(xrs("DESCR254_MIXED")&"")
  		session("SGU_DEPTO")   =trim(xrs("DESCR") &"-" & xrs("DESCR1")&"")
  		ximg=trim(xrs("BC_DESCR250_1") &"")
  		session("SGU_IMG")   =right(ximg,len(ximg)- instrrev(ximg,"\"))

           ELSE
     		session("SGU_userid")      = Legajo
  		session("SGU_nombre")      = "NO"
  		session("SGU_mail")	   = "NO"
  		session("SGU_descrisector")= "NO"
  		session("SGU_idcostcenter")= "0"
  		session("SGU_descripuesto")= "NO"
  		session("SGU_codigopuesto")= "NO"
  		session("SGU_direccion")   ="NO"
  		session("SGU_telefono")   ="NO"
  		session("SGU_DEPTO")   =""
  		session("SGU_IMG")=""
           END IF

           call ofv.cerrarconsulta(xrs)
           call ofv.cerrarconn(xcnPS)
           call ofv.cerrarconn(xcnPSB)


         CASE "PUESTO"

           xgusep                   = ""
	   session("SGU_UIGuserid") = ""
	   session("SGU_UIGdescri") = ""

	   select case GTUcase

	        case "I"

	            set xcn = ofv.conectar(ofv.strconn2)

	            xStrSql = "Select cod From relapuestos WHERE idpst = '" & Legajo
	            Set xrs = OFV.crearconsultaEx(xStrSql, xcn, 1, "")
	            IF NOT xRS.EOF THEN

	               DO WHILE NOT XRS.EOF

	                 session("SGU_UIGuserid") = session("SGU_UIGuserid") & xgusep & xrs(0)
	                 xx=getinfouser(xrs(0) & "")
	                 session("SGU_UIGdescri") = session("SGU_UIGdescri") & xgusep & session("SGU_nombre")

	                 xgusep                   = chr(9)

	               XRS.MOVENEXT
	               LOOP

	            ELSE

	            END IF

	            call ofv.cerrarconsulta(xrs)
	            call ofv.cerrarconn(xcn)
            END SELECT

         END SELECT
end function
'---------------------------------------------------------------------------




function getUsrInTEXTBCEX(XGU_TEXT)



   set xcn = ofv.conectar(ofv.strconn7)

   XGU_TEXT=replace(ucase(XGU_TEXT & "" ),"'","''")

    xStrSql="select ALTER_EMPLID,NAME1 from PS_BC_PS_AD_VW where upper(NAME1) LIKE '" & XGU_TEXT &"%' ORDER BY NAME1"
   ' response.write xStrSql
      Set xrs = OFV.crearconsultaEx(xStrSql, xcn, 1, "")




   xgusep                   = ""
   session("SGU_UIGuserid") = ""
   session("SGU_UIGdescri") = ""




	    IF NOT xRS.EOF THEN

	       DO WHILE NOT XRS.EOF

		 session("SGU_UIGuserid") = session("SGU_UIGuserid") & xgusep & trim(xrs(0)&"")

		 session("SGU_UIGdescri") = session("SGU_UIGdescri") & xgusep & replace(trim(XRS(1)&""),"'" ," ")

		 xgusep                   = chr(9)

	       XRS.MOVENEXT
	       LOOP

	    ELSE

	    END IF

	    call ofv.cerrarconsulta(xrs)
	    call ofv.cerrarconn(xcn)

end function
'---------------------------------------------------------------------------

FUNCTION getUsrIngroupBCRED(GRUPO)


session("SGU_grupos")=""
xgusep=""



 SET cnAD = ofv.conectar(ofv.strconn2)
      Strsql = "Select MapGrp From grupos where id='" & grupo & "'"
      SET rsADX = ofv.crearconsultaEx(StrSql, cnAd, 1, parametros)
         IF NOT rsADX.EOF THEN
             GrupoADX=rsADX(0)

         ELSE
            GrupoADX = ""
         END IF

  CALL ofv.cerrarconsulta(rsADX)

 Dirc=  Request.ServerVariables("APPL_PHYSICAL_PATH")

 Dirc=Dirc  & "coressdll\USERWITHROLEv1.xml"

 set xmlHttp2=SERVER.createobject("Microsoft.XMLHTTP")

    xml4=leerArc(dirc)
    xml4=REPLACE(XML4,"XROLE",GrupoADX)
    xml4=REPLACE(XML4,"XAPP",session("USRPRV"))


    xmlHttp2.open "POST", sRBAC, FALSE
    xmlHttp2.send (xml4)

    set xml3 =xmlHTTP2.responsexml


    Set xmldoc=xml3.selectSingleNode("//return")


     Set nList = xmldoc.selectNodes("users")

        for Each n In nList
          session("SGU_UIGuserid") = session("SGU_UIGuserid") & xgusep & n.text
          xxx=getinfoUserADBCRED(n.text)
		  session("SGU_UIGdescri") = session("SGU_UIGdescri") & xgusep & session("SGU_nombre")
          xgusep = CHR(9)
        next

 set xmlHTTP2=nothing




END FUNCTION

'--------------------------------------------------------------------------

FUNCTION GETMAIL(LEGAJO)

Dirc=  Request.ServerVariables("APPL_PHYSICAL_PATH")

 Dirc=Dirc  & "coressdll\getmailLdap.xml"

 set xmlHttp2=SERVER.createobject("Microsoft.XMLHTTP")

 xml4=leerArc(dirc)
 xml4=REPLACE(XML4,"XUSUARIO",legajo)



 xmlHttp2.open "POST", sMailLdap, FALSE
 xmlHttp2.send (xml4)

  set xml3 =xmlHTTP2.responsexml

  Set xmldoc=xml3.selectSingleNode("//return")
  Set n = xmldoc.selectSingleNode("mail")
  if Not (n Is Nothing) then
     GETMAIL=n.text
  else
     getmail="NO"
  end if

 set xmlHTTP2=nothing

END FUNCTION


function getsectorExt(Sdato)

 'Sdato=mid(sdato,1,11)

 IF SDATO & "" ="" THEN
 getsectorExt="NO"
 ELSE
 SET cnext = ofv.conectar(ofv.strconn0)

 Strsqle3 = "Select * From eco WHERE codinter='" & sdato & "'"

 SET rsdomeco = ofv.crearconsultaEx(StrSqle3, cnext, 1, parametros)

 if not rsdomeco.eof then
   getsectorExt=rsdomeco(0)
 else
  getsectorExt="0" 'Sdato
 end if
 CALL ofv.cerrarconsulta(rsdomeco)
 CALL ofv.cerrarconn(cnext)
END IF
end function





function getnivelExt(Sdato)

' Sdato=RTRIM(Sdato & "")
 SET cnext = ofv.conectar(ofv.strconn0)

 Strsqle3 = "Select Nccod  From NivelConfiden WHERE CODEXTERNO  >=0" & sdato

 SET rsdomeco = ofv.crearconsultaEx(StrSqle3, cnext, 1, parametros)

 if not rsdomeco.eof then
   getnivelExt=rsdomeco(0)
 else
   getnivelExt ="9" 'Sdato
 end if
 CALL ofv.cerrarconsulta(rsdomeco)
 CALL ofv.cerrarconn(cnext)

end function

function getnivelExtBanelco(Sdato)

 Sdato=RTRIM(Sdato & "")
 IF  SDATO="" THEN
 getnivelExtBanelco="9"
 ELSE
 SET cnext = ofv.conectar(ofv.strconn0)

 Strsqle3 = "Select Nccod  From NivelConfiden WHERE CODEXTERNO  ='" & sdato & "'"

 SET rsdomeco = ofv.crearconsultaEx(StrSqle3, cnext, 1, parametros)

 if not rsdomeco.eof then
   getnivelExtBanelco=rsdomeco(0)
 else
   getnivelExtBanelco ="9" 'Sdato
 end if
 CALL ofv.cerrarconsulta(rsdomeco)
 CALL ofv.cerrarconn(cnext)
END IF
end function


function getperfilExtBanelco(Sdato)

 Sdato=RTRIM(Sdato & "")
 IF SDATO="" THEN
 getperfilExtBanelco="NO"
 ELSE
 SET cnext = ofv.conectar(ofv.strconn0)

 Strsqle3 = "Select tipoPerfil  From TipoPerfil WHERE CODpuestoE  ='" & sdato & "'"
'RESPONSE.WRITE Strsqle3
 SET rsdomeco = ofv.crearconsultaEx(StrSqle3, cnext, 1, parametros)

 if not rsdomeco.eof then
   getperfilExtBanelco=rsdomeco(0)
 else
   getperfilExtBanelco ="NO" 'Sdato
 end if
 CALL ofv.cerrarconsulta(rsdomeco)
 CALL ofv.cerrarconn(cnext)
END IF
end function



FUNCTION getInfouserBCEx(legajo)

  ' legajo=RTRIM(legajo &"")
   SET cnDOM = ofv.conectar(ofv.strconn0)
   Strsqlq = "Select * From dominios WHERE RAIZLDAP IS NOT NULL and DESCRIPCION <> 'DEFAULT'"
    SET rsdom = ofv.crearconsultaEx(StrSqlq, cnDOM, 1, parametros)

  do while (not rsdom.eof )
   sldap =rsdom("raizldap")
   usernx =  rsdom("usrldap")
   passx = cstr(ofv.getCript("D",rsdom("pwdldap")&""))'rsdom("pwdldap")

   SET connaX = server.CREATEOBJECT("ADODB.Connection")


   connaX.provider = "ADSDSOObject"
   connaX.Properties("User ID") = usernx
   connaX.Properties("Password") = passx
   connaX.Properties("Encrypt Password") = TRUE

   connaX.OPEN

   strAttributes = "sAMAccountName"

   strFilter = "(&(objectCategory=person)(objectClass=user)" & "(samAccountName=*" & legajo &"))"


   strQuery = sldap & ";" & strFilter & ";" & strAttributes & ";subtree"

   SET RSAX = connaX.EXECUTE(strQuery)


   IF RSAX.EOF THEN
      getInfouserBCEx=legajo
   ELSE
      getInfouserBCEx=RSAX("samAccountName")
    exit do
  END IF

   RSAX.CLOSE
   SET RSAX = NOTHING
   connaX.CLOSE
   SET connaX = NOTHING
   rsdom.MOVENEXT
   LOOP
   CALL ofv.cerrarconsulta(rsdom)
   CALL ofv.cerrarconn(cnDOM)
END FUNCTION

Function ChangeMail(Smail)
    Smail = Smail & ""
    Armail = Split(Smail, "@")

    If UBound(Armail) <> -1 Then
        ChangeMail = Armail(0) & "@ciudad.bco"
    Else
        ChangeMail = "NO"
    End If

End Function


function GrupoDelUsrLoguin(Legajo)
	' Create the http request text
	Dim strSoapReq
	strSoapReq = GenerateSoapBodyStart()
	strSoapReq = strSoapReq & GenerateSoapFunctionCallStart("GrupoDelLegajo")

	strSoapReq = strSoapReq & "<GrupoDelLegajoRequest>" & GenerateSoapParameter("Legajo", legajo) & "</GrupoDelLegajoRequest>"
	strSoapReq = strSoapReq & GenerateSoapFunctionCallEnd("GrupoDelLegajo")
	strSoapReq = strSoapReq & GenerateSoapBodyEnd()

	Dim oHttp
	Dim strResult
	Set oHttp = CreateObject("Msxml2.XMLHTTP")
	oHttp.open "POST", URLRiva, false
	oHttp.setRequestHeader "Content-Type", "text/xml"
	oHttp.setRequestHeader "SOAPAction", URLRiva & "GrupoDelLegajo"
	oHttp.send strSoapReq
	'response.write oHttp.responseText
	SET strResult = oHttp.responseXML
    xgusep=""
    GrupoDelUsrLoguin=""
  Set nList = strResult.selectNodes("//Grupo")

 	   	Dim iLoop
 	 	   For Each n In nList
 	 	       if n.text <>"" then
 	 	       GrupoDelUsrLoguin = GrupoDelUsrLoguin & xgusep  &  n.text &""
               xgusep = CHR(9)

 	 	       end if

	Next

	Set oHttp=NOTHING
    SET strResult=NOTHING
end function


function LegajosSector(sector)
	' Create the http request text
	Dim strSoapReq
	strSoapReq = GenerateSoapBodyStart()
	strSoapReq = strSoapReq & GenerateSoapFunctionCallStart("LegajosSector")

	strSoapReq = strSoapReq & "<LegajosSectorRequest>" & GenerateSoapParameter("Sector", sector) & "</LegajosSectorRequest>"
	strSoapReq = strSoapReq & GenerateSoapFunctionCallEnd("LegajosSector")
	strSoapReq = strSoapReq & GenerateSoapBodyEnd()

	Dim oHttp
	Dim strResult
	Set oHttp = CreateObject("Msxml2.XMLHTTP")
	oHttp.open "POST", URLRiva, false
	oHttp.setRequestHeader "Content-Type", "text/xml"
	oHttp.setRequestHeader "SOAPAction", URLRiva & "LegajosSector"
	oHttp.send strSoapReq
'	response.write oHttp.responseText
	SET strResult = oHttp.responseXML
Set nList = strResult.selectNodes("//Empleados/item")
  xgusep=""
  session("SGU_UIGuserid")=""
  session("SGU_UIGdescri")=""
	   	Dim iLoop
	 	    For Each n In nList
		   	 	    Legajo=n.selectSingleNode("Legajo").text &""

		  		     nombre=n.selectSingleNode("Nombre").text & ""

		  		   if Legajo <>"" then
		  			session("SGU_UIGuserid") = session("SGU_UIGuserid") & xgusep & Legajo '& objRecordSet("sAMAccountName")
		  			session("SGU_UIGdescri") = session("SGU_UIGdescri") & xgusep & nombre
		  			xgusep = CHR(9)



		  	 	   end if

	Next



 session("SGU_UIGdescri")=Konvert(session("SGU_UIGdescri"))

end function

function GrupoDelLegajo(Legajo)
	' Create the http request text
	Dim strSoapReq
	strSoapReq = GenerateSoapBodyStart()
	strSoapReq = strSoapReq & GenerateSoapFunctionCallStart("GrupoDelLegajo")

	strSoapReq = strSoapReq & "<GrupoDelLegajoRequest>" & GenerateSoapParameter("Legajo", legajo) & "</GrupoDelLegajoRequest>"
	strSoapReq = strSoapReq & GenerateSoapFunctionCallEnd("GrupoDelLegajo")
	strSoapReq = strSoapReq & GenerateSoapBodyEnd()

	Dim oHttp
	Dim strResult
	Set oHttp = CreateObject("Msxml2.XMLHTTP")
	oHttp.open "POST", URLRiva, false
	oHttp.setRequestHeader "Content-Type", "text/xml"
	oHttp.setRequestHeader "SOAPAction", URLRiva & "GrupoDelLegajo"
	oHttp.send strSoapReq
	'response.write oHttp.responseText
	SET strResult = oHttp.responseXML
    xgusep=""
    session("SGU_grupos")=""
  Set nList = strResult.selectNodes("//Grupo")

 	   	Dim iLoop
 	 	   For Each n In nList
 	 	       if n.text <>"" then
 	 	       session("SGU_grupos") = session("SGU_grupos") & xgusep  &  n.text &""
               xgusep = CHR(9)

 	 	       end if

	Next

	Set oHttp=NOTHING
    SET strResult=NOTHING
end function

Function Konvert(sIn)
    Dim oIn
    Set oIn = SERVER.CreateObject("ADODB.Stream")
    oIn.Open
    oIn.CharSet = "ISO-8859-1"
    oIn.WriteText sIn
    oIn.Position = 0
    oIn.CharSet =  "UTF-8"
    Konvert = oIn.ReadText
    oIn.Close
    SET OIN=NOTHING
End Function

function LegajosGrupo(grupo)
	' Create the http request text
	Dim strSoapReq

	SET cnAD = ofv.conectar(ofv.strconn2)
	     Strsql = "Select MapGrp From grupos where id='" & grupo & "'"
	     SET rsADX = ofv.crearconsultaEx(StrSql, cnAd, 1, parametros)
	        IF NOT rsADX.EOF THEN
	           grupo=rsADX(0)

	        END IF

     CALL ofv.cerrarconsulta(rsADX)


	strSoapReq = GenerateSoapBodyStart()
	strSoapReq = strSoapReq & GenerateSoapFunctionCallStart("LegajosGrupo")

	strSoapReq = strSoapReq & "<LegajosGrupoRequest>" & GenerateSoapParameter("Grupo", grupo) & "</LegajosGrupoRequest>"
	strSoapReq = strSoapReq & GenerateSoapFunctionCallEnd("LegajosGrupo")
	strSoapReq = strSoapReq & GenerateSoapBodyEnd()

	Dim oHttp
	Dim strResult
	Set oHttp = CreateObject("Msxml2.XMLHTTP")
	oHttp.open "POST", URLRiva, false
	oHttp.setRequestHeader "Content-Type", "text/xml"
	oHttp.setRequestHeader "SOAPAction", URLRiva & "LegajosGrupo"
	oHttp.send strSoapReq
	'response.write oHttp.responseText
	SET strResult = oHttp.responseXML

  Set nList = strResult.selectNodes("//Empleados/item")
    xgusep=""
    session("SGU_UIGuserid")=""
    session("SGU_UIGdescri")=""
  	   	Dim iLoop
  	 	    For Each n In nList
  		   	 	    Legajo=n.selectSingleNode("Legajo").text &""

  		  		     nombre=n.selectSingleNode("Nombre").text & ""

  		  		   if Legajo <>"" then
  		  			session("SGU_UIGuserid") = session("SGU_UIGuserid") & xgusep & Legajo '& objRecordSet("sAMAccountName")
  		  			session("SGU_UIGdescri") = session("SGU_UIGdescri") & xgusep & nombre
  		  			xgusep = CHR(9)



  		  	 	   end if

	Next

     session("SGU_UIGdescri")=Konvert(session("SGU_UIGdescri"))
	Set oHttp=NOTHING
    SET strResult=NOTHING
end function

function Loguin(p1,p2)
	' Create the http request text
	Dim strSoapReq
	strSoapReq = GenerateSoapBodyStart()
	strSoapReq = strSoapReq & GenerateSoapFunctionCallStart("Login")

	strSoapReq = strSoapReq & "<LoginRequest>" & GenerateSoapParameter("Legajo", p1)
	strSoapReq = strSoapReq & GenerateSoapParameter("Identificacion", p2) &   "</LoginRequest>"
	strSoapReq = strSoapReq & GenerateSoapFunctionCallEnd("Login")
	strSoapReq = strSoapReq & GenerateSoapBodyEnd()

	Dim oHttp
	Dim strResult
	Set oHttp = SERVER.CreateObject("Msxml2.XMLHTTP")
	oHttp.open "POST", URLRiva, false
	oHttp.setRequestHeader "Content-Type", "text/xml"
	oHttp.setRequestHeader "SOAPAction", URLRiva & "Login"
	oHttp.send strSoapReq
	'response.write oHttp.responseText
	SET strResult = oHttp.responseXML


    ErrorT=strResult.selectSingleNode("//Error ").Text &""

	Loguin=Konvert(ErrorT)
    Set oHttp=NOTHING
    SET strResult=NOTHING
end function





	Function GetResult(byval responseText, byval resultParam)

		Dim oXml
		Set oXml = CreateObject("Msxml2.DOMDocument")
		oXml.Async = true
		oXml.LoadXml responseText

		Dim strPath
		strPath = "/*/*/*/" & resultParam
		Dim oNode
		Set oNode = oXml.documentElement.SelectSingleNode(strPath)
		GetResult = oNode.Text

	End Function


	Function GenerateSoapBodyStart()

		Dim strSoap
		strSoap = "<?xml version=""1.0"" encoding=""utf-8""?>"
		strSoap = strSoap & "<soap12:Envelope "
		strSoap = strSoap & "xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" "
		strSoap = strSoap & "xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" "
		strSoap = strSoap & "xmlns:soap12=""http://www.w3.org/2003/05/soap-envelope""> "
		strSoap = strSoap & "<soap12:Body>"
		GenerateSoapBodyStart = strSoap

	End Function

	Function GenerateSoapBodyEnd()

		Dim strSoap
		strSoap = "</soap12:Body>"
		strSoap = strSoap & "</soap12:Envelope>"
		GenerateSoapBodyEnd = strSoap

	End Function

	Function GenerateSoapFunctionCallStart(byval strFunction)

		Dim strSoap
		strSoap = "<" & strFunction & " xmlns=""" & nsUrl & """>"
		GenerateSoapFunctionCallStart = strSoap

	End Function

	Function GenerateSoapFunctionCallEnd(byval strFunction)

		Dim strSoap
		strSoap = "</" & strFunction & ">"
		GenerateSoapFunctionCallEnd = strSoap

	End Function

	Function GenerateSoapParameter(byval strParam, byval strValue)

		Dim strSoap
		strSoap = "<" & strParam & ">"
		strSoap = strSoap & strValue
		strSoap = strSoap & "</" + strParam & ">"
		GenerateSoapParameter = strSoap

End Function





function LOGUINEx(Legajo)
	' Create the http request text
	Dim strSoapReq
	strSoapReq = GenerateSoapBodyStart()
	strSoapReq = strSoapReq & GenerateSoapFunctionCallStart("LoginLegajo")

	strSoapReq = strSoapReq & "<LoginLRequest >" & GenerateSoapParameter("Legajo", legajo) & "</LoginLRequest >"
	strSoapReq = strSoapReq & GenerateSoapFunctionCallEnd("LoginLegajo")
	strSoapReq = strSoapReq & GenerateSoapBodyEnd()
 ' RESPONSE.WRITE strSoapReq
	Dim oHttpstrSoapReq
	Dim strResult
	Set oHttp = server.CreateObject("Msxml2.XMLHTTP")
	oHttp.open "POST", URLRiva, false
	oHttp.setRequestHeader "Content-Type", "text/xml"
	oHttp.setRequestHeader "SOAPAction", URLRiva & "LoginLegajo"
	oHttp.send strSoapReq

	SET strResult = oHttp.responseXML



     if Not (strResult Is Nothing)  THEN


				LOGUINEx=Konvert(strResult.selectSingleNode("//Error").Text &"")
                IF LOGUINEx ="" THEN
                 LOGUINEx="OK"
                 Session("Usuario")   = Legajo
                 call InfousrLOGUIN(Legajo)
                END IF

     ELSE

     LOGUINEx="NO"

     END IF


Set oHttp=NOTHING
SET strResult=NOTHING



end function

function InfoUsrLOGUIN(Legajo)
	' Create the http request text
	Dim strSoapReq
	strSoapReq = GenerateSoapBodyStart()
	strSoapReq = strSoapReq & GenerateSoapFunctionCallStart("InfoDelLegajo")

	strSoapReq = strSoapReq & "<InfoDelLegajoRequest>" & GenerateSoapParameter("Legajo", legajo) & "</InfoDelLegajoRequest>"
	strSoapReq = strSoapReq & GenerateSoapFunctionCallEnd("InfoDelLegajo")
	strSoapReq = strSoapReq & GenerateSoapBodyEnd()
 ' RESPONSE.WRITE strSoapReq
	Dim oHttpstrSoapReq
	Dim strResult
	Set oHttp = server.CreateObject("Msxml2.XMLHTTP")
	oHttp.open "POST", URLRiva, false
	oHttp.setRequestHeader "Content-Type", "text/xml"
	oHttp.setRequestHeader "SOAPAction", URLRiva & "InfoDelLegajo"
	oHttp.send strSoapReq

	SET strResult = oHttp.responseXML

     if Not (strResult Is Nothing)  THEN
	            session("SGU_userid") = legajo

				UsuarioNombre=Konvert(strResult.selectSingleNode("//UsuarioNombre").Text &"")
				UsuarioMail=Konvert(strResult.selectSingleNode("//UsuarioMail").Text &"")
				UsuarioGrupo=strResult.selectSingleNode("//UsuarioGrupo").Text &""
				UsuarioSector=strResult.selectSingleNode("//UsuarioSector").Text &""
				UsuarioNivel=strResult.selectSingleNode("//UsuarioNivel").Text &""
				UsuarioPerfil=strResult.selectSingleNode("//UsuarioPerfil").Text &""
				Error=Konvert(strResult.selectSingleNode("//Error ").Text &"")


	            	Session("Usuario")   = legajo
					Session("UsrOpcion") = UsuarioNombre 'session("SGU_nombre")
					Session("SkrDefT")   = "0"
					Session("SkrDefE")   = "0"
					Session("SkrDefI")   = "0"
					Session("UsrPmin")   = "0"
					Session("UsrPmax")   = "0"
					Session("UsrHist")   = "0"
					Session("SegUsuario") = UsuarioNombre 'session("SGU_nombre")
					Session("SegNivel")  = getnivelExtRiva(UsuarioNivel)
					Session("SegGrp")    = "DEF"
					Session("SegPerfil") = getperfilExtRiva(UsuarioPerfil) 'Session("Usuario")
					Session("UsuarNT")   = legajo 'Session("Usuario")

					Session("UsrSector") = getsectorExtRiva(UsuarioSector) '"" 'session("SGU_idcostcenter")
					Session("UsrSecDsc") = "" 'session("SGU_descrisector")
					Session("UsrPuesto") = "" 'session("SGU_codigopuesto")
					Session("UsrPstDsc") = "" 'session("SGU_descripuesto")
					Session("UsrMail")   = UsuarioMail 'session("SGU_mail")

					Session("SegGrpOpt")    = "S"
				   	Session("SegGrpLst")    = GrupoDelUsrLoguin(Legajo)


         END IF


Set oHttp=NOTHING
SET strResult=NOTHING



end function

Function ValidaMail(email)
'Funcion que comprueba si un e-mail es valido. Devuelve:
' 0 -> si es nulo o esta vacio
' 1 -> Si el e-mail es incorrecto
' 2 -> Si el e-mail es valido

    If(email)="" Then
        ValidaMail=0
        exit function
   End if
dim partes, parte, i, c
    'rompo el email en dos partes, antes y después de la arroba
    partes = Split(email, "@")
    if UBound(partes) <> 1 then
       'si el mayor indice del array es distinto de 1 es que no he obtenido las dos partes
       ValidaMail = 1
       exit function
    end if
    'para cada parte, compruebo varias cosas
    for each parte in partes
       'Compruebo que tiene algún caracter
       if Len(parte) <= 0 then
          ValidaMail = 1
          exit function
       end if
       'para cada caracter de la parte
       for i = 1 to Len(parte)
          'tomo el caracter actual
          c = Lcase(Mid(parte, i, 1))
          'miro a ver si ese caracter es uno de los permitidos
          if InStr("._-abcdefghijklmnopqrstuvwxyz", c) <= 0 and not IsNumeric(c) then
             ValidaMail = 1
             exit function
          end if
       next
       'si la parte actual acaba o empieza en punto la dirección no es válida
       if Left(parte, 1) = "." or Right(parte, 1) = "." then
          ValidaMail = 1
          exit function
       end if
    next
    'si en la segunda parte del email no tenemos un punto es que va mal
    if InStr(partes(1), ".") <= 0 then
       ValidaMail = 1
       exit function
    end if
    'calculo cuantos caracteres hay después del último punto de la segunda parte del mail
    i = Len(partes(1)) - InStrRev(partes(1), ".")
    'si el número de caracteres es distinto de 2 y 3
    if not (i = 2 or i = 3) then
       ValidaMail = 1
       exit function
    end if
    'si encuentro dos puntos seguidos tampoco va bien
    if InStr(email, "..") > 0 then
       ValidaMail=0
       exit function
    end if
    ValidaMail = 2
End Function

function getnivelExtBCR(Sdato)

' Sdato=RTRIM(Sdato & "")
 SET cnext = ofv.conectar(ofv.strconn0)

 Strsqle3 = "Select Nccod  From NivelConfiden WHERE CODEXTERNO  ='" & sdato & "'"

 SET rsdomeco = ofv.crearconsultaEx(StrSqle3, cnext, 1, parametros)

 if not rsdomeco.eof then
   getnivelExtBCR=cdbl("0" & rsdomeco(0))
 else
   getnivelExtBCR =cdbl("8") 'Sdato
 end if
 CALL ofv.cerrarconsulta(rsdomeco)
 CALL ofv.cerrarconn(cnext)

end function




function InfoDelLegajo(Legajo)
	' Create the http request text
	Dim strSoapReq
	strSoapReq = GenerateSoapBodyStart()
	strSoapReq = strSoapReq & GenerateSoapFunctionCallStart("InfoDelLegajo")

	strSoapReq = strSoapReq & "<InfoDelLegajoRequest>" & GenerateSoapParameter("Legajo", legajo) & "</InfoDelLegajoRequest>"
	strSoapReq = strSoapReq & GenerateSoapFunctionCallEnd("InfoDelLegajo")
	strSoapReq = strSoapReq & GenerateSoapBodyEnd()

	Dim oHttpex2
	Dim strResultex2
	Set oHttpex2 = server.CreateObject("Msxml2.XMLHTTP")
	oHttpex2.open "POST", URLRiva, false
	oHttpex2.setRequestHeader "Content-Type", "text/xml"
	oHttpex2.setRequestHeader "SOAPAction", URLRiva & "InfoDelLegajo"
	oHttpex2.send strSoapReq

	SET strResultex2 = oHttpex2.responseXML

     if Not (strResultex2 Is Nothing)  THEN
	            session("SGU_userid") = legajo

				UsuarioNombre=Konvert(strResultex2.selectSingleNode("//UsuarioNombre").Text &"")
				UsuarioMail=Konvert(strResultex2.selectSingleNode("//UsuarioMail").Text &"")
				UsuarioGrupo=strResultex2.selectSingleNode("//UsuarioGrupo").Text &""
				UsuarioSector=strResultex2.selectSingleNode("//UsuarioSector").Text &""
				UsuarioNivel=strResultex2.selectSingleNode("//UsuarioNivel").Text &""
				UsuarioPerfil=strResultex2.selectSingleNode("//UsuarioPerfil").Text &""
				Error=strResultex2.selectSingleNode("//Error ").Text &""


	            session("SGU_nombre") =UsuarioNombre
	            session("SGU_mail") = UsuarioMail
	            session("SGU_descrisector") =""
	            session("SGU_idcostcenter") = getsectorExtRiva(UsuarioSector)
	            session("SGU_descripuesto") = "NO"
				session("SGU_codigopuesto") =getperfilExtRiva(UsuarioPerfil)
	            session("SGU_idnivel")      =getnivelExtRiva(UsuarioNivel)
	         ELSE
	            session("SGU_userid") = Legajo
	            session("SGU_nombre") = "NO"
	            session("SGU_mail") = "NO"
	            session("SGU_descrisector") = "NO"
	            session("SGU_idcostcenter") = "0"
	            session("SGU_descripuesto") = "NO"
	            session("SGU_codigopuesto") = "NO"
	            session("SGU_idnivel")      ="9"
         END IF


Set oHttpex2=NOTHING
SET strResultex2=NOTHING



end function
function InfoDelLegajow(Legajo)
	' Create the http request text
	Dim strSoapReq
	strSoapReq = GenerateSoapBodyStart()
	strSoapReq = strSoapReq & GenerateSoapFunctionCallStart("InfoDelLegajo")

	strSoapReq = strSoapReq & "<InfoDelLegajoRequest>" & GenerateSoapParameter("Legajo", legajo) & "</InfoDelLegajoRequest>"
	strSoapReq = strSoapReq & GenerateSoapFunctionCallEnd("InfoDelLegajo")
	strSoapReq = strSoapReq & GenerateSoapBodyEnd()
   'response.write  strSoapReq
	Dim oHttpex
	Dim strResultex
	Set oHttpex = server.CreateObject("Msxml2.XMLHTTP")
	oHttpex.open "POST", URLRiva, false
	oHttpex.setRequestHeader "Content-Type", "text/xml"
	oHttpex.setRequestHeader "SOAPAction", URLRiva & "InfoDelLegajo"
	oHttpex.send strSoapReq

	SET strResultex = oHttpex.responseXML

     if Not (strResultex Is Nothing)  THEN
	            session("SGU_userid") = legajo

				UsuarioNombre=Konvert(strResultex.selectSingleNode("//UsuarioNombre").Text &"")
				UsuarioMail=Konvert(strResultex.selectSingleNode("//UsuarioMail").Text &"")
				UsuarioGrupo=strResultex.selectSingleNode("//UsuarioGrupo").Text &""
				UsuarioSector=strResultex.selectSingleNode("//UsuarioSector").Text &""
				UsuarioNivel=strResultex.selectSingleNode("//UsuarioNivel").Text &""
				UsuarioPerfil=strResultex.selectSingleNode("//UsuarioPerfil").Text &""
				Error=strResultex.selectSingleNode("//Error ").Text &""


	            session("SGU_nombre") =UsuarioNombre
	            session("SGU_mail") = UsuarioMail
	            session("SGU_descrisector") =""
	            session("SGU_idcostcenter") = getsectorExtRiva(UsuarioSector)
	            session("SGU_descripuesto") = "NO"
				session("SGU_codigopuesto") = getperfilExtRiva(UsuarioPerfil)
	            session("SGU_idnivel")      ="00" & UsuarioNivel
	         ELSE
	            session("SGU_userid") = Legajo
	            session("SGU_nombre") = "NO"
	            session("SGU_mail") = "NO"
	            session("SGU_descrisector") = "NO"
	            session("SGU_idcostcenter") = "0"
	            session("SGU_descripuesto") = "NO"
	            session("SGU_codigopuesto") = "NO"
	            session("SGU_idnivel")      ="9"
         END IF


Set oHttpex=NOTHING
SET strResultex=NOTHING



end function
function getsectorExtBCR(Sdato)

 'Sdato=mid(sdato,1,11)

 IF SDATO & "" ="" THEN
 getsectorExtBCR="0"
 ELSE
 SET cnext = ofv.conectar(ofv.strconn0)

 Strsqle3 = "Select * From eco WHERE codinter='" & sdato & "'"

 SET rsdomeco = ofv.crearconsultaEx(StrSqle3, cnext, 1, parametros)

 if not rsdomeco.eof then
   getsectorExtBCR=rsdomeco(0)
 else
  getsectorExtBCR="0" 'Sdato
 end if
 CALL ofv.cerrarconsulta(rsdomeco)
 CALL ofv.cerrarconn(cnext)
END IF
end function
function getperfilExtRiva(Sdato)

 Sdato=RTRIM(Sdato & "")
 IF SDATO="" THEN
 getperfilExtRiva="NO"
 ELSE
 SET cnext = ofv.conectar(ofv.strconn0)

 Strsqle3 = "Select tipoPerfil  From TipoPerfil WHERE CODpuestoE  ='" & sdato & "'"
'RESPONSE.WRITE Strsqle3
 SET rsdomeco = ofv.crearconsultaEx(StrSqle3, cnext, 1, parametros)

 if not rsdomeco.eof then
   getperfilExtRiva=rsdomeco(0)
 else
   getperfilExtRiva ="NO" 'Sdato
 end if
 CALL ofv.cerrarconsulta(rsdomeco)
 CALL ofv.cerrarconn(cnext)
END IF
end function

function getnivelExtRiva(Sdato)

' Sdato=RTRIM(Sdato & "")
 SET cnext = ofv.conectar(ofv.strconn0)

 Strsqle3 = "Select Nccod  From NivelConfiden WHERE CODEXTERNO  ='" & sdato & "'"

 SET rsdomeco = ofv.crearconsultaEx(StrSqle3, cnext, 1, parametros)

 if not rsdomeco.eof then
   getnivelExtRiva=cdbl("0" & rsdomeco(0))
 else
   getnivelExtRiva =cdbl("8") 'Sdato
 end if
 CALL ofv.cerrarconsulta(rsdomeco)
 CALL ofv.cerrarconn(cnext)

end function

function getsectorExtRiva(Sdato)

 'Sdato=mid(sdato,1,11)

 IF SDATO & "" ="" THEN
 getsectorExtRiva="0"
 ELSE
 SET cnext = ofv.conectar(ofv.strconn0)

 Strsqle3 = "Select * From eco WHERE codinter='" & sdato & "'"

 SET rsdomeco = ofv.crearconsultaEx(StrSqle3, cnext, 1, parametros)

 if not rsdomeco.eof then
   getsectorExtRiva=rsdomeco(0)
 else
  getsectorExtRiva="0" 'Sdato
 end if
 CALL ofv.cerrarconsulta(rsdomeco)
 CALL ofv.cerrarconn(cnext)
END IF
end function


function buscarGrupoAdGP(mapgrp)

  XKEYGRP = SPLIT(mapgrp,CHR(9))
  XDELIMG=""
  XName=""

  FOR XGUINX=0 TO UBOUND(XKEYGRP)
   XName= XName & XDELIMG & "  Name='" & XKEYGRP(XGUINX) & "'"
   XDELIMG= " OR "
  NEXT


  SET cnext = ofv.conectar(ofv.strconn0)

  SET connaX = server.CREATEOBJECT("ADODB.Connection")
       connaX.provider = "ADSDSOObject"
       connaX.Properties("User ID") = session("usrEmpre")
       connaX.Properties("Password") = cstr(ofv.getCript("D",session("pwdEmpre")))
       connaX.Properties("Encrypt Password") = TRUE

      connaX.OPEN
       ldapEmpreldapEmpre= replace(session("ldapEmpre"),"<","")
       ldapEmpreldapEmpre= replace(ldapEmpreldapEmpre ,">","")



    strQuery="SELECT Name FROM '" & ldapEmpreldapEmpre & "' WHERE objectCategory='group' AND  (" & XName & ")"

   ' response.write strQuery
   SET RSAX = connaX.EXECUTE(strQuery)

    session("SGU_UIGuserid")=""
    session("SGU_UIGdescri")=""
    xgusep=""

    DO WHILE NOT   rsax.eof

   '    session("SGU_UIGuserid") = session("SGU_UIGuserid") & xgusep & rsax(0)
       Strsqle3 = "Select descripcion,ID  From grupos WHERE mapgrp  ='" & rsax(0) & "'"
     SET rsdomeco = ofv.crearconsultaEx(StrSqle3, cnext, 1, parametros)

       if not rsdomeco.eof then
        session("SGU_UIGdescri") = session("SGU_UIGdescri") & xgusep & rsdomeco(0)
        session("SGU_UIGuserid") = session("SGU_UIGuserid") & xgusep & rsdomeco(1)
       else
       session("SGU_UIGdescri") = session("SGU_UIGdescri") & xgusep & "SD"
       session("SGU_UIGuserid") = session("SGU_UIGuserid") & xgusep & "SD"
      end if
     CALL ofv.cerrarconsulta(rsdomeco)


     xgusep = CHR(9)
     RSAX.MOVENEXT
   LOOP

   RSAX.CLOSE
   SET RSAX = NOTHING
   connaX.CLOSE
   SET connaX = NOTHING
   CALL ofv.cerrarconn(cnext)
end function

%>
