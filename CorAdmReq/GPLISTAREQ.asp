<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0


  CantRegMostrarx = 8
  CantRegMoverx   = cdbl(request.querystring("CantRegMover"))
  vuelta          = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)
  sMODAL          = request.querystring("MODAL")
  sVISTA          = request.querystring("VISTA")

  sMODALMAIN      = request.querystring("MODALMAIN")
  sCODTRACER      = request.querystring("CODTRACER")
  sFECHA          = request.querystring("FECHA")






  if sMODAL = "" then sMODAL = "DEF"

  if sMODAL = "SET" then

     sMODALVISTA     = request.FORM("NVISTA") & ""

     IF sMODALVISTA  = "" and SESSION("CVISTA") = "VISTATEMP" then sMODAL = "VISTATEMP"

  end if

 ''' sHTML = ofv.FormHeader(session("FormN"))

  WKVISION = ""


  DIM TABVISTA(20,10)
  TTABVISTA = 0



  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn5)

    sForm = "formlis"




  ''   WKMSTSEPITM = CHR(9)
  ''   WKMSTSEPDAT = CHR(8)

     WKMSTSEPITM = "|||"
     WKMSTSEPDAT = "###"




  if sMODAL = "SET"  then


     sMODALVISTA     = request.FORM("NVISTA") & ""

     IF sMODALVISTA  = "" THEN sMODALVISTA  = SESSION("CVISTA")

     Strsql = "Select DETVISTA1, DETVISTA2, CODVISTA From GPVISTAS WHERE CODVISTA = '" & sMODALVISTA & "' "


  ELSEif sMODAL = "DEF"  then


     Strsql = ""
     Strsql = Strsql & "Select A.DETVISTA1, A.DETVISTA2, A.CODVISTA, A.USUARIO, "
     Strsql = Strsql & " CASE A.USUARIO "
     Strsql = Strsql & "    WHEN '" & SESSION("USUARIO") & "' THEN '0' "
     Strsql = Strsql & "    ELSE '1' "
     Strsql = Strsql & " END AS PROPIA  "
     Strsql = Strsql & "From GPVISTAS A LEFT JOIN GPVISTASARC B ON (A.CODVISTA = B.CODVISTA) "
     Strsql = Strsql & "WHERE A.TIPO = 'REQ' AND ESTADO = 'S' AND "
     Strsql = Strsql & "( "
     Strsql = Strsql & "( A.USUARIO = '" & SESSION("USUARIO") & "') OR (A.ESDEFAULT = 'S' AND "
     Strsql = Strsql & "( A.ESPUBLIC = 'S' OR (A.ESPUBLIC = 'N' AND B.USUARIOSECTOR = '" & SESSION("USUARIO") & "' ) ) ) "
     Strsql = Strsql & ") "
     Strsql = Strsql & " ORDER BY  PROPIA, A.ESDEFAULT DESC, A.ESPUBLIC  "



  ELSEif sMODAL = "VISTATEMP"  then

       WKFVISTA = "VISTATEMP"

       SESSION("CVISTA") = "VISTATEMP"

       WKTXTVISTA = SESSION("VISTATEMP") & ""

  END IF


 IF SMODAL <> "WKF" THEN

  if sMODAL = "VISTATEMP"  then


  ELSE


  ' response.write strsql & "<br>"

    set rsx1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    IF RSX1.EOF THEN


    ELSE
 
       WKFVISTA = RSX1(2)

       SESSION("CVISTA") = RSX1(2)

       WKTXTVISTA = RSX1(0) & RSX1(1) & ""

    END IF

    call ofv.cerrarconsulta(RSX1)


  END IF

     WKTABLAITEM = SPLIT(WKTXTVISTA,WKMSTSEPITM) 


     FOR II=1 TO UBOUND(WKTABLAITEM) 

         WKTABLADATOS = SPLIT(WKTABLAITEM(II),WKMSTSEPDAT) 

         FOR YY=0 TO UBOUND(WKTABLADATOS) 

             IF YY > 9 THEN
                EXIT FOR
             END IF 

             TABVISTA(II,YY) = WKTABLADATOS(YY)
             TABVISTA(II,10) = YY

         NEXT

         TTABVISTA = II

     NEXT


    WKTXTWHERE = ""
    WKSEP      = ""

    WKHOY      = CLNG(DATE()) - CLNG(CDATE("01-01-1900")) 


     FOR II=1 TO TTABVISTA 

       IF TABVISTA(II,1) <> "TODOS" THEN

         WKTXTWHERE = WKTXTWHERE & WKSEP & " ( "
         WKSEP2     = ""

         WKTXTCAMPO = "A." & TABVISTA(II,0)

         FOR YY=1 TO TABVISTA(II,10) 

             WKTXTWHERE = WKTXTWHERE & WKSEP2 & " ( "

             SELECT CASE  TABVISTA(II,YY)

                    CASE "-MIUSR-"

                         WKTXTWHERE = WKTXTWHERE & WKTXTCAMPO & " = '" & SESSION("USUARIO") & "' "

                    CASE "-USRAC-"


                         WKTXTWHERE = WKTXTWHERE & WKTXTCAMPO & " = '" & SESSION("USUARIO") & "' "

                    CASE "-U7-"


                         WKTXTWHERE = WKTXTWHERE & WKTXTCAMPO & " > '" & (WKHOY - 8) & "' "

                    CASE "-U15-"

                         WKTXTWHERE = WKTXTWHERE & WKTXTCAMPO & " > '" & (WKHOY - 16) & "' "

                    CASE "-UM-"

                         WKTXTWHERE = WKTXTWHERE & WKTXTCAMPO & " > '" & (WKHOY - 31) & "' "

                    CASE "-UB-"

                         WKTXTWHERE = WKTXTWHERE & WKTXTCAMPO & " > '" & (WKHOY - 61) & "' "

                    CASE "-US-"

                         WKTXTWHERE = WKTXTWHERE & WKTXTCAMPO & " > '" & (WKHOY - 181) & "' "

                    CASE "-UA-"

                         WKTXTWHERE = WKTXTWHERE & WKTXTCAMPO & " > '" & (WKHOY - 366) & "' "

                    CASE ELSE

                         WKTXTWHERE = WKTXTWHERE & WKTXTCAMPO & " = '" & TABVISTA(II,YY) & "' "


             END SELECT 



             WKTXTWHERE = WKTXTWHERE & " )  "
             WKSEP2     = " OR "

         NEXT

         WKTXTWHERE = WKTXTWHERE & " )  "
         WKSEP      = " AND "

       END IF

     NEXT

       IF (WKTXTWHERE & "") <> "" THEN 
          WKTXTWHERE = " AND " & WKTXTWHERE

       END IF

       SESSION("ORWHERE") = WKTXTWHERE


    RESPONSE.REDIRECT "GPLISTAREQLIS.ASP?MODAL=" & SMODAL & "&VISTA=" & WKFVISTA


''  sHTML = sHTML & "<br><p style='border-style:solid;border-width:1;line-height:20px;' >&nbsp;" &  WKTXTWHERE & "&nbsp;</p>"






 ELSE



                        SEGWHERE = "   B.SEGFFINAL = '0' AND ( (B.RESPASG = '" & Session("Usuario") & "') OR ("

                         xxxx     = getInfoUser(Session("Usuario"))
                         xusrlog  = session("SGU_nombre")


                         if  xusrlog <> "NO" THEN


                             SEGWHERE = SEGWHERE & " ( (B.RESPTIPO = 'USR') AND (B.PASORSP = '" & Session("Usuario") & "') ) "
                             SEGWHERE = SEGWHERE & " OR ( (B.RESPTIPO = 'GRP') AND ( "


                             if  Session("DELLOG") = "S" then

                                 XMFTGRP = split(session("SegGrpLst"),chr(9))
                                 xcount  = ubound(XMFTGRP) + 1
                                 XMFSEP  = "" 
                                 SEGWHERE = SEGWHERE & " B.PASORSP in ("

                                 for XMFI = 0 to ubound(XMFTGRP)
                                     SEGWHERE = SEGWHERE  & XMFSEP & "'" & XMFTGRP(XMFI) & "'"
                                     XMFSEP  = "," 
                                 next

                                 SEGWHERE = SEGWHERE & ") ) )"

                            else
                                 SEGWHERE = SEGWHERE & " B.PASORSP = '" & SESSION("SEGGRP") & "') ) "
                            end if 





                             if  cstr(session("SGU_idcostcenter") & "") <> "0" then
                                 SEGWHERE = SEGWHERE & " OR ( (B.RESPTIPO = 'SEC') AND ( "
                                 SEGWHERE = SEGWHERE & " (B.PASORSP = '" & session("SGU_idcostcenter") & "'  ) "
                                 SEGWHERE = SEGWHERE & " AND (NOT B.SEGNIVELCNF < " & SESSION("SEGNIVEL") & " ) "
                                 SEGWHERE = SEGWHERE & " ) ) "
                             END IF


                         END IF

                         SEGWHERE = SEGWHERE & " ) )  "



          WKTXTWHERE = " AND " &  SEGWHERE 



       SESSION("ORWHERE") = WKTXTWHERE


    RESPONSE.REDIRECT "GPLISTAREQLIS.ASP?MODAL=" & SMODAL & "&VISTA=" & WKFVISTA

 END IF


  call ofv.cerrarconn(cn)

end if

Response.Write sHTML



%>