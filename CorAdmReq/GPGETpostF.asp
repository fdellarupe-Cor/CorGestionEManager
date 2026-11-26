
<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->

<%


  GPITEM = Request.querystring("GPITEM")

 ''' RESPONSE.WRITE "GPITEM=" & GPITEM & "<BR>"


  set cn = ofv.conectar(ofv.strconn6)


    AFiles=split(session("Sfilename"),"#zzzz#")
    ATam=split(session("Stamano"),"#zzzz#")
    for z=0 to ubound(afiles)-1

       ''  RESPONSE.WRITE "ARCHIVO " & Z & " =" & afiles(z) & "<BR>"


             strsql = "UPDATE REQINFO  SET ARCHIVO = '" & afiles(z) & "' where ID = 0" & GPITEM
             call ofv.crearconsultaEx(StrSql,cn,1,parametros)

       ''  RESPONSE.WRITE "SQL " & strsql & "<BR>"

    next

    call ofv.cerrarconn(cn)

    response.REDIRECT "edrequernv.asp"



  %>