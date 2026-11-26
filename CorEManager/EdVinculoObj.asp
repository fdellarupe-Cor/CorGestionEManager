<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/CorCopyGen.inc" -->
<%


ok=ofv.CheckUsuario()



  nCodigo   = request.querystring("Codobjeto")
  sTipo     = request.querystring("Tipobjeto")
  sASP      = request.querystring("ASP")


  oCodigo    = 0

    Set cn = OFV.conectar(OFV.strconn2)

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''' VERIFICA EL VINCULO DEL DOCUMENTO ''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Strsql = "Select * From Objlnk Where codobjeto = " & nCodigo & " and ltipo = 'LPL'"
    set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    if not rs.eof then 
       oCodigo    = RS(6)
       oTipo      = RS(7)
    end if

    call ofv.cerrarconsulta(rs)
    Call OFV.cerrarconn(cn)


  IF oCodigo <> 0 THEN

  CCGIXORG       = "PL"
  CCGIXLNK       = "N"
  CCGIOCOD       = oCodigo
  CCGIOTIP       = oTipo

  CCGINCOD       = nCodigo
  CCGINTIP       = sTipo

  bSector        = "SI"
  bEvento        = "SI"
  bProcedimiento = "SI"



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''' SECTORES ''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    'Sectores
        If bSector = "SI" Then
'''''''''''''''''' invoca generacion bloque sectores (BLKSEC)

           xxxx = CopyGenFn("BLKSEC","NO")
        End If



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''' EVENTOS ''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        If bEvento = "SI" Then
'''''''''''''''''' invoca generacion bloque eventos (BLKEVE)

           xxxx = CopyGenFn("BLKEVE","NO")
        End If


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''' PROCEDIMIENTOS ''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

      If bProcedimiento = "SI" Then

'''''''''''''''''' invoca generacion bloque procedimientos (BLKPCD)

           xxxx = CopyGenFn("BLKPCD","NO")


      End If



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Set cn = OFV.conectar(OFV.strconn2)

           strsql = "update Objetos set codvinOBJ = 'SC' "
           strsql = strsql & " where  codobjeto = " & nCodigo
           call ofv.crearconsultaEx(StrSql,cn,1,volver)

    Call OFV.cerrarconn(cn)

 END IF



 sASP = sASP & "?" & REQUEST.QUERYSTRING

' response.write sASP

 RESPONSE.REDIRECT sASP 


%>