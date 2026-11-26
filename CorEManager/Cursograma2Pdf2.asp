<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!--#INCLUDE FILE="asp/clsUpload.asp"-->

<%

ok=ofv.CheckUsuario()

    Set cn = OFV.conectar(OFV.strconn2)


Dim Upload
Dim FileName1
Dim FileName2
Dim Folder


response.buffer=true

ncodigo=request.querystring("codigo")
stipo=request.querystring("tipo")

Set Upload = New clsUpload
upload.Sbase64="SI"


UPLOAD.Sbase64NameArch= ncodigo &stipo & ".PDF"


if session("APLMDCLI")="BANELCO" THEN
 FOLDER=  Request.ServerVariables("PATH_TRANSLATED")
 FOLDER=left(FOLDER,instrRev(FOlDER,"\",-1))
 FOLDER=FOLDER  & "PnlPrint\"

ELSE
folder= session ("DirUpload") & "PnlPrint\"
END IF


 'response.write shtml
 'response.write "<center><hr><h2>Transferencia Ok.</h2><hr><br></center>"
'response.write folder
For i = 0 To Upload.Count - 1

		'FileName = Upload.Fields(Upload(i).Name).FileName
		upload(Upload(i).Name).saveas folder &  UPLOAD.Sbase64NameArch
Next


'''response.write "<a href=PnlPrint/" &UPLOAD.Sbase64NameArch & ">Link al documento..  " & UPLOAD.Sbase64NameArch & "</a>"


  StrSqlx = "select * FROM OBJPROCURSOCMP where codobjeto = " & ncodigo
'response.write StrSqlx
  Set rs2 = OFV.crearconsultaEx(StrSqlx, cn, 1, "")

  IF  NOT RS2.EOF THEN

      Strsql = "DELETE FROM OBJPROCURSOCMP where codobjeto = " & ncodigo
      call ofv.crearconsultaEx(StrSql,cn,1,volver)

  END IF

      Strsql = "insert into OBJPROCURSOCMP  (codobjeto,tipobjeto,TIPOCMP,NOMCMP,ACTIVADO) values ("
      Strsql = Strsql & "'" & ncodigo & "','" & stipo & "','PDF','" & UPLOAD.Sbase64NameArch & "','S')"
      call ofv.crearconsultaEx(StrSql,cn,1,volver)


   response.redirect "CursogramaAdmin.asp?Codproced=" & nID & "&Codobjeto=" & ncodigo & "&Tipobjeto=" & sTipo

Set Upload = Nothing



%>
