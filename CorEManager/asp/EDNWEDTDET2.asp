<!--#INCLUDE FILE="clsUpload.asp"-->

<%

If SESSION("FILETIPO")="L" Then

	ruta=server.mappath(request.servervariables("script_name"))
	pos=InStrRev(ruta,"\")
	ruta=Mid(ruta,1,pos-1)
	pos=InStrRev(ruta,"\")
	ruta=Mid(ruta,1,pos)

	If session("APLMDCLI")="BGAL" or session("APLMDCLI")="BC" Then ruta=session ("DirUpload")

Else
	ruta=session ("DirUpload")
End If

concod=request.querystring("concod")

nCodigo = request.querystring("Codobjeto")
sTipo   = request.querystring("Tipobjeto")
concod  = request.querystring("concod")
OPC     = request.querystring("OPC")

sss=request.querystring

'response.write concod
Set fff = server.CreateObject("scripting.filesystemobject")


xarchivo=ruta & "carpetascmd\" &  concod & ".CSE"


'For I = 1 To Request.Form("txt1").Count
'	XNPSRFL = XNPSRFL & Request.Form("txt1")(I)
'Next

 ' Definir el tamaño del bloque (CHUNK) para la lectura
    ' Un tamaño común y seguro es de 4KB (4096 bytes)
    Const CHUNK_SIZE = 4096

    Dim totalBytes      ' Total de bytes a leer
    Dim bytesLeidos     ' Bytes leídos en el loop
    Dim bytesRestantes  ' Bytes restantes por leer
    Dim bytesEnBloque   ' Tamaño del bloque actual
    Dim datosBloque     ' Matriz de bytes del bloque actual
    Dim datosCompletos  ' La matriz completa de bytes (solo si el total es manejable)

    ' 1. Inicialización
    totalBytes = Request.TotalBytes
    bytesRestantes = totalBytes
    bytesLeidos = 0

    ' Si necesitas almacenar todo el resultado, es mejor usar un ADODB.Stream
    ' para concatenar los bloques binarios, ya que la concatenación de arrays
    ' en VBScript es compleja.

    Set objStream = Server.CreateObject("ADODB.Stream")
    objStream.Type = 1 ' adTypeBinary
    objStream.Open

    ' 2. Loop de Lectura
    Do While bytesRestantes > 0

        ' 3. Determinar el tamaño del bloque actual
        If bytesRestantes >= CHUNK_SIZE Then
            bytesEnBloque = CHUNK_SIZE
        Else
            bytesEnBloque = bytesRestantes ' Último bloque
        End If

        ' 4. Leer el bloque de datos binarios
        ' ESTE ES EL PASO CLAVE DENTRO DEL LOOP
        datosBloque = Request.BinaryRead(bytesEnBloque)

        ' 5. Escribir el bloque en el Stream (para acumular los datos)
        objStream.Write datosBloque

        ' 6. Actualizar contadores
        bytesLeidos = bytesLeidos + bytesEnBloque
        bytesRestantes = bytesRestantes - bytesEnBloque

    Loop

    ' 7. Acceder a los datos completos desde el Stream
    If totalBytes > 0 Then
        ' Mover al inicio
        objStream.Position = 0

        ' Si necesitas pasarlo a un componente (ej. clsUpload)
        ' Puedes pasar el objStream a un componente si este lo soporta,
        ' o usar objStream.Read(totalBytes) para obtener la matriz de bytes completa.



        ' Si los datos son texto (requiere conversión a texto como en la respuesta anterior):
        objStream.Position = 0
         objStream.Type = 2
         objStream.Charset = "UTF-8"
        datosTexto = objStream.ReadText()

    End If

    objStream.Close
    Set objStream = Nothing



function ParseMultipartFormData(texto)

	    Dim totalBytes      : totalBytes = Request.TotalBytes
	    Dim datosBinarios
	    Dim cuerpoTexto
	    Dim boundary
	    Dim partes
	    Dim parte
	    Dim nombreCampo
	    Dim valorCampo

	    ' 1. Obtener el cuerpo de la solicitud en formato binario
	    If totalBytes = 0 Then
	        Response.Write "No se encontraron datos POST."
	        Exit function
	    End If

	    ' Lectura en bloque (idealmente, pero aquí se usa simple por simplicidad)
	    datosBinarios = texto 'BytesToText(texto, "UTF-8")

	    ' 2. Convertir el contenido binario a texto (asumiendo UTF-8)
	    ' ?? El Content-Type real debe ser parseado para el charset
	    cuerpoTexto = texto 'BytesToText(datosBinarios, "UTF-8")

	    ' 3. Extraer el delimitador (BOUNDARY) del encabezado Content-Type
	    ' Esta parte es compleja y se simula aquí:
	    Dim contentTypeHeader : contentTypeHeader = Request.ServerVariables("HTTP_CONTENT_TYPE")

	    ' Buscamos la cadena 'boundary=' y tomamos lo que sigue.
	    If InStr(contentTypeHeader, "boundary=") > 0 Then
	        boundary = Mid(contentTypeHeader, InStr(contentTypeHeader, "boundary=") + 9)
	    Else
	        Response.Write "Error: No se encontró el delimitador (boundary)."
	        Exit function
	    End If

	    ' 4. Dividir el cuerpo del formulario usando el delimitador
	    ' Se añade '--' al inicio del delimitador real para la división
	    partes = Split(cuerpoTexto, "--" & boundary)


	    ' 5. Analizar cada parte
	    For Each parte In partes

	        ' Eliminar espacios en blanco y saltos de línea al inicio y fin
	        parte = Trim(parte)

	        ' Ignorar partes vacías o el delimitador final '--'
	        If Len(parte) > 0 AND parte <> "--" Then

	            ' 6. Localizar el encabezado y el valor. Se usa el doble salto de línea (vbCrLf & vbCrLf)
	            Dim posSeparador : posSeparador = InStr(parte, vbCrLf & vbCrLf)

	            If posSeparador > 0 Then

	                ' Extraer los encabezados de la parte (donde está el nombre del campo)
	                Dim encabezados : encabezados = Left(parte, posSeparador - 1)

	                ' Extraer el valor del campo (todo lo que viene después)
	                valorCampo = Mid(parte, posSeparador + Len(vbCrLf & vbCrLf))

	                ' Limpiar el valor (quitar saltos de línea y espacios sobrantes)
	                valorCampo = Trim(valorCampo)

	                ' 7. Extraer el nombre del campo (name="") del encabezado
	                ' Buscar 'name=' y extraer el texto entre comillas
	                Dim posName : posName = InStr(encabezados, "name=")
	                If posName > 0 Then
	                    Dim resto : resto = Mid(encabezados, posName + 5)
	                   ' nombreCampo = Split(resto, """")(1) ' Tomamos la cadena dentro de las comillas

	                    ' 8. Mostrar el resultado
	                   ' Response.Write "<b>Campo:</b> " & nombreCampo & "<br>"
	                    Response.Write "<b>Valor:</b> " & valorCampo & "<hr>"
	                End If
	            End If
	        End If
	    Next
      ParseMultipartFormData=valorCampo

	End function

	' Ejecutar el Sub


Function BytesToText(bytesArray, charset)
	    Const adTypeBinary = 1
	    Const adTypeText = 2

	    Dim objStream
	    Set objStream = Server.CreateObject("ADODB.Stream")

	    objStream.Type = adTypeBinary
	    objStream.Open
	    objStream.Write bytesArray

	    objStream.Position = 0
	    objStream.Type = adTypeText
	    objStream.Charset = charset

	    BytesToText = objStream.ReadText()

	    objStream.Close
	    Set objStream = Nothing
	End Function
  xx=ParseMultipartFormData(datosTexto)
 XNPSRFL=xx
Set MyFile = fff.CreateTextFile(xarchivo)
MyFile.Write XNPSRFL
MyFile.Close


If SESSION("FILETIPO") = "R" Then

	response.redirect "../../../coremanager/EDNWEDTDET.asp?" & sss
Else

	response.redirect "../EDNWEDTDET.asp?" & sss
End If


%>