<%@ Page Language="vb" %>


<html>
   <Script Language="VB" RunAt="Server">

Sub Page_Load(Sender as Object, e as EventArgs)

    Dim MyPath, MyName as string
    ' Display the names in C:\ that represent directories.

    MyPath = "C:\temp\"                      ' Set the path.
    MyName = Dir(MyPath, vbDirectory)   ' Retrieve the first entry.

    if MyName=""          ' The folder is not there & to be created
        MkDir("C:\temp\")       'Folder created
        span2.InnerHtml="A Folder (C:\ServerFolder) is created at the Page_Load"
    end if
  
End Sub

      Sub Upload_Click(Sender as Object, e as EventArgs)

         ' Display properties of the uploaded file

         FileName.InnerHtml = MyFile.PostedFile.FileName 
         FileContent.InnerHtml = MyFile.PostedFile.ContentType 
         FileSize.InnerHtml = MyFile.PostedFile.ContentLength
         UploadDetails.visible = True

' Let us recover only the file name from its fully qualified path at client 

        Dim strFileName as string
        strFileName = MyFile.PostedFile.FileName
        Dim c as string = System.IO.Path.GetFileName(strFileName) ' only the attched file name not its path

                 ' Let us Save uploaded file to server at C:\ServerFolder\

       ' response.write (c)
                 Try 
             '   response.write(session.sessionid)
	        MyFile.PostedFile.SaveAs("C:\temp\" + c)
              '  MyFile.PostedFile.SaveAs("\\server-a\data\" + c)
                catch Exp as exception
                  span1.InnerHtml = ""'Exp.description '"An Error occured. Please check the attached  file"
                       UploadDetails.visible = false
                       span2.visible=false
                 End Try

      
      End Sub
   </Script>
   <Body>
      
         <Font Color="DarkGreen" Face=Helvetica Size=5> <B>Uploading In ASP.Net - A Rich Uploading !
         </Font>
          <HR Size="2" Color=Black >
          <P>
                   
          <Form Method="Post"  EncType="Multipart/Form-Data" RunAt="Server">
         Choose Your File  To Upload : <BR>
         <Input ID="MyFile" Type="File" RunAt="Server" Size="40"> <BR>
         <BR>
         <Input Type="Submit" Value="Upload" OnServerclick="Upload_Click" RunAt="Server">
         <P>
         <Div ID="UploadDetails" Visible="False" RunAt="Server">
            File Name: <Span ID="FileName" RunAt="Server"/> <BR>
            File Content: <Span ID="FileContent" RunAt="Server"/> 

           <BR>
            File Size: <Span ID="FileSize" RunAt="Server"/>bytes
           <BR></Div>
           <Span ID="Span1" Style="Color:Red" RunAt="Server"/>
           <Span ID="Span2" Style="Color:Red" RunAt="Server"/>
         
      </Form>
<HR Size="2" Color=Black>

   </Body>
</html>

