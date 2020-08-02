<%

Dim strLinea, FileObject, Instream, intVisite, intTipo, OutStream,allMsg


strFile = Server.MapPath("msg.txt")

' creo l'oggetto
Set FileObject = Server.CreateObject("Scripting.FileSystemObject")


'apro in lettura

intTipo = 1 


Set InStream = FileObject.OpenTextFile(strFile,IntTipo,true)


'apro il file di testo e leggo tutti i messaggi riga per riga e li metto dentro una variabile chiamata allMsg separandoli dal <br>
'in modo da andare a capo dopo la fine di ogni messaggio

Do While not InStream.AtEndOfStream

	StrLinea = InStream.ReadLine()
	allMsg=allMsg&StrLinea&"<br>"
		
Loop

	



InStream.Close

Set InStream = Nothing



' libero risorse


Set FileObject = Nothing

'invio a Flash tutti i messaggi della chat
Response.Write("corpo="&allMsg)



%>

