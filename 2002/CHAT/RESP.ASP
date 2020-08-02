<%

Dim strLinea, FileObject, Instream, intVisite, intTipo, OutStream, i, k, arrMsg(100), fine, allMsg

'stringa è la variabile che viene passata da Flash e contiene il nuovo messaggio dell'utente
'così composto : "nickname - messaggio"
stringa = Trim(Request("stringa"))


strFile = Server.MapPath("msg.txt")

' creo l'oggetto
Set FileObject = Server.CreateObject("Scripting.FileSystemObject")


'apro in lettura

intTipo = 1 ' lettura


Set InStream = FileObject.OpenTextFile(strFile,IntTipo,true)

'aggiungo la data al nuovo messaggio inviato dall'interfaccia Flash e metto questo messaggio al primo posto
'dell'array che conterrà tutti i messaggi della chat
i = 0
arrMsg(i) = FormatDateTime(Now(),3)&" - "&stringa



'con questo ciclo while leggo tutte le righe del file di testo e le metto ad una ad una nell'array. su ogni riga c'è un messaggio diverso

Do While not InStream.AtEndOfStream

	StrLinea = InStream.ReadLine()
	i=i+1
	arrMsg(i) = StrLinea
	'Response.Write(arrMsg(i))
	
Loop


'questo è un commento che può servirvi per visualizzare tutto il contenuto dell'array
'Response.Write(arrMsg(i))



Set OutStream = FileObject.CreateTextFile(strFile, true)

'verifico quanti messaggi sono presenti nella chat

if i>20 then
	fine=20
else
	fine=i
end if


k=0

'adesso leggo il contenuto dell'array dove ho memorizzato i messaggi della chat e scrivo sul file di testo su ogni riga un messaggio
'e allo stasso modo metto tutti i messaggi separandoli dal <br> in una variabile chiamata allMsg
'questa variabile verrà quindi inviata a Flash per visualizzare i messaggi della chat

Do While k<=fine 

	OutStream.WriteLine(arrMsg(k))
	allMsg=allMsg&arrMsg(k)&"<br>"
	'Response.Write(arrMsg(k)&"<br>")
	k=k+1


Loop


OutStream.Close
Set OutStream = Nothing



InStream.Close
Set InStream = Nothing



' libero risorse


Set FileObject = Nothing


'Response.Write("ok")


'invio a Flash tutti i messaggi della chat
Response.Write("corpo="&allMsg)



%>

