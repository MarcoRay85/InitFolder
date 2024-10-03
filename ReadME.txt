PROCEDURA PER AGGIORNARE LE ISTALLAZIONI

1) Copiare il File INTACT.exe nella cartella D:/Intact
	1A) Copiare la cartella HTML in D:/Intact e sovrascrivere

2) Se il progetto richiede TDV:
	2A) Copiare il contenuto della cartella "TDV_compiled_V2.1.2" nella cartella D:/Tools/TestdocViewer2
	2B) Quindi ti sposti nella cartella D:/Tools/TestdocViewer2 ed esegui 2C	
	2C) Se da Kickoff NON si richiede invio dei risultati tramite stored procedure
            Aprire con Netepad++ il file "TestdocViewer.exe.config" ed impostare il parametro "isSendingEnable" a 0 => <add key="isSendingEnable" value="0" />
            (IMPOSTARLO AD 1 Viceversa)

3) Copiare il contenuto, escluse le istruzioni, di VS nella directory D:/Tools.
       3A) Seguire le istruzioni del file txt.