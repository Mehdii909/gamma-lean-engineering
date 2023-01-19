<%
Function EcrireDansFichier(Fichier, Contenu, Ajouter)

On Error Resume Next

If Ajouter = True Then
   LeMode = 8
Else 
   LeMode = 2
End If

  Set Fs = Server.Createobject("Scripting.FileSystemObject")
  Set LeFichierTexte = Fs.OpenTextFile(Fichier, LeMode, True)

  LeFichierTexte.Write Contenu

  LeFichierTexte.Close
  Set LeFichierTexte = Nothing
  Set Fs = Nothing

End Function
%>
<%
    'Pour creer/ecraser un fichier
    Call EcrireDansFichier(Monfichier,"Ceci est mon texte", True)
    'Pour ajouter à la suite d'un fichier
    Call EcrireDansFichier(Monfichier,vbCrLf & "Ceci est la 2ème ligne de mon texte", False)

%>
