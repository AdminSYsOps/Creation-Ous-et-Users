

Option Explicit

' Déclaration des objets
Dim objFSO, objFile, objShell, line, fields
Dim baseDN, basePath, basePathUser, csvPath

' Configuration de base
baseDN = "ou=ou-site,DC=catkingdom,DC=local"
basePath = "\\SRV-1\services$\" 
basePathUser = "\\SRV-1\users$\" 
csvPath = ".\UsersAD.csv"

' Initialisation des objets
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("WScript.Shell")

' Lecture du fichier CSV
If objFSO.FileExists(csvPath) Then
    Set objFile = objFSO.OpenTextFile(csvPath, 1)
    line = objFile.ReadLine
    Do Until objFile.AtEndOfStream
        line = objFile.ReadLine
        If Trim(line) <> "" And Not Left(line, 1) = "#" Then
            fields = Split(line, ";")
            
            ' Vérification du nombre de champs
            If UBound(fields) >= 8 Then
                ' Récupération des informations utilisateur
                Dim matricule, nom, prenom, adresse, cp, ville, dateembauche, site, service
                matricule = Trim(fields(0))
                nom = Trim(fields(1))
                prenom = Trim(fields(2))
                adresse = Trim(fields(3))
                cp = Trim(fields(4))
                ville = Trim(fields(5))
                dateembauche = Trim(fields(6))
                site = Trim(fields(7))
                service = Trim(fields(8))
                
                ' Création des éléments
                
                CreateOrganizationalUnit site, service
                CreateGroups site, service
                CreateServiceFolders site, service
                CreateUserFolder prenom, nom, site
            Else
                WScript.Echo "Ligne invalide (nombre de champs insuffisant) : " & line
            End If
        End If
    Loop
    objFile.Close
Else
    WScript.Echo "Fichier CSV introuvable : " & csvPath
End If

' Fonction pour créer une Unité Organisationnelle
Sub CreateOrganizationalUnit(site, service)
    Dim villeDN, serviceDN, command

    villeDN = "OU=ou-" & site & "," & baseDN
    serviceDN = "OU=ou-" & service & "," & villeDN

    ' Créer l'OU de la ville si elle n'existe pas
    On Error Resume Next
    objShell.Run "dsadd ou """ & villeDN & """", 0, True
    If Err.Number <> 0 Then
        WScript.Echo "Erreur lors de la création de l'OU site : " & Err.Description
    End If
    
   
    ' Créer l'OU du service dans le site
    objShell.Run "dsadd ou """ & serviceDN & """", 0, True
    If Err.Number <> 0 Then
        WScript.Echo "Erreur lors de la création de l'OU service : " & Err.Description
    End If
    On Error GoTo 0
End Sub

' Fonction pour créer les groupes locaux et globaux
Sub CreateGroups(site, service)
    Dim baseOU, globalGroupDN, localGroupDN
    baseOU = "OU=ou-" & service & ",OU=ou-" & site & "," & baseDN

    globalGroupDN = "CN=GG-" & service & "," & baseOU
    localGroupDN = "CN=GL-" & service & "," & baseOU
 
    ' Créer les groupes si nécessaire
    On Error Resume Next
    objShell.Run "dsadd group """ & globalGroupDN & """ -scope g", 0, True
    If Err.Number <> 0 Then
        WScript.Echo "Erreur lors de la création du groupe global : " & Err.Description
    End If

    objShell.Run "dsadd group """ & localGroupDN & """ -scope domainlocal", 0, True  ' Correction ici
    If Err.Number <> 0 Then
        WScript.Echo "Erreur lors de la création du groupe local : " & Err.Description
    End If
    On Error GoTo 0

    ' Ajouter le groupe global dans le groupe local
    objShell.Run "dsmod group """ & localGroupDN & """ -addmbr """ & globalGroupDN & """", 0, True
    If Err.Number <> 0 Then
        WScript.Echo "Erreur lors de l'ajout du groupe global au groupe local : " & Err.Description
    End If
End Sub




' Fonction pour créer les dossiers de services
Sub CreateServiceFolders(site, service)
    Dim sitePath, servicePath

    sitePath = basePath & site
    servicePath = sitePath & "\" & service

    ' Créer le dossier site si nécessaire
    If Not objFSO.FolderExists(sitePath) Then
        objFSO.CreateFolder sitePath
    End If

    ' Créer le dossier service si nécessaire
    If Not objFSO.FolderExists(servicePath) Then
        objFSO.CreateFolder servicePath
    End If
End Sub

' Fonction pour créer le dossier utilisateur
Sub CreateUserFolder(prenom, nom, site)
    Dim login, villePath, userPath
    login = LCase(Left(prenom, 3) & "." & Left(nom, 3))
    villePath = basePathUser & site
    userPath = villePath & "\" & login

    ' Créer le dossier ville si nécessaire
    If Not objFSO.FolderExists(villePath) Then
        objFSO.CreateFolder villePath
    End If

    ' Créer le dossier utilisateur si nécessaire
    If Not objFSO.FolderExists(userPath) Then
        objFSO.CreateFolder userPath
    End If
End Sub