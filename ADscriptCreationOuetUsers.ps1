# --- Fonctions ---
function Create-OrganizationalUnit {
    param($Site, $Service, $baseDN)

    $villeDN = "OU=ou-$Site,$baseDN"
    $serviceDN = "OU=ou-$Service,$villeDN"

    # Créer l'OU du site si elle n'existe pas
    if (-not (Get-ADOrganizationalUnit -LDAPFilter "(name=ou-$Site)" -SearchBase $baseDN -ErrorAction SilentlyContinue)) {
        try {
            New-ADOrganizationalUnit -Name "ou-$Site" -Path $baseDN
            Write-Host "OU du site '$Site' créée."
        } catch {
            Write-Host "Erreur création OU site $Site : $_" -ForegroundColor Red
        }
    }

    # Créer l'OU du service si elle n'existe pas
    if (-not (Get-ADOrganizationalUnit -LDAPFilter "(name=ou-$Service)" -SearchBase $villeDN -ErrorAction SilentlyContinue)) {
        try {
            New-ADOrganizationalUnit -Name "ou-$Service" -Path $villeDN
            Write-Host "OU du service '$Service' créée."
        } catch {
            Write-Host "Erreur création OU service $Service : $_" -ForegroundColor Red
        }
    }
}

function Create-Groups {
    param($Site, $Service, $baseDN)

    $baseOU = "OU=ou-$Service,OU=ou-$Site,$baseDN"
    $globalGroupName = "GG-$Service"
    $localGroupName = "GL-$Service"

    # Créer le groupe global
    if (-not (Get-ADGroup -LDAPFilter "(name=$globalGroupName)" -SearchBase $baseOU -ErrorAction SilentlyContinue)) {
        try {
            New-ADGroup -Name $globalGroupName -GroupScope Global -Path $baseOU -Description "Groupe Global pour $Service"
            Write-Host "Groupe global '$globalGroupName' créé."
        } catch {
            Write-Host "Erreur création groupe global : $_" -ForegroundColor Red
        }
    }

    # Créer le groupe local
    if (-not (Get-ADGroup -LDAPFilter "(name=$localGroupName)" -SearchBase $baseOU -ErrorAction SilentlyContinue)) {
        try {
            New-ADGroup -Name $localGroupName -GroupScope DomainLocal -Path $baseOU -Description "Groupe Local pour $Service"
            Write-Host "Groupe local '$localGroupName' créé."
        } catch {
            Write-Host "Erreur création groupe local : $_" -ForegroundColor Red
        }
    }

    # Ajouter le groupe global dans le groupe local
    try {
        Add-ADGroupMember -Identity $localGroupName -Members $globalGroupName
        Write-Host "Ajout de $globalGroupName dans $localGroupName réussi."
    } catch {
        Write-Host "Erreur ajout de $globalGroupName dans $localGroupName : $_" -ForegroundColor Red
    }
}

function Create-ServiceFolders {
    param($Site, $Service, $basePath)

    $sitePath = Join-Path -Path $basePath -ChildPath $Site
    $servicePath = Join-Path -Path $sitePath -ChildPath $Service

    # Créer dossier site si besoin
    if (-not (Test-Path $sitePath)) {
        New-Item -ItemType Directory -Path $sitePath | Out-Null
        Write-Host "Dossier site '$sitePath' créé."
    }

    # Créer dossier service si besoin
    if (-not (Test-Path $servicePath)) {
        New-Item -ItemType Directory -Path $servicePath | Out-Null
        Write-Host "Dossier service '$servicePath' créé."
    }
}

function Create-UserFolder {
    param($Prenom, $Nom, $Site, $basePathUser)

    $login = ($Prenom.Substring(0,3) + "." + $Nom.Substring(0,3)).ToLower()
    $villePath = Join-Path -Path $basePathUser -ChildPath $Site
    $userPath = Join-Path -Path $villePath -ChildPath $login

    # Créer dossier site si besoin
    if (-not (Test-Path $villePath)) {
        New-Item -ItemType Directory -Path $villePath | Out-Null
        Write-Host "Dossier site '$villePath' créé."
    }

    # Créer dossier utilisateur si besoin
    if (-not (Test-Path $userPath)) {
        New-Item -ItemType Directory -Path $userPath | Out-Null
        Write-Host "Dossier utilisateur '$userPath' créé."
    }
}

# --- Déclarations ---
$baseDN = "OU=ou-site,DC=catkingdom,DC=local"
$basePath = "\\SRV-1\services$"
$basePathUser = "\\SRV-1\users$"
$csvPath = ".\UsersAD.csv"

# --- Traitement CSV ---
if (Test-Path $csvPath) {
    $users = Import-Csv -Path $csvPath -Delimiter ";"

    foreach ($user in $users) {
        # Vérification du contenu de la ligne
        if ($user.MATRICULE -and $user.NOM -and $user.PRENOM -and $user.ADRESSE -and $user.CP -and $user.VILLE -and $user.DATEEMBAUCHE -and $user.SITE -and $user.SERVICE) {
            $matricule = $user.MATRICULE.Trim()
            $nom = $user.NOM.Trim()
            $prenom = $user.PRENOM.Trim()
            $adresse = $user.ADRESSE.Trim()
            $cp = $user.CP.Trim()
            $ville = $user.VILLE.Trim()
            $dateembauche = $user.DATEEMBAUCHE.Trim()
            $site = $user.SITE.Trim()
            $service = $user.SERVICE.Trim()

            # Appels de fonctions
            Create-OrganizationalUnit -Site $site -Service $service -baseDN $baseDN
            Create-Groups -Site $site -Service $service -baseDN $baseDN
            Create-ServiceFolders -Site $site -Service $service -basePath $basePath
            Create-UserFolder -Prenom $prenom -Nom $nom -Site $site -basePathUser $basePathUser

            # --- Création de l'utilisateur AD ---
            $ouUserPath = "OU=ou-$service,OU=ou-$site,$baseDN"
            $samAccountName = ($prenom.Substring(0,1) + $nom).ToLower()
            $userPrincipalName = "$samAccountName@catkingdom.local"
            $password = "P@ssw0rd123!" | ConvertTo-SecureString -AsPlainText -Force

            # Vérifier si l'utilisateur existe déjà
            if (-not (Get-ADUser -Filter { SamAccountName -eq $samAccountName } -ErrorAction SilentlyContinue)) {
                try {
                    New-ADUser -Name "$prenom $nom" `
                               -GivenName $prenom `
                               -Surname $nom `
                               -SamAccountName $samAccountName `
                               -UserPrincipalName $userPrincipalName `
                               -Path $ouUserPath `
                               -AccountPassword $password `
                               -Enabled $true `
                               -ChangePasswordAtLogon $true `
                               -StreetAddress $adresse `
                               -PostalCode $cp `
                               -City $ville `
                               -Description "Embauché le $dateembauche - Service : $service"
                    
                    Write-Host "Utilisateur '$prenom $nom' créé dans $ouUserPath." -ForegroundColor Green
                } catch {
                    Write-Host "Erreur création utilisateur $prenom $nom : $_" -ForegroundColor Red
                }
            } else {
                Write-Host "Utilisateur '$prenom $nom' existe déjà." -ForegroundColor Yellow
            }

            # --- Ajout de l'utilisateur dans son Groupe Global ---
            $globalGroupName = "GG-$service"

            try {
                Add-ADGroupMember -Identity $globalGroupName -Members $samAccountName
                Write-Host "Ajout de l'utilisateur $samAccountName dans le groupe $globalGroupName." -ForegroundColor Cyan
            } catch {
                Write-Host "Erreur ajout utilisateur dans groupe $globalGroupName : $_" -ForegroundColor Red
            }
        }
        else {
            Write-Host "Ligne invalide (nombre de champs insuffisant)." -ForegroundColor Red
        }
    }
} else {
    Write-Host "Fichier CSV introuvable : $csvPath" -ForegroundColor Red
}
