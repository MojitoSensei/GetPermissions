Add-Type -AssemblyName "System.Windows.Forms"

# Fonction pour demander à l'utilisateur de choisir un dossier
function Get-UserFolderPath {
    $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderDialog.Description = "Choisissez un répertoire à analyser"
    $folderDialog.ShowNewFolderButton = $false

    # Si un répertoire est sélectionné, renvoyer son chemin
    if ($folderDialog.ShowDialog() -eq "OK") {
        return $folderDialog.SelectedPath
    } else {
        Write-Host "Aucun répertoire sélectionné"
        return $null
    }
}

# Fonction pour demander à l'utilisateur de choisir un fichier de sortie avec un nom par défaut
function Get-UserFilePath {
    $fileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $fileDialog.Filter = "Fichier Excel (*.xlsx)|*.xlsx"
    $fileDialog.DefaultExt = "xlsx"
    $fileDialog.AddExtension = $true
    $fileDialog.Title = "Choisissez où enregistrer le fichier"
    $fileDialog.FileName = "Permissions.xlsx"  # Nom par défaut

    # Si un fichier est sélectionné, renvoyer son chemin
    if ($fileDialog.ShowDialog() -eq "OK") {
        return $fileDialog.FileName
    } else {
        Write-Host "Aucun fichier sélectionné"
        return $null
    }
}

# Demander à l'utilisateur de choisir le répertoire
$parentFolderPath = Get-UserFolderPath
if ($parentFolderPath) {
    # Récupérer tous les dossiers présents dans le répertoire choisi
    $foldersToAnalyze = Get-ChildItem -Path $parentFolderPath -Directory | Select-Object -ExpandProperty Name

    # Fonction pour déterminer le niveau d'accès
    function Get-AccessLevel {
        param ([string]$rights)

        switch -Regex ($rights) {
            "ReadAndExecute|Read" { "RO" }
            "Write|Modify|CreateFiles" { "RW" }
            "FullControl" { "ALL" }
            default { "X" } # X signifie pas de permission significative
        }
    }

    # Fonction pour récupérer les permissions d'un groupe/utilisateur sur un dossier
    function Get-GroupPermissions {
        param ([string]$Path)

        try {
            $access = Get-Acl -Path $Path | Select-Object -ExpandProperty Access
        } catch {
            Write-Warning "Impossible de récupérer les permissions pour $Path : $_"
            return $null
        }

        $permissionsList = @()
        foreach ($entry in $access) {
            # Vérification du niveau d'accès avec priorité sur DENY
            $rights = $entry.FileSystemRights
            $isInherited = $entry.IsInherited

            # Identifier les permissions explicites et héritées
            $level = Get-AccessLevel -rights $rights
            if ($entry.AccessControlType -eq "Deny") {
                $level = "DENY" # Priorité pour les règles DENY
            }

            $permissionsList += [PSCustomObject]@{
                UserOrGroup = $entry.IdentityReference
                AccessLevel = $level
                Inherited   = if ($isInherited) { "Oui" } else { "Non" }
            }
        }

        return $permissionsList
    }

    # Fonction pour analyser les permissions des dossiers et sous-dossiers
    function Get-FolderPermissionsRecursively {
        param ([string]$Path)

        $allPermissions = @()

        # Vérifier les permissions du dossier parent
        $parentPermissions = Get-GroupPermissions -Path $Path
        if ($parentPermissions) {
            foreach ($permission in $parentPermissions) {
                $allPermissions += [PSCustomObject]@{
                    Niveau1       = (Get-Item -Path $Path).Name
                    Niveau2       = ""
                    Niveau3       = ""
                    UserOrGroup   = $permission.UserOrGroup
                    AccessLevel   = $permission.AccessLevel
                    Inherited     = $permission.Inherited
                }
            }
        }

        # Analyser les sous-dossiers
        $subFolders = Get-ChildItem -Path $Path -Directory
        foreach ($folder in $subFolders) {
            $Niveau1 = (Get-Item -Path $Path).Name
            $Niveau2 = $folder.Name
            $Niveau3 = ""

            $permissions = Get-GroupPermissions -Path $folder.FullName
            if ($permissions) {
                foreach ($permission in $permissions) {
                    $allPermissions += [PSCustomObject]@{
                        Niveau1       = $Niveau1
                        Niveau2       = $Niveau2
                        Niveau3       = $Niveau3
                        UserOrGroup   = $permission.UserOrGroup
                        AccessLevel   = $permission.AccessLevel
                        Inherited     = $permission.Inherited
                    }
                }
            }

            # Analyser les sous-sous-dossiers
            $subSubFolders = Get-ChildItem -Path $folder.FullName -Directory
            foreach ($subFolder in $subSubFolders) {
                $Niveau3 = $subFolder.Name

                $permissions = Get-GroupPermissions -Path $subFolder.FullName
                if ($permissions) {
                    foreach ($permission in $permissions) {
                        $allPermissions += [PSCustomObject]@{
                            Niveau1       = $Niveau1
                            Niveau2       = $Niveau2
                            Niveau3       = $Niveau3
                            UserOrGroup   = $permission.UserOrGroup
                            AccessLevel   = $permission.AccessLevel
                            Inherited     = $permission.Inherited
                        }
                    }
                }
            }
        }
        return $allPermissions
    }

    # Fonction pour transformer les données en format tableau pivot pour Excel
    function Transform-ToExcelFormat {
        param ([array]$Data)

        $groupedData = $Data | Group-Object -Property UserOrGroup
        $table = @()

        $folders = $Data | Group-Object -Property Niveau1, Niveau2, Niveau3
        foreach ($folderGroup in $folders) {
            $row = [Ordered]@{
                Niveau1       = $folderGroup.Group[0].Niveau1
                Niveau2       = $folderGroup.Group[0].Niveau2
                Niveau3       = $folderGroup.Group[0].Niveau3
            }

            foreach ($group in $groupedData) {
                $permission = $folderGroup.Group | Where-Object { $_.UserOrGroup -eq $group.Name } |
                              Select-Object -ExpandProperty AccessLevel -First 1
                $row[$group.Name] = if ($permission) { $permission } else { "X" }
            }
            $table += New-Object PSObject -Property $row
        }
        return $table
    }

    # Demander à l'utilisateur où enregistrer le fichier Excel
    $outputPath = Get-UserFilePath
    if ($outputPath) {
        $allResults = @()

        foreach ($folderName in $foldersToAnalyze) {
            $folderPath = Join-Path -Path $parentFolderPath -ChildPath $folderName
            $results = Get-FolderPermissionsRecursively -Path $folderPath
            if ($results.Count -ne 0) { $allResults += $results }
        }

        if ($allResults.Count -eq 0) {
            Write-Output "Aucune permission trouvée pour les dossiers spécifiés."
        } else {
            $formattedResults = Transform-ToExcelFormat -Data $allResults

            # Exporter les données avec mise en forme
            $formattedResults | Export-Excel -Path $outputPath -AutoSize -Title "Permissions des Dossiers" -WorksheetName "Permissions" -TableName "Permissions" -Show -CellStyleSB {
                # Couleurs basées sur les permissions
                $_.AddConditionalFormatting("RO", { $_ -eq "RO" }, "#DFF2BF")    # Vert clair
                $_.AddConditionalFormatting("RW", { $_ -eq "RW" }, "#FEEFB3")    # Orange clair
                $_.AddConditionalFormatting("ALL", { $_ -eq "ALL" }, "#FFBABA")  # Rouge clair
                $_.AddConditionalFormatting("X", { $_ -eq "X" }, "#E6E6E6")      # Gris clair

                # Couleur des lignes alternées
                $_.AddAlternatingRowColor("#F9F9F9", "#FFFFFF")
            }

            # Ajouter une légende
            @(
                [PSCustomObject]@{ Type = "RO"; Description = "Lecture seule" }
                [PSCustomObject]@{ Type = "RW"; Description = "Lecture et écriture" }
                [PSCustomObject]@{ Type = "ALL"; Description = "Contrôle total" }
                [PSCustomObject]@{ Type = "X"; Description = "Pas de permission significative" }
            ) | Export-Excel -Path $outputPath -WorksheetName "Légende" -TableName "Legende" -AutoSize -Show -Append
        }
    } else {
        Write-Host "Aucun fichier sélectionné"
    }
} else {
    Write-Host "Aucun répertoire sélectionné"
}
