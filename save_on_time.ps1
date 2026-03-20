Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Création de la fenêtre principale
$form = New-Object System.Windows.Forms.Form
$form.Text = "Save on Time"
$form.Size = New-Object System.Drawing.Size(500, 780)
$form.StartPosition = "CenterScreen"

# Ensuite, je fais en sorte que l'on ne puisse pas modifier la taille de l'application : 
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle 
$form.MaximizeBox = $false   # Désactive le bouton "agrandir"
$form.MinimizeBox = $true    # On garde le bouton "réduire"

# Maintenant, je vais définir le logo du logiciel (chemin relatif à l'exécutable, dans le même répertoire)
$cheminIcone = Join-Path $PSScriptRoot "logo_Save_on_Time.ico"
if (Test-Path $cheminIcone) {
    $form.Icon = New-Object System.Drawing.Icon($cheminIcone)
}

# Bouton permettant de choisir le dossier en entrée 
$boutonBrowseEntree            = New-Object System.Windows.Forms.Button
$boutonBrowseEntree.Text       = 'Choisir le repertoire à copier'
$boutonBrowseEntree.Size       = New-Object System.Drawing.Size(150,50)
$boutonBrowseEntree.Location   = New-Object System.Drawing.Point(20,20)

# Zone de texte pour afficher le chemin choisi du dossier en entrée
$txtSourceEntree               = New-Object System.Windows.Forms.TextBox
$txtSourceEntree.ReadOnly      = $true
$txtSourceEntree.Width         = 300
$txtSourceEntree.Location      = New-Object System.Drawing.Point(180,23)



# Boîte de dialogue fichier (OpenFileDialog)
$FolderBrowserDialogEntree = New-Object System.Windows.Forms.FolderBrowserDialog

# Événement du bouton
$boutonBrowseEntree.Add_Click({
    if ($FolderBrowserDialogEntree.ShowDialog() -eq 'OK') {                # Affiche la boîte et teste OK
        $txtSourceEntree.Text = $FolderBrowserDialogEntree.SelectedPath              # Chemin complet sélectionné
    }
})

# Bouton permettant de choisir le dossier en sortie
$boutonBrowseSortie            = New-Object System.Windows.Forms.Button
$boutonBrowseSortie.Text       = 'Choisir le repertoire de destination'
$boutonBrowseSortie.Size       = New-Object System.Drawing.Size(150,50)
$boutonBrowseSortie.Location   = New-Object System.Drawing.Point(20,120)

# Zone de texte pour afficher le chemin choisi du dossier en sortie
$txtSourceSortie               = New-Object System.Windows.Forms.TextBox
$txtSourceSortie.ReadOnly      = $true
$txtSourceSortie.Width         = 300
$txtSourceSortie.Location      = New-Object System.Drawing.Point(180,123)


# Boîte de dialogue fichier (OpenFileDialog)
$FolderBrowserDialogSortie = New-Object System.Windows.Forms.FolderBrowserDialog

# Événement du bouton
$boutonBrowseSortie.Add_Click({
    if ($FolderBrowserDialogSortie.ShowDialog() -eq 'OK') {                # Affiche la boîte et teste OK
        $txtSourceSortie.Text = $FolderBrowserDialogSortie.SelectedPath              # Chemin complet sélectionné
    }
})

# Maintenant, je vais séparer les sélections de la partie Copie du logiciel :
$separateurSelectionCopie = New-Object System.Windows.Forms.Label
$separateurSelectionCopie.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
$separateurSelectionCopie.AutoSize = $false
$separateurSelectionCopie.Width = 500
$separateurSelectionCopie.Height = 2
$separateurSelectionCopie.Location = New-Object System.Drawing.Point(0, 185)
$separateurSelectionCopie.BackColor = [System.Drawing.Color]::LightGray

# Création de la barre de progression
$ProgressBar = New-Object System.Windows.Forms.ProgressBar
$ProgressBar.Location = New-Object System.Drawing.Point(30, 250)
$ProgressBar.Size = New-Object System.Drawing.Size(415, 25)
$ProgressBar.Minimum = 0
$ProgressBar.Value = 0
$ProgressBar.Style = "Continuous"

# Label pour afficher le statut
$StatusLabel = New-Object System.Windows.Forms.Label
$StatusLabel.Location = New-Object System.Drawing.Point(50, 280)
$StatusLabel.Size = New-Object System.Drawing.Size(450, 20)
$StatusLabel.Text = "Prêt à copier..."

# Je vréer la variable copieEnCours indiquant si une copie est en cours afin de prévenir la fermeture de l'app
$script:copieEnCours = $false

# Création d'un bouton pour valider la copie
$boutonfin                 = New-Object System.Windows.Forms.Button
$boutonfin.Text            = "Effectuer la copie"
$boutonfin.Location        = New-Object System.Drawing.Point(320, 200)
$boutonfin.Size            = New-Object System.Drawing.Size(140,20)
$boutonfin.Add_Click({
                        # On commence par vérifier si les deux répertoires ont été spécifiés et sinon on en informe l'utilisateur
                        if($($txtSourceEntree.Text -eq '') -or $($txtSourceSortie.Text -eq '')) {
                            [System.Windows.Forms.MessageBox]::Show("Veuillez sélectionner en répertoire d'origine et un répertoire de destination", "Erreur", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                        }
                        # Ensuite on vérifie si les deux répertoires sont différents
                        elseif($($txtSourceEntree.Text -eq $txtSourceSortie.Text)) {
                            [System.Windows.Forms.MessageBox]::Show("Veuillez sélectionner en répertoire d'origine différent du répertoire de destination", "Erreur", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                        }
                        # Enfin si toutes les conditions sont respectées, on lance la copie
                        else {
                            # Tant que la copie est en cours, tous les contrôles interactifs sont désactivés
                            $script:copieEnCours = $true
                            $boutonfin.Enabled = $false
                            $boutonTache.Enabled = $false
                            $boutonBrowseEntree.Enabled = $false
                            $boutonBrowseSortie.Enabled = $false
                            $boutonRafraichir.Enabled = $false
                            $boutonSupprimer.Enabled = $false
                            $groupeBoutonsCheckBox.Enabled = $false
                            $selectionneurHeure.Enabled = $false
                            # Compter le nombre total de fichiers
                            $Fichiers = Get-ChildItem -Path "$($txtSourceEntree.Text)\*" -Recurse -File
                            $FichiersTotal = $Fichiers.Count

                            # Ensuite on prépare la barre de chargement
                            $ProgressBar.Maximum = $FichiersTotal
                            $ProgressBar.Value = 0

                            # Maintenant, on modifie le label de la barra afin qu'elle affiche l'état du téléversement
                            $StatusLabel.Text = "Copie en cours... 0/$FichiersTotal fichiers"
                            $form.refresh()

                            # Copier les fichiers un par un avec mise à jour de la progression
                            $NombreFichierCopies = 0
                            $FichiersEnErreur = @()

                            foreach ($File in $Fichiers) {
                                try {
                                    # D'abord, on récupère la partie relative du nom du fichier (c'est à dire la partie après le répertoire d'origine sélectionné)
                                    $CheminRelatif = $File.FullName.Substring($txtSourceEntree.Text.Length + 1)

                                    # On rajoute à ce chemin relatif le chemin absolu de notre répertoire de destination pour faire le chemin absolu du fichier de destination
                                    $FichierDestinationCheminAbsolu = Join-Path $txtSourceSortie.Text $CheminRelatif

                                    # Enfin on récupère le dossier parent pour voir s'il existe et le créer sinon
                                    $RepertoireDestinationCheminAbsolu = Split-Path $FichierDestinationCheminAbsolu -Parent

                                    if (-not (Test-Path $RepertoireDestinationCheminAbsolu)) {
                                        New-Item -ItemType Directory -Path $RepertoireDestinationCheminAbsolu -Force | Out-Null
                                    }

                                    # Copier le fichier
                                    Copy-Item -Path $File.FullName -Destination $FichierDestinationCheminAbsolu -Force -ErrorAction Stop
                                } catch {
                                    # On note le fichier en erreur et on continue la copie des autres
                                    $FichiersEnErreur += $File.Name
                                }

                                # Mettre à jour la progression
                                $NombreFichierCopies++
                                $ProgressBar.Value = $NombreFichierCopies
                                $StatusLabel.Text = "Copie en cours... $NombreFichierCopies/$FichiersTotal fichiers - Fichier actuel: $($File.Name)"

                                # Rafraîchir uniquement les contrôles visuels (sans traiter les événements utilisateur)
                                $ProgressBar.Refresh()
                                $StatusLabel.Refresh()
                            }

                            # Bilan de la copie
                            if ($FichiersEnErreur.Count -gt 0) {
                                $StatusLabel.Text = "Copie terminée avec $($FichiersEnErreur.Count) erreur(s) sur $FichiersTotal fichiers"
                                [System.Windows.Forms.MessageBox]::Show(
                                    "Copie terminée avec $($FichiersEnErreur.Count) fichier(s) en erreur :`n`n$($FichiersEnErreur -join "`n")",
                                    "Copie partielle",
                                    [System.Windows.Forms.MessageBoxButtons]::OK,
                                    [System.Windows.Forms.MessageBoxIcon]::Warning
                                )
                            } else {
                                $StatusLabel.Text = "Copie terminée ! $FichiersTotal fichiers copiés"
                                [System.Windows.Forms.MessageBox]::Show("Copie terminée avec succès", "Succès")
                            }
                            $form.refresh()
                            # Réactivation de tous les contrôles
                            $script:copieEnCours = $false
                            $boutonfin.Enabled = $true
                            $boutonTache.Enabled = $true
                            $boutonBrowseEntree.Enabled = $true
                            $boutonBrowseSortie.Enabled = $true
                            $boutonRafraichir.Enabled = $true
                            $boutonSupprimer.Enabled = $true
                            $groupeBoutonsCheckBox.Enabled = $true
                            $selectionneurHeure.Enabled = $true
                        }
})


# Maintenant, je créé le titre associé à la partie copie directe :

$titreCopie = New-Object System.Windows.Forms.Label
$titreCopie.Text = "Copie Directe"
$titreCopie.Font = New-Object System.Drawing.Font("Segoe UI", 13, ([System.Drawing.FontStyle]::Bold -bor [System.Drawing.FontStyle]::Underline))
$titreCopie.ForeColor = [System.Drawing.Color]::DarkSlateGray
$titreCopie.Location = New-Object System.Drawing.Point(30, 200)
$titreCopie.AutoSize = $true

# Je continue en faisant un séparateur entre ces deux parties : 

$separateurCopieTache = New-Object System.Windows.Forms.Label
$separateurCopieTache.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
$separateurCopieTache.AutoSize = $false
$separateurCopieTache.Width = 500
$separateurCopieTache.Height = 2
$separateurCopieTache.Location = New-Object System.Drawing.Point(0, 305)
$separateurCopieTache.BackColor = [System.Drawing.Color]::LightGray



# Et je fais ensuite le titre associé à la partie Tâche automatique

$titreTache = New-Object System.Windows.Forms.Label
$titreTache.Text = "Créer une tâche planifiée"
$titreTache.Font = New-Object System.Drawing.Font("Segoe UI", 13, ([System.Drawing.FontStyle]::Bold -bor [System.Drawing.FontStyle]::Underline))
$titreTache.ForeColor = [System.Drawing.Color]::DarkSlateGray
$titreTache.Location = New-Object System.Drawing.Point(30, 318)
$titreTache.AutoSize = $true

# Ici, nous mettons un petit champ permettant de sélectionner l'heure à laquelle la tâche planifiée sera effectuée 
$selectionneurHeure = New-Object System.Windows.Forms.DateTimePicker

# On définit le format comme Heure:Minutes car ce sont les deux seuls paramètres importants
$selectionneurHeure.Format = [System.Windows.Forms.DateTimePickerFormat]::Custom
$selectionneurHeure.CustomFormat = "HH:mm"

# Cela active l'apparition des flèches pour pouvoir choisir plus facilement
$selectionneurHeure.ShowUpDown = $true

# On indique où se trouve le champ sur l'app et on lui donne une longueur et une taille d'écriture
$selectionneurHeure.Location = New-Object System.Drawing.Point(145,467)
$selectionneurHeure.Width = 100
$selectionneurHeure.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9)

# Enfin on initie sa valeur par défaut à 12h30 qui est l'heure recommandée pour des tâches comme cela
$selectionneurHeure.Value = (Get-Date).Date.AddHours(12).AddMinutes(30)

# Je vais maintenant faire un label afin d'indiquer à quoi correspond l'heure
$labelHeure = New-Object System.Windows.Forms.Label
$labelHeure.Text = "Heure d'exécution :"
$labelHeure.Location = New-Object System.Drawing.Point(30, 470) # Ajuste l'alignement selon besoin
$labelHeure.AutoSize = $true
$labelHeure.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", 9.5)

# Je commence par initialiser l'ensemble de boutons radio
$groupeBoutonsCheckBox                 = New-Object System.Windows.Forms.GroupBox
$groupeBoutonsCheckBox.Text            = "Sélectionner les jours durant lesquels la tâche automatique s'effectuera"
$groupeBoutonsCheckBox.Location        = New-Object System.Drawing.Point(30,350)
$groupeBoutonsCheckBox.Size            = New-Object System.Drawing.Size(440,90)



$radioLundi               = New-Object System.Windows.Forms.CheckBox
$radioLundi.Location      = New-Object System.Drawing.Point(10,20)
$radioLundi.Text          = "Lundi"
$groupeBoutonsCheckBox.Controls.Add($radioLundi)


$radioMardi               = New-Object System.Windows.Forms.CheckBox
$radioMardi.Location      = New-Object System.Drawing.Point(120,20)
$radioMardi.Text          = "Mardi"
$groupeBoutonsCheckBox.Controls.Add($radioMardi)

$radioMercredi            = New-Object System.Windows.Forms.CheckBox
$radioMercredi.Location   = New-Object System.Drawing.Point(230,20)
$radioMercredi.Text       = "Mercredi"
$groupeBoutonsCheckBox.Controls.Add($radioMercredi)

$radioJeudi               = New-Object System.Windows.Forms.CheckBox
$radioJeudi.Location      = New-Object System.Drawing.Point(340,20)
$radioJeudi.Text          = "Jeudi"
$groupeBoutonsCheckBox.Controls.Add($radioJeudi)

$radioVendredi            = New-Object System.Windows.Forms.CheckBox
$radioVendredi.Location   = New-Object System.Drawing.Point(10,50)
$radioVendredi.Text       = "Vendredi"
$groupeBoutonsCheckBox.Controls.Add($radioVendredi)

$radioSamedi              = New-Object System.Windows.Forms.CheckBox
$radioSamedi.Location     = New-Object System.Drawing.Point(120,50)
$radioSamedi.Text         = "Samedi"
$groupeBoutonsCheckBox.Controls.Add($radioSamedi)

$radioDimanche            = New-Object System.Windows.Forms.CheckBox
$radioDimanche.Location   = New-Object System.Drawing.Point(230,50)
$radioDimanche.Text       = "Dimanche"
$groupeBoutonsCheckBox.Controls.Add($radioDimanche)

# Enfin, je créé une table de hash (hashtable) permettant de traduire les noms français des Checkbox avec les noms anglais reconnus par Windows
$conversionJours = @{
    "Lundi" = "Monday"
    "Mardi" = "Tuesday" 
    "Mercredi" = "Wednesday"
    "Jeudi" = "Thursday"
    "Vendredi" = "Friday"
    "Samedi" = "Saturday"
    "Dimanche" = "Sunday"
}


# Je créé ici le bouton pour valider la tâche
$boutonTache              = New-Object System.Windows.Forms.Button
$boutonTache.Text         = "Ajouter la tâche de copie planifiée"
$boutonTache.Location     = New-Object System.Drawing.Point(320,450)
$boutonTache.Size         = New-Object System.Drawing.Size(140,40)

# Ici on fait la fonction on click qui permet d'ajouter la tâche planifiée
$boutonTache.Add_Click({
                        # On commence par vérifier si les deux répertoires ont été spécifiés et sinon on en informe l'utilisateur
                        if($($txtSourceEntree.Text -eq '') -or $($txtSourceSortie.Text -eq '')) {
                            [System.Windows.Forms.MessageBox]::Show("Veuillez sélectionner en répertoire d'origine et un répertoire de destination", "Erreur", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                        }
                        # Ensuite on vérifie si les deux répertoires sont différents
                        elseif($($txtSourceEntree.Text -eq $txtSourceSortie.Text)) {
                            [System.Windows.Forms.MessageBox]::Show("Veuillez sélectionner en répertoire d'origine différent du répertoire de destination", "Erreur", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                        }
                        # Enfin si toutes les conditions sont respectées, on lance la mise en place de la tâche
                        else {
                            # D'abord, on définis l'action a effectué : ici c'est la copie des fichiers 
                            $commandePlanifiee = '"'+"Copy-Item -Path '$($txtSourceEntree.Text)/*' -Destination '$($txtSourceSortie.Text)/*' -Recurse -Force"+'"'
                            $copyAction = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument "-Command $commandePlanifiee"

                            # --- Ensuite, on définis le Trigger çad quand l'action se produit-elle. Ici c'est l'heure selectionnée par l'user.
                            $heureMinute = $selectionneurHeure.Value.ToString("HH:mm")

                            # --- De plus, on choisit les jours sélectionnés par l'utilisateur : 
                            # D'abord on créé un tableau qui contiendra les jours sélectionnés :
                            $joursSelectionnesAnglais = @()
                            foreach ($controle in $groupeBoutonsCheckBox.Controls) {
                                # Vérifier si la checkbox est cochée
                                if ($controle.Checked) {
                                    # Ajoute une virgule avant les suivants
                                    $joursSelectionnesAnglais += $conversionJours[$controle.Text]
                                }
                            }

                            # --- Si aucun jour n'est sélectionné, alors on affiche une erreur : 

                            if ($joursSelectionnesAnglais.Count -eq 0) {
                                [System.Windows.Forms.MessageBox]::Show("Veuillez sélectionner au moins un jour pour la tâche planifiée.", "Erreur", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                                return
                            }

                            # --- Et enfin, on tr
                            $Trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek $joursSelectionnesAnglais -At $heureMinute

                            # Là c'est les options : On indique donc que l'action se produit qu'on soit branché ou non et dès que possible.
                            $Settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable

                            #Enfin, on créé la tâche, en indiquant dans son nom sa date de création pour éviter les doublons
                            $dateActuelle = Get-Date
                            $dateActuelle = $dateActuelle.ToString('dd-MM-yyyy_HH-mm')
                            Register-ScheduledTask -Action $copyAction -Trigger $Trigger -Settings $Settings -TaskName "Tache_de_Copie_du_$dateActuelle" -Description "Copie Journaliere de $($txtSourceEntree.Text) vers $($txtSourceSortie.Text) à $heureMinute" -User "NT AUTHORITY\SYSTEM" -RunLevel Highest
                            $nomTache = "Tache_de_Copie_du_$dateActuelle"
                            [System.Windows.Forms.MessageBox]::Show(
                                "La tâche planifiée '$nomTache' a été ajoutée avec succès et s'exécutera à $heureMinute les jours : $($joursSelectionnesAnglais -join ', ').",
                                "Succès",
                                [System.Windows.Forms.MessageBoxButtons]::OK,
                                [System.Windows.Forms.MessageBoxIcon]::Information
                            )
                            # On rafraîchit la liste pour que la nouvelle tâche apparaisse immédiatement
                            Rafraichir-ListeTaches
                        }
})


# =============================================
# SECTION : Gestion des tâches planifiées existantes
# =============================================

# Séparateur entre la création et la gestion des tâches
$separateurTacheGestion = New-Object System.Windows.Forms.Label
$separateurTacheGestion.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
$separateurTacheGestion.AutoSize = $false
$separateurTacheGestion.Width = 500
$separateurTacheGestion.Height = 2
$separateurTacheGestion.Location = New-Object System.Drawing.Point(0, 505)
$separateurTacheGestion.BackColor = [System.Drawing.Color]::LightGray

# Titre de la section
$titreGestion = New-Object System.Windows.Forms.Label
$titreGestion.Text = "Gérer les tâches planifiées"
$titreGestion.Font = New-Object System.Drawing.Font("Segoe UI", 13, ([System.Drawing.FontStyle]::Bold -bor [System.Drawing.FontStyle]::Underline))
$titreGestion.ForeColor = [System.Drawing.Color]::DarkSlateGray
$titreGestion.Location = New-Object System.Drawing.Point(30, 518)
$titreGestion.AutoSize = $true

# ListView pour afficher les tâches existantes (uniquement celles créées par Save on Time)
$listeTaches = New-Object System.Windows.Forms.ListView
$listeTaches.Location = New-Object System.Drawing.Point(30, 550)
$listeTaches.Size = New-Object System.Drawing.Size(430, 130)
$listeTaches.View = [System.Windows.Forms.View]::Details
$listeTaches.FullRowSelect = $true
$listeTaches.GridLines = $true
$listeTaches.Columns.Add("Nom de la tâche", 200) | Out-Null
$listeTaches.Columns.Add("Prochaine exécution", 110) | Out-Null
$listeTaches.Columns.Add("Statut", 100) | Out-Null

# Fonction pour rafraîchir la liste des tâches créées par l'application
function Rafraichir-ListeTaches {
    $listeTaches.Items.Clear()
    try {
        $taches = Get-ScheduledTask | Where-Object { $_.TaskName -like "Tache_de_Copie_*" }
        foreach ($tache in $taches) {
            $item = New-Object System.Windows.Forms.ListViewItem($tache.TaskName)
            # Récupérer les infos de la prochaine exécution
            $infos = Get-ScheduledTaskInfo -TaskName $tache.TaskName -ErrorAction SilentlyContinue
            if ($infos -and $infos.NextRunTime) {
                $item.SubItems.Add($infos.NextRunTime.ToString("dd/MM/yyyy HH:mm")) | Out-Null
            } else {
                $item.SubItems.Add("N/A") | Out-Null
            }
            # Statut de la tâche
            $item.SubItems.Add($tache.State.ToString()) | Out-Null
            $listeTaches.Items.Add($item) | Out-Null
        }
    } catch {
        # Si on n'a pas les droits pour lire les tâches, on affiche un message dans la liste
        $item = New-Object System.Windows.Forms.ListViewItem("Erreur de lecture des tâches")
        $item.SubItems.Add("") | Out-Null
        $item.SubItems.Add("") | Out-Null
        $listeTaches.Items.Add($item) | Out-Null
    }
}

# Bouton pour rafraîchir la liste
$boutonRafraichir = New-Object System.Windows.Forms.Button
$boutonRafraichir.Text = "Rafraîchir"
$boutonRafraichir.Location = New-Object System.Drawing.Point(30, 690)
$boutonRafraichir.Size = New-Object System.Drawing.Size(100, 30)
$boutonRafraichir.Add_Click({ Rafraichir-ListeTaches })

# Bouton pour supprimer la tâche sélectionnée
$boutonSupprimer = New-Object System.Windows.Forms.Button
$boutonSupprimer.Text = "Supprimer la tâche"
$boutonSupprimer.Location = New-Object System.Drawing.Point(320, 690)
$boutonSupprimer.Size = New-Object System.Drawing.Size(140, 30)
$boutonSupprimer.Add_Click({
    if ($listeTaches.SelectedItems.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "Veuillez sélectionner une tâche dans la liste.",
            "Aucune sélection",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }

    $nomTacheSelectionnee = $listeTaches.SelectedItems[0].Text

    # Demander confirmation avant suppression
    $confirmation = [System.Windows.Forms.MessageBox]::Show(
        "Voulez-vous vraiment supprimer la tâche '$nomTacheSelectionnee' ?",
        "Confirmation de suppression",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )

    if ($confirmation -eq [System.Windows.Forms.DialogResult]::Yes) {
        try {
            Unregister-ScheduledTask -TaskName $nomTacheSelectionnee -Confirm:$false
            [System.Windows.Forms.MessageBox]::Show(
                "La tâche '$nomTacheSelectionnee' a été supprimée.",
                "Succès",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
            Rafraichir-ListeTaches
        } catch {
            [System.Windows.Forms.MessageBox]::Show(
                "Erreur lors de la suppression : $($_.Exception.Message)",
                "Erreur",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
        }
    }
})

# Charger la liste au démarrage de l'application
Rafraichir-ListeTaches

$form.Controls.Add($boutonBrowseEntree)

$form.Controls.Add($txtSourceEntree)

$form.Controls.Add($boutonBrowseSortie)

$form.Controls.Add($txtSourceSortie)

$form.Controls.Add($separateurSelectionCopie)

$form.Controls.Add($ProgressBar)

$form.Controls.Add($StatusLabel)

$form.Controls.Add($titreCopie)

$form.Controls.Add($separateurCopieTache)

$form.Controls.Add($titreTache)

$form.Controls.Add($selectionneurHeure)

$form.Controls.Add($labelHeure)

$form.Controls.Add($boutonfin)

$form.Controls.Add($boutonTache)

$form.Controls.Add($groupeBoutonsCheckBox)

$form.Controls.Add($separateurTacheGestion)
$form.Controls.Add($titreGestion)
$form.Controls.Add($listeTaches)
$form.Controls.Add($boutonRafraichir)
$form.Controls.Add($boutonSupprimer)

# Maintenant, je vais modifier comment se comporte le logiciel en fonction de quand on veut le fermer : 
$form.Add_FormClosing({
    param($sender, $e)
    
    if ($script:copieEnCours) {
        # Annuler la fermeture
        $e.Cancel = $true
        
        # Afficher le message d'avertissement
        [System.Windows.Forms.MessageBox]::Show(
            "Une copie est actuellement en cours. Veuillez attendre la fin de l'opération avant de fermer l'application.",
            "Opération en cours",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
    } else {
        # Fermeture normale
        $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
    }
})


$form.Show()
[System.Windows.Forms.Application]::Run($form)