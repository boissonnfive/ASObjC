--
--  AppDelegate.applescript
--  Facture
--
--  Created by Bruno Boissonnet on 07/09/2014.
--  Copyright (c) 2014 BInfoService. All rights reserved.
--

script AppDelegate
	property parent : class "NSObject"
	
	
    ----------------------------------------------------------------------------
    
    --                        VARIABLES A MODIFIER
    
    ----------------------------------------------------------------------------
    
    -- Prix horaire
    property prixHoraire                     :  25.0
    -- dossier contenant le modèle excel
    property cheminFichierModeleFactureExcel :  "/Users/bruno/Documents/binfoservice/Modele_Facture_v0_5.xltx"
    -- dossier d'archivage des factures
    property cheminDossierFactures :            "Macintosh HD:Users:bruno:Documents:binfoservice:Factures:"
    -- mail
    property monAdresseCourrier :               "binfoservice@gmail.com"
    property maSignature :                      "Signature nº1"
    property monSujet :                         "Facture BInfoService"
    property contenuMessage1 :                  "Bonjour,

veuillez trouver ci-jointe la facture de l'intervention du "
    property contenuMessage2 :                      ".

Cordialement.
Bruno.

---

"

    ----------------------------------------------------------------------------
    
    --                        VARIABLES MODIFIABLES
    
    ----------------------------------------------------------------------------


    -- Format du nom du fichier facture.
    property prefixNomFichierFacture :          "Facture_00"
    property extensionFichierFacture :          ".pdf"

    -- Fichier de log.
    property fichierLog :                       POSIX file "/Users/bruno/Desktop/log.txt"
    property referenceVersFichierLog :          ""
    
    -- Dossier temporaire où va être enregistrée la facture avant d'être envoyée au client.
    -- ATTENTION ! Chemin au format HFS car Excel ne comprend pas le POSIX.
    property cheminDossierTempFacture :         "Macintosh HD:Users:bruno:Desktop:_"
    
    
    
    
    ----------------------------------------------------------------------------
    
    --                      VARIABLES A NE PAS TOUCHER
    
    ----------------------------------------------------------------------------
    
    
    -- Variables de l'interface graphique
    -- IBOutlets
    property theWindow              : missing value
    property datePicker          : missing value
    property tfNomClient         : missing value
    property tfTypeIntervention  : missing value
    property tfDureeIntervention : missing value
    property tfNumeroDeFacture   : missing value
    property tfPrixHoraire       : missing value
    
    -- Variables nécessaire au fonctionnement du programme
    -- Autres propriétés du script
    property numeroFacture :                    ""
    property cheminFichierTempFacture :         ""
    property nomFichierFacture :                ""
    property cheminFactureFinal :               ""
    property dossierFacturesAlias :             ""
    --property dateDuJourCourte :                 ""
    property adresseCourrierClient :            ""
    property nomClient :                        ""
    --property ficheClient :                      ""
    property typeIntervention :                 ""
    property dureeIntervention :                ""
    property dateIntervention :                 ""


    ----------------------------------------------------------------------------
    
    --                      applicationWillFinishLaunching
    
    ----------------------------------------------------------------------------
	
    (*
     applicationWillFinishLaunching :   Fonction appelée juste avant que la
                                        fenêtre n'apparaîsse à l'écran.
                                        Utile pour initialiser les variables.
     aNotification                  :   ???.
    *)
    
	on applicationWillFinishLaunching_(aNotification)
		-- Insert code here to initialize your application before any files are opened
        
        set referenceVersFichierLog to open for access fichierLog with write permission
        
        my logToFileAndToConsole("\n---------- Début du programme ------------\n")
        
        
        -- Initialisation des variables du programme ---------------------------
        
        my logToFileAndToConsole("----> Initialisation des variables")
        
        
        (* numeroFacture : récupéré à partir du numéro de fichier de la dernière facture *)
        
        -- set dossierFacturesAlias to (POSIX file chemindossierFactures) as alias
        set dossierFacturesAlias to cheminDossierFactures as alias
        set numeroFacture to my getNumeroFacture(dossierFacturesAlias)
        my logToFileAndToConsole("Numéro : " & numeroFacture)
        
        
        (* nomFichierFacture : nom du fichier facture de la forme "Facture_000X.pdf" *)
        
        set nomFichierFacture to prefixNomFichierFacture & (numeroFacture as text) & extensionFichierFacture
        my logToFileAndToConsole("Nom du fichier : " & nomFichierFacture)
        
        
        (* cheminFichierTempFacture : Chemin vers le dossier temporaire de stockage de la facture *)
        
        set cheminFichierTempFacture to cheminDossierTempFacture & space & nomFichierFacture
        my logToFileAndToConsole("Chemin temporaire : " & (POSIX path of cheminFichierTempFacture) as text)
        
        
        (* cheminFactureFinal : Chemin vers le dossier final de stockage de la facture *)
        
        set cheminFactureFinal to cheminDossierFactures & nomFichierFacture
        my logToFileAndToConsole("Chemin final : " & (POSIX path of cheminFactureFinal) as text)
        
        (* contenuMessage1 et contenuMessage2 : contenu du message à envoyer au client *)
        
        -- ATTENTION ! Il ne faut pas de tabulations ou d'espaces dans le texte.
        -- Attention lors de l'indentation du code.
        (*
        set contenuMessage1 to "Bonjour,

veuillez trouver ci-jointe la facture de l'intervention du "
         *)
        my logToFileAndToConsole("Contenu du message (1ére partie ) : \n" & contenuMessage1)
        (*
        set contenuMessage2 to ".

Cordialement.
Bruno.

---

"
         *)
        my logToFileAndToConsole("Contenu du message (2ème partie ) : \n" & contenuMessage2)
        
        (* dateDuJourCourte : Date du jour au format JJ/MM/AAAA *)
        
        --set dateDuJourCourte to short date string of (current date) -- format JJ/MM/AAAA
        --my logToFileAndToConsole("Date du jour (forme JJ/MM/AA) : " & dateDuJourCourte)
        
        -- Initialisation de l'interface du programme ---------------------------
        
        (* Initialise le calendrier de l'interface avec la date du jour *)
        
        --set thisDate to current date
        --my logToFileAndToConsole("Date du jour : " & (thisDate as text))
        --datePicker's setDateAS:thisDate
        --datePicker's setDateAS(current date)
        datePicker's setDateValue_(current date)
        
        tfNumeroDeFacture's setStringValue_(numeroFacture)
        
        tfPrixHoraire's setStringValue_(prixHoraire)
        
        my logToFileAndToConsole("<---- Initialisation des variables")
        
	end applicationWillFinishLaunching_
    
    
    
	
    
    ----------------------------------------------------------------------------
    
    --                      applicationShouldTerminate
    
    ----------------------------------------------------------------------------

    (*
     applicationShouldTerminate :   Fonction appelée juste avant que le programme
                                    se termine. Utile pour fermer des fichiers et
                                    tout mettre en forme une dernière fois.
     sender                     :   ???.
     *)

	on applicationShouldTerminate_(sender)
        
		-- Insert code here to do any housekeeping before your application quits
        my logToFileAndToConsole("\n----------- Fin du programme -------------\n")
        
        close access referenceVersFichierLog
        
		return current application's NSTerminateNow
        
	end applicationShouldTerminate_
    
    
    
    
    ----------------------------------------------------------------------------
    
    --                              clickButton
    
    ----------------------------------------------------------------------------
    
    
    (*
     clickButton :   Fonction appelée lorsque l'on clique sur le bouton "Facturer".
     sender      :   Le bouton "Facturer".
     *)
    
    on clickButton_(sender)
        
        my logToFileAndToConsole("\n----> Nouvelle facture")
        
        
        (***************************** Récupération des informations du client dans Contacts ********************************)
        
        set nomClient to (tfNomClient's stringValue()) as text
        my logToFileAndToConsole("Client recherché : " & nomClient)
        
        
        (* Cherche le client dans l'application « Contacts » *)
        tell application "Contacts"
            
            --tell current application to my logToFileAndToConsole("On appelle l'application \"Contacts\"")
            
            (* Récupération de la fiche client à partir du nom *)
            
            -- Liste de toutes les personnes qui portent ce nom
            set ficheclientTrouvees to every person whose name contains nomClient
            
            if (count of ficheclientTrouvees) = 0 then
                
                tell current application to my logToFileAndToConsole("Aucune personne ne porte ce nom.")
                (* Aucune personne ne porte ce nom, on déclenche une erreur	*)
                
                display dialog "Aucun contact dont le nom de famille est \"" & ¬
                nomClient & "\" n'a été trouvé." & return & ¬
                "Veuillez créer le contact et relancer le programme." & return & ¬
                "Fin du programme." with title ¬
                "Facture" buttons {"Terminer"} cancel button "Terminer" default button "Terminer"
                
            else if (count of ficheclientTrouvees) = 1 then
                
                --tell current application to my logToFileAndToConsole("Une personne porte ce nom, on récupère sa fiche.")
                (* Une personne porte ce nom, on récupère sa fiche. *)
                
                set ficheClient to first item of ficheclientTrouvees
                
            else if (count of ficheclientTrouvees) > 1 then
                
                --tell current application to my logToFileAndToConsole("Plusieurs personnes portent ce nom.")
                (*
                 Plusieurs personnes portent ce nom.
                 On met les noms complets (propriété name) dans une liste
                 et on demande à l'utilisateur de sélectionner la personne recherchée.
                 *)
                
                set nomsClients to {}
                
                repeat with unClient in ficheclientTrouvees
                    set end of nomsClients to name of unClient
                end repeat
                
                set listeReponse to choose from list nomsClients ¬
                with title ¬
                "Facture" with prompt "Il y a plus d'un contact qui porte le nom " & nomClient & return & ¬
                "Choisissez dans la liste le contact que vous recherchez." OK button name ¬
                "OK" cancel button name ¬
                "Annuler" without multiple selections allowed
                
                set nomClient to item 1 of listeReponse
                --display dialog nomClient
                
                (*
                 De la position du nom dans la liste,
                 on trouve la fiche dans la liste des clients
                 *)
                set pos to my list_position(nomClient, nomsClients)
                --display dialog pos
                
                set ficheClient to item pos of ficheclientTrouvees
                
            end if
            
            
            (* Récupération du nom complet du client *)
            
            set nomClient to name of ficheClient
            tell current application to my logToFileAndToConsole("Nom du client : " & nomClient)
            
            
            (* Récupération de la forme juridique *)
            if company of ficheClient is true then
                set formeJuridique to "Entreprise"
            else
                set formeJuridique to "Particulier"
            end if
            
            tell current application to my logToFileAndToConsole("Forme juridique du client : " & formeJuridique)
            
            
            (* Récupération de l'adresse mail à partir de la fiche client *)
            
            --repeat
                
            set contactEmail to emails of ficheClient
                
            if contactEmail ≠ {} then
                    
                    --exit repeat
                    
            else
                
                display dialog "Aucune adresse mèle renseignée pour le client \"" & nomClient & "\"." & ¬
                return & ¬
                "Veuillez compléter la fiche client et relancer le programme." & return & ¬
                "Fin du programme." with title ¬
                "Facture" buttons {"Terminer"} cancel button "Terminer" default button "Terminer"
                    
            end if
                
            --end repeat
            
            if (count of contactEmail) > 1 then
                
                set theEmails to {}
                repeat with anEmail in contactEmail
                    set end of theEmails to label of anEmail & ": " & value of anEmail
                end repeat
                set contactEmail to first item of (choose from list theEmails with prompt "Veuillez sélectionner l'adresse mail à utiliser :")
                set AppleScript's text item delimiters to ": "
                set adresseCourrierClient to text item -1 of contactEmail
                set AppleScript's text item delimiters to {""}
                
                --tell current application to my logToFileAndToConsole("Adresse mail principale du client : " & adresseCourrierClient)
                
            else if (count of contactEmail) = 1 then
            
                set adresseCourrierClient to value of first item of contactEmail
                --tell current application to my logToFileAndToConsole("Adresse mail du client : " & adresseCourrierClient)
                
            end if
            
            tell current application to my logToFileAndToConsole("Adresse mail du client : " & adresseCourrierClient)
            
            (* Récupération de l'adresse à partir de la fiche client *)
            
            try
                set adresseClient to first item of address of ficheClient
                tell adresseClient
                    set adresseClientFormatee to street & space & zip & space & city
                end tell
                
                tell current application to my logToFileAndToConsole("Adresse du client : " & adresseClientFormatee)
                
            on error
            
                display dialog "Aucune adresse renseignée pour le client \"" & nomClient & "\"." & ¬
                return & "Veuillez compléter la fiche client et relancer le programme." & return & ¬
                "Fin du programme." with title ¬
                "Facture" buttons {"Terminer"} cancel button "Terminer" default button "Terminer"
                
            end try
            
            
            quit
            
        end tell -- application "Contacts"
        
        -- display dialog nomClient
        
        -- log "Nom du client : " & nomClient
        -- log "Adresse mail du client : " & adresseCourrierClient
        -- log "Adresse du client : " & adresseClientFormatee
        
        tfNomClient's setStringValue_(nomClient)
        
        --set theText to leTexte's stringValue()
        --display alert "Vous avez écrit : " & theText
        
        --leTexte's setStringValue_("rien du tout")
        --unAutreTexte's setStringValue_("que dalle, oui !")
        
        
        (************************ Récupération des données de l'intervention ********************)
        
        
        (* Récupération du type d'intervention *)
        
        set typeIntervention to (tfTypeIntervention's stringValue()) as text
        my logToFileAndToConsole("Le type d'intervention est : " & typeIntervention)
        
        
        (* Récupération de la durée de l'intervention *)
        
        set dureeIntervention to (tfDureeIntervention's stringValue()) as text
        my logToFileAndToConsole("La durée d'intervention est : " & dureeIntervention)
        
        
        (* Récupération de la date de l'intervention *)
        set dateIntervention to (datePicker's dateAS() as date)
		my logToFileAndToConsole("La date de l'intervention est : " & dateIntervention as text)
        
        (*
         
         -- Si on est arrivé ici, c'est que l'on a toutes les données de l'intervention
         display alert "Type d'intervention : " & typeIntervention & return & ¬
         "Durée de l'intervention : " & dureeIntervention & return & ¬
         "Date de l'intervention : " & dateIntervention
         --return
         
         *)
        
        
        (***************************** Création de la facture dans Excel ***************************)
        
        
        (* Crée un fichier excel à partir d'un modèle, le modifie en fonction des données entrées et l'enregistre en pdf sur le bureau.
         Le fichier aura le nom: " Facture_00XXX.pdf" avec une espace devant ajoutée automatiquement par excel qui concatène le nom du fichier avec le nom de la feuille de calcul.*)
        tell application "Microsoft Excel"
            
            --tell current application to my logToFileAndToConsole("Appel de l'application Excel.")
            
            open cheminFichierModeleFactureExcel as POSIX file
            
            tell worksheet "Sheet1" of active workbook
                
                -- Mise à jour du numéro de facture (dans les 2 cellules!)
                set value of cell "D4" to numeroFacture
                set value of cell "B11" to numeroFacture
                
                -- Mise à jour de la date d'intervention
                set value of cell "B12" to short date string of dateIntervention
                
                -- Mise à jour du nom du client
                set value of cell "B14" to nomClient
                
                -- Mise à jour de l'adresse du client
                set value of cell "B15" to adresseClientFormatee
                
                -- Mise à jour de la forme juridique du client
                set value of cell "B16" to formeJuridique
                
                -- Mise à jour du type d'intervention
                set value of cell "A21" to typeIntervention
                
                -- Mise à jour de la durée de l'intervention
                set value of cell "C21" to dureeIntervention
                
                -- Mise à jour du coût de l'intervention
                set value of cell "D21" to dureeIntervention * prixHoraire
                
            end tell --worksheet "Sheet1" of active workbook
            
            -- Une fois la facture remplie, on l'enregistre au format PDF
            -- On donne un nom à la feuille (ex : "Facture_00" & "234")
            set name of active sheet to (prefixNomFichierFacture & (numeroFacture as text))
            -- On enregistre au format PDF en indiquant le chemin du fichier de la forme
            -- chemin + extension (ex : "Macintosh HD:Users:bruno:Desktop:" & ".pdf"
            -- le fichier produit aura la forme : chemin + une espace + nom de la feuille + extension
            -- (ex : "Macintosh HD:Users:bruno:Desktop: Facture_00234.pdf"
            -- Il faudra donc renommer le fichier pour enlever l'espace inutile
            tell active sheet to save as filename (cheminDossierTempFacture & extensionFichierFacture) file format PDF file format
            
            -- On n'enregistre pas les modifications dans le fichier modèle
            quit without saving
            
        end tell --application "Microsoft Excel"
        
        
        
        (************************** Renommage du fichier facture avec Finder ************************)
        
        (*
         On renome le fichier facture correctement (Impossible de le faire depuis Excel)
         et on le déplace dans le dossier des factures.
         *)
        tell application "Finder"
            
            --tell current application to my logToFileAndToConsole("Appel de l'application Finder.")
            
            set fichierTempFactureAlias to cheminFichierTempFacture as alias
            
            -- On renome le fichier pour enlever l'espace inutile
            set the name of fichierTempFactureAlias to nomFichierFacture
            
            set dossierFacturesAlias to cheminDossierFactures as alias
            
            try
                -- Si on déplace un fichier, il faut mettre à jour la variable qui pointe dessus
                -- car sinon on ne peut plus l'utiliser
                set myNewFile to move fichierTempFactureAlias to dossierFacturesAlias
                
                --tell current application to my logToFileAndToConsole("Copie de " & cheminFichierTempFacture & " vers " & cheminDossierFactures & ".")
                
            on error
                display alert "Impossible de copier le fichier: " & ¬
                cheminFichierTempFacture & return ¬
                & "Un fichier portant le même nom existe peut-être déjà." & return ¬
                & "Opération de déplacement annulée."
            end try
            
        end tell
        
        
        
        -- Transformation de la date dans le format : nomDuJour NuméroDuJour NomDuMois Année.
        -- Exemple : lundi 29 septembre 2015
        -- Pour utilisation dans le message à envoyer au client.
        set dateInterventionFormatLong to my getDate(dateIntervention)
        --my logToFileAndToConsole("Date de l'intervention au format long : " & dateInterventionFormatLong)
        
        
        
        (***************************** Envoi de la facture avec Mail ***************************)
        
        (*
         Crée le message contenant la facture à envoyer au client.
         Impossible d'ajouter la signature avant la pièce jointe,
         sinon elle est supprimée ou considérée comme du texte
         et non plus comme une signature (elle est même soulignée).
         
         Pour répondre à ce problème, j'ajoute la pièce jointe,
         j'attends 5 secondes et j'ajoute la signature à la fin
         (via l'interface graphique).
         *)
        tell application "Mail"
            
            activate
            
            --tell current application to my logToFileAndToConsole("Affichage de mail")
            
            set contenuMessage to contenuMessage1 & dateInterventionFormatLong & contenuMessage2
            set nouveauMessage to make new outgoing message with properties {sender:monAdresseCourrier, subject:monSujet, visible:true, content:contenuMessage}
            
            tell nouveauMessage
                
                make new recipient at end of to recipients with properties {address:adresseCourrierClient}
                
                make new attachment with properties {file name:(POSIX path of cheminFactureFinal)} -- inséré à la fin du message
                delay 5
                
            end tell
            
            activate
            
            (* Insère la signature par GUI Scripting *)
            tell application "System Events"
                tell process "Mail"
                    --delay 1.3
                    tell window 1
                        click pop up button 1
                        click menu item 3 of menu 1 of pop up button 1
                    end tell
                    
                end tell
            end tell
        end tell
        
        
        --display dialog "Fin du programme"
        -- create an alert
        --set theAlert to current application's NSAlert's makeAlert:"An alert" buttons:{"Cancel", "OK"} |text|:(theDate as text)
        --theAlert's showModal() -- returns name of the button
        --log result
        
        
        (***************************** MAJ du numéro de facture ***************************)
        
        my logToFileAndToConsole("----> MAJ des variables de la facture")
        
        (* Incrémentation du numéro de facture *)
        
        set numeroFacture to numeroFacture + 1
        my logToFileAndToConsole("Numéro : " & numeroFacture)
        tfNumeroDeFacture's setStringValue_(numeroFacture) -- MAJ de l'interface
        
        
        (* Création du nom du fichier facture *)
        
        set nomFichierFacture to prefixNomFichierFacture & (numeroFacture as text) & extensionFichierFacture
        my logToFileAndToConsole("Nom du fichier : " & nomFichierFacture)
        
        
        (* Chemin vers le dossier temporaire de stockage de la facture *)
        
        set cheminFichierTempFacture to cheminDossierTempFacture & space & nomFichierFacture
        my logToFileAndToConsole("Chemin temporaire : " & (POSIX path of cheminFichierTempFacture) as text)
        
        
        (* Chemin vers le dossier final de stockage de la facture *)
        
        set cheminFactureFinal to cheminDossierFactures & nomFichierFacture
        my logToFileAndToConsole("Chemin final : " & (POSIX path of cheminFactureFinal) as text)
        
        
        my logToFileAndToConsole("<---- MAJ des variables de la facture")
        
        my logToFileAndToConsole("<---- Nouvelle facture\n")
        
    end clickButton_
    
    

    ----------------------------------------------------------------------------
    
    --                              showAlertModal

    ----------------------------------------------------------------------------

    -- shows modal alert
	on showAlertModal:sender
		-- create an alert
		set theAlert to current application's NSAlert's makeAlert:"An alert" buttons:{"Cancel", "OK"} |text|:"Further explanation"
		theAlert's showModal() -- returns name of the button
		log result
	end showAlertModal:
    
    
    ----------------------------------------------------------------------------
    
    --                              datePickerGet
    
    ----------------------------------------------------------------------------
    
    on datePickerGet:sender
        
		set theDate to (datePicker's dateAS() as date)
		log theDate as text
		log "##"
		log "##"
        
	end datePickerGet:
    
    
    
    ----------------------------------------------------------------------------
    
    --                              list_position
    
    ----------------------------------------------------------------------------
    
    (*
     list_position : renvoie la position d'un élément dans une liste.
     this_item : élément à rechercher.
     this_list : liste dans laquelle on recherche.
     résultat : 0 si non trouvé, la position de l'élément
     dans la liste sinon.
     *)
    on list_position(this_item, this_list)
        
        local position
        set position to 0
        
        repeat with i from 1 to the count of this_list
            
            if item i of this_list is this_item then
                set position to i
            end if
            
        end repeat
        
        return position
        
    end list_position
    
    
    ----------------------------------------------------------------------------
    
    --                              getDate
    
    ----------------------------------------------------------------------------
    
    (*
     getDate :      renvoie la date au format "lundi 31 mars 2014" à partir
                    d'une date complète.
     shortDate :    date complète (format "lundi 31 mars 2014 17:19").
     résultat :     date au format "jour n° jour mois année".
     *)
    on getDate(dateComplete)
        
        --tell current application to my logToFileAndToConsole("Entrée dans la fonction \"getDate\".")
        
        tell current application to log dateComplete as text
        
        set monInstant to dateComplete as text
        set mon_jour_semaine_fr to 1st word of monInstant
        set mon_jour_fr to 2nd word of monInstant
        set mon_mois_fr to 3rd word of monInstant
        set mon_annee_fr to 4th word of monInstant
        set mon_heure_fr to 5th word of monInstant
        set ma_minute_fr to 6th word of monInstant
        set ma_seconde_fr to 7th word of monInstant
        return mon_jour_semaine_fr & " " & mon_jour_fr & " " & mon_mois_fr & " " & mon_annee_fr
        
    end getDate
    

    ----------------------------------------------------------------------------
    
    --                              getNumeroFacture
    
    ----------------------------------------------------------------------------

    (*
     getNumeroFacture : renvoie le numéro de facture à partir du nom du dernier fichier
     dans le dossier des factures (Facture_00212.pdf).
     dossierFactures :  dossier contenant les fichiers de facture.
     résultat :         le numéro de facture à utiliser pour facturer.
     *)
    on getNumeroFacture(dossierFactures)
        
        set listeNomFichiers to {}
        
        tell application "Finder"
            set tousLesFichiers to every file of folder dossierFactures
            repeat with unFichier in tousLesFichiers
                set nomFichier to the name of unFichier
                --set end of listeNomFichiers to nomFichier
            end repeat
            set numeroFacture to text 11 through 13 of nomFichier
            
            return (numeroFacture as integer) + 1
        end tell
        
        
    end getNumeroFacture
    


    ----------------------------------------------------------------------------
    
    --           applicationShouldTerminateAfterLastWindowClosed
    
    ----------------------------------------------------------------------------

    (*
     applicationShouldTerminateAfterLastWindowClosed : Fonction à redéfinir si
                                                       on veut que le programme
                                                       se termine quand on
                                                       clique sur le bouton de
                                                       fermeture de la fenêtre.
     nsApplication                                   : Application.
    *)

    on applicationShouldTerminateAfterLastWindowClosed_(nsApplication)
        
        tell nsApplication to terminate_(me)
        
    end applicationShouldTerminateAfterLastWindowClosed
    



    ----------------------------------------------------------------------------
    
    --                              logToFile
    
    ----------------------------------------------------------------------------
    
    (*
     logToFile :    écrit une ligne (en UTF-8) à la fin du fichier referenceVersFichierLog.
     ligne :        le texte à écrire dans le fichier.
     résultat :     Aucun.
     *)
    on logToFile(ligne)
        
        set ligne to ligne & "\n"
        
        try
            write ligne to referenceVersFichierLog starting at eof as «class utf8»
        on error
            close access referenceVersFichierLog
        end try
        
    end logToFile
    
    
    
    on logToFileAndToConsole(ligne)
        
        logToFile(ligne)
        log(ligne)
        
    end logToFileAndToConsole
    
	
end script
