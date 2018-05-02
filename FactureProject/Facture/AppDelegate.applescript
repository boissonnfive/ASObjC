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
    property cheminDossierFactures           :  "Macintosh HD:Users:bruno:Documents:binfoservice:Factures:"
    -- mail
    property monAdresseCourrier              : "binfoservice@gmail.com"
    property maSignature                     : "Signature nº1"
    property monSujet                        : "Facture BInfoService"
    property contenuMessage1                 : "Bonjour,\n\nveuillez trouver ci-jointe la facture de l'intervention du "
    property contenuMessage2                 : ".\n\nCordialement.\nBruno.\n\n---\n\n"

    ----------------------------------------------------------------------------
    
    --                        VARIABLES MODIFIABLES
    
    ----------------------------------------------------------------------------


    -- Format du nom du fichier facture.
    property prefixNomFichierFacture :          "Facture_00"
    property extensionFichierFacture :          ".pdf"

    -- Fichier de log.
    property fichierLog              :          POSIX file "/Users/bruno/Desktop/log.txt"
    property referenceVersFichierLog :          ""
    
    -- Dossier temporaire où va être enregistrée la facture avant d'être envoyée au client.
    -- ATTENTION ! Chemin au format HFS car Excel ne comprend pas le POSIX.
    property cheminDossierTempFacture :         "Macintosh HD:Users:bruno:Desktop:_"
    
    
    
    
    ----------------------------------------------------------------------------
    
    --                      VARIABLES A NE PAS TOUCHER
    
    ----------------------------------------------------------------------------
    
    
    -- Variables de l'interface graphique
    -- IBOutlets
    property theWindow           : missing value
    property datePicker          : missing value
    property tfNomClient         : missing value
    property tfTypeIntervention  : missing value
    property tfDureeIntervention : missing value
    property tfNumeroDeFacture   : missing value
    property tfPrixHoraire       : missing value
    
    -- Variables nécessaire au fonctionnement du programme
    -- Autres propriétés du script
    property numeroFacture            : ""
    property cheminFichierTempFacture : ""
    property nomFichierFacture        : ""
    property cheminFactureFinal       : ""
    property dossierFacturesAlias     : ""
    -- Client
    property adresseCourrierClient    : ""
    property nomClient                : ""
    property adresseFormateeClient    : ""
    property formeJuridiqueClient     : ""
    -- Intervention
    property typeIntervention         : ""
    property dureeIntervention        : ""
    property dateIntervention         : ""
    

    ----------------------------------------------------------------------------

    --                     MÉTHODES DE LA FENÊTRE "window"

    ----------------------------------------------------------------------------
    
    
    
    --------------------  applicationWillFinishLaunching  -----------------------
     
    -- Fonction appelée juste avant que la fenêtre n'apparaisse à l'écran.
    -- Utile pour initialiser les variables.

    on applicationWillFinishLaunching:aNotification
		-- Insert code here to initialize your application before any files are opened
        
        set referenceVersFichierLog to open for access fichierLog with write permission
        
        my logToFileAndToConsole("")
        my logToFileAndToConsole("---------- Début du programme ------------")
        
        my recupParametres()
        my initialisationVariables()
        my initialisationInterface()
        
    end applicationWillFinishLaunching:
    
    
    ---------------------  applicationShouldTerminate  ------------------------
     
    -- Fonction appelée juste avant que le programme se termine.
    -- Utile pour fermer des fichiers et tout mettre en forme une dernière fois.

    on applicationShouldTerminate:sender
        
		-- Insert code here to do any housekeeping before your application quits
        my logToFileAndToConsole("----------- Fin du programme -------------")
        my logToFileAndToConsole("")
        
        close access referenceVersFichierLog
        
		return current application's NSTerminateNow
        
    end applicationShouldTerminate:
    
    
    ------------  applicationShouldTerminateAfterLastWindowClosed  -------------
     
    -- Fonction à redéfinir si on veut que le programme se termine quand on
    -- clique sur le bouton de fermeture de la fenêtre.
    
    on applicationShouldTerminateAfterLastWindowClosed:nsApplication
        
        tell nsApplication to terminate:me
        
    end applicationShouldTerminateAfterLastWindowClosed:
    
    
    ---------------------  clickBtnFacturer  ------------------------
    (*
     clickBtnFacturer : Fonction appelée lorsque l'on clique sur le bouton "Facturer".
     sender           : Le bouton "Facturer".
    *)
    
    on clickBtnFacturer:sender
        
        my logToFileAndToConsole("----> Nouvelle facture")
        
        -- Récupère les données utilisateurs entrées dans l'interface
        my recupDonneesDeLUtilisateur()
        
        -- Récupère les informations du client dans l'application Contacts
        my recupInfosClientDansContacts()
        
        -- Création de la facture dans Excel
        my creerFactureDansExcel()
        
        
        -- Renommage du fichier facture avec Finder
        my renommeFichierFacture()
        
        
        -- Envoi de la facture avec Mail
        my envoieFactureAvecMail()
        
        
        -- MAJ du numéro de facture et de l'interface        
        my initialisationVariables()
        my initialisationInterface()
        
        my logToFileAndToConsole("<---- Nouvelle facture\n")
        
    end clickBtnFacturer
    
   
   
   ----------------------------------------------------------------------------
    
    --                           AUTRES MÉTHODES
    
    ----------------------------------------------------------------------------
    
    
    
    ---------------------  recupParametres  ------------------------
    
    -- Récupération des paramètres du programme.
    
    on recupParametres()
        
        my logToFileAndToConsole("----> Récupération des paramètres")
        
        my logToFileAndToConsole("prixHoraire                      : " & prixHoraire)
        my logToFileAndToConsole("cheminFichierModeleFactureExcel  : " & cheminFichierModeleFactureExcel)
        my logToFileAndToConsole("cheminDossierFactures            : " & cheminDossierFactures)
        my logToFileAndToConsole("monAdresseCourrier               : " & monAdresseCourrier)
        my logToFileAndToConsole("maSignature                      : " & maSignature)
        my logToFileAndToConsole("monSujet                         : " & monSujet)
        my logToFileAndToConsole("contenuMessage1                  : " & contenuMessage1)
        my logToFileAndToConsole("contenuMessage2                  : " & contenuMessage2)
        my logToFileAndToConsole("prefixNomFichierFacture          : " & prefixNomFichierFacture)
        my logToFileAndToConsole("extensionFichierFacture          : " & extensionFichierFacture)
        my logToFileAndToConsole("fichierLog                       : " & fichierLog)
        my logToFileAndToConsole("referenceVersFichierLog          : " & referenceVersFichierLog)
        my logToFileAndToConsole("cheminDossierTempFacture         : " & cheminDossierTempFacture)
        
        my logToFileAndToConsole("<---- Récupération des paramètres")
        
    end recupParametres
    
    
    ---------------------  initialisationVariables  ------------------------
     
    -- Initialisation des variables du programme.
    
    on initialisationVariables()
        
        my logToFileAndToConsole("----> Initialisation des variables")
        
        -- numeroFacture : récupéré à partir du numéro de fichier de la dernière facture
        
        -- set dossierFacturesAlias to (POSIX file chemindossierFactures) as alias
        set dossierFacturesAlias to cheminDossierFactures as alias
        set numeroFacture to my getNumeroFacture(dossierFacturesAlias)
        my logToFileAndToConsole("Numéro de la facture        : " & numeroFacture)
        
        
        -- nomFichierFacture : nom du fichier facture de la forme "Facture_000X.pdf"
        
        set nomFichierFacture to prefixNomFichierFacture & (numeroFacture as text) & extensionFichierFacture
        my logToFileAndToConsole("Nom du fichier facture      : " & nomFichierFacture)
        
        
        -- cheminFichierTempFacture : Chemin vers le dossier temporaire de stockage de la facture
        
        set cheminFichierTempFacture to cheminDossierTempFacture & space & nomFichierFacture
        my logToFileAndToConsole("Chemin temporaire           : " & (POSIX path of cheminFichierTempFacture) as text)


        -- cheminFactureFinal : Chemin vers le dossier final de stockage de la facture
        
        set cheminFactureFinal to cheminDossierFactures & nomFichierFacture
        my logToFileAndToConsole("Chemin final                : " & (POSIX path of cheminFactureFinal) as text)


        my logToFileAndToConsole("<---- Initialisation des variables")
        
    end initialisationVariables
    
    
    
    ---------------------  initialisationInterface  ------------------------
     
    -- Initialisation de l'interface du programme.
    
    on initialisationInterface()
        
        -- Initialise le calendrier de l'interface avec la date du jour
        -- datePicker's setDateAS:(current date) -- Passage à 10.12
        datePicker's setDateValue:(current date)
        
        -- Affiche le numéro de facture en cours
        tfNumeroDeFacture's setStringValue:numeroFacture
        --tfNumeroDeFacture's setStringValue_(numeroFacture)
        
        -- Affiche le prix horaire
        tfPrixHoraire's setStringValue:prixHoraire
        --tfPrixHoraire's setStringValue_(prixHoraire)
        
        -- Affiche un exemple de nom de client
        tfNomClient's setStringValue:"ex: Boissonnet"
        
        tfTypeIntervention's setStringValue:"Réparation\nAssistance à domicile\nFormation"
        tfDureeIntervention's setStringValue:""
        
    end initialisationInterface
    
    

     ---------------------  dossierRessources  ------------------------
 
     -- Renvoi le chemin du dossier Ressources de l'application
     
    
    on dossierRessources()
        
        set pathToMe to ""
        
        tell current application's class "NSBundle"
            tell its mainBundle()
                set pathToMe to resourcePath() as string
            end tell
        end tell
        
        return pathToMe
        
    end dossierRessources


    ---------------------  recupDonneesDeLUtilisateur  ------------------------
 
    -- Récupère les données utilisateurs entrées dans l'interface
 
    on recupDonneesDeLUtilisateur()
        
        my logToFileAndToConsole("")
        
        set nomClient to (tfNomClient's stringValue()) as text
        my logToFileAndToConsole("Nom du Client à rechercher       : " & nomClient)
        
        set typeIntervention to (tfTypeIntervention's stringValue()) as text
        my logToFileAndToConsole("Type d'intervention sélectionné  : " & typeIntervention)
        
        set dureeIntervention to (tfDureeIntervention's stringValue()) as text
        my logToFileAndToConsole("Durée d'intervention entrée      : " & dureeIntervention)
       
        set dateIntervention to (datePicker's dateAS() as date)
        my logToFileAndToConsole("Date d'intervention sélectionnée : " & (date string of dateIntervention) as text)
        
    end recupDonneesDeLUtilisateur



    ---------------------  recupInfosClientDansContacts  ------------------------
 
    -- Récupère les informations du client dans l'application Contacts
    
    on recupInfosClientDansContacts()
        
        my logToFileAndToConsole("")
        
        set nomClient to (tfNomClient's stringValue()) as text
        my logToFileAndToConsole("Client recherché                 : " & nomClient)
        
        
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
            tell current application to my logToFileAndToConsole("Nom du client                    : " & nomClient)
            
            
            (* Récupération de la forme juridique *)
            if company of ficheClient is true then
                set formeJuridiqueClient to "Entreprise"
                else
                set formeJuridiqueClient to "Particulier"
            end if
            
            tell current application to my logToFileAndToConsole("Forme juridique du client        : " & formeJuridiqueClient)

            
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
            
            tell current application to my logToFileAndToConsole("Adresse mail du client           : " & adresseCourrierClient)
            
            (* Récupération de l'adresse à partir de la fiche client *)
            
            try
                set adresseClient to first item of address of ficheClient
                tell adresseClient
                    
                    -- gère les cas où il manque qqch dans l'adresse
                    if  (street is missing value) then
                        set adresseFormateeClient to zip & space & city
                        else
                        set adresseFormateeClient to street & space & zip & space & city
                    end if
                end tell
                
                tell current application to my logToFileAndToConsole("Adresse formatée du client       : " & adresseFormateeClient)
                
                on error
                
                display dialog "Aucune adresse renseignée pour le client \"" & nomClient & "\"." & ¬
                return & "Veuillez compléter la fiche client et relancer le programme." & return & ¬
                "Fin du programme." with title ¬
                "Facture" buttons {"Terminer"} cancel button "Terminer" default button "Terminer"
                
            end try
            
            
            quit
            
        end tell -- application "Contacts"
        
        
        -- MAJ de l'interface avec le nom complet du client
        tfNomClient's setStringValue:nomClient
        
        
    end recupInfosClientDansContacts
    
    
    
    ---------------------  creerFactureDansExcel  ------------------------
    
    -- Crée un fichier excel à partir d'un modèle, le modifie en fonction des données entrées et l'enregistre en pdf sur le bureau.
    -- Le fichier aura le nom: " Facture_00XXX.pdf" avec une espace devant ajoutée automatiquement par excel qui concatène le nom du fichier avec le nom de la feuille de calcul.
    
    on creerFactureDansExcel()
        
        
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
                set value of cell "B15" to adresseFormateeClient
                
                -- Mise à jour de la forme juridique du client
                set value of cell "B16" to formeJuridiqueClient
                
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
        
    
    end creerFactureDansExcel
    
    
    ---------------------  renommeFichierFacture  ------------------------
    
    -- On renome le fichier facture correctement (Impossible de le faire depuis Excel)
    -- et on le déplace dans le dossier des factures.
    
    on renommeFichierFacture()
        
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
        
    end renommeFichierFacture
    
    
    ---------------------  envoieFactureAvecMail  ------------------------
    
    -- Crée le message contenant la facture à envoyer au client.
    
    on envoieFactureAvecMail()
        
        (*
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
            
            set contenuMessage to contenuMessage1 & (date string of dateIntervention) & contenuMessage2
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
        
    end envoieFactureAvecMail
    
    
    ---------------------  list_position  ------------------------
    
    (*
     list_position : renvoie la position d'un élément dans une liste.
     this_item     : élément à rechercher.
     this_list     : liste dans laquelle on recherche.
     résultat      : la position de l'élément dans la liste, 0 si non trouvé.
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
    
    
    ---------------------  getNumeroFacture  ------------------------

    (*
     getNumeroFacture : renvoie le numéro de facture à partir du nom du dernier fichier
                        dans le dossier des factures (Facture_00212.pdf).
     dossierFactures  : dossier contenant les fichiers de facture.
     résultat         : le numéro de facture à utiliser pour facturer.
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



    ---------------------  logToFile  ------------------------
    
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
    
    ---------------------  logToFileAndToConsole  ------------------------
    
    -- Enregistre dans le fichier log et affiche dans la console (voir logToFile)
    
    on logToFileAndToConsole(ligne)
        
        logToFile(ligne)
        log(ligne)
        
    end logToFileAndToConsole
    
	
end script
