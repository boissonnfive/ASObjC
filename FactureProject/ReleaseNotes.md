# Release Notes

## Version 1.2.5

### Meilleure factorisation du code

- renommeElement
- deplaceElement
- envoieAvecMail
- ajouteSignatureDansMail
 
Et meilleur nommage des variables et des fonctions.


## Trace log ##

- On ne trace plus sur le bureau mais dans **~/Library/Logs/Facture.log**.
- Ajout de nombreuses traces dans le code


## Dossier temporaire ##

- On n'utilise plus le bureau comme dossier temporaire pour la facture, mais le dossier **Temporary items** de macOS.



### Numéro de version 1.2.5

- fenêtre de l'application (fichier MainMenu.xib)
- fenêtre « À propos » (fichier Facture-Info.plist)



## Version 1.2.4

### Modifications

- Création d'un fichier icône à partir d'une image "Invoice Icon" de http://www.visualpharm.com
- Ajout d'un remerciement dans le fichier **Credits.rtf**
- Modification du fichier **README.md** pour y ajouter la procédure de création d'une image .dmg

### Numéro de version 1.2.4

- fenêtre de l'application (fichier MainMenu.xib)
- fenêtre « À propos » (fichier Facture-Info.plist)


## Version 1.2.3

### Factorisation du code ... suite

- Création d'un procédure « recupDonneesDeLUtilisateur » qui récupère les données utilisateurs entrées dans l'interface
- Création d'un procédure « recupInfosClientDansContacts » qui écupère les informations du client dans l'application Contacts
- Création d'un procédure « creerFactureDansExcel » qui crée la facture dans Excel
- Création d'un procédure « renommeFichierFacture » qui renomme le fichier facture avec Finder
- Création d'un procédure « envoieFactureAvecMail » qui envoie la facture avec Mail


### Modifications

- Fonction clickButton => clickBtnFacturer
- Fonction « initialisationInterface » : MAJ TOUS les champs de l'interface
- Reformatage du log


### Créations

- Fonction « dossierRessources » qui renvoie le chemin du dossier ressources
- Fonction « recupParametres » qui trace les paramètres du programme


### Suppressions

- Fonction « getDate »
- Fonction « datePickerGet »
- Fonction « majVariablesFacture »
- Fonction « showAlertModal »


### Date du copyright

- fenêtre de l'application (fichier MainMenu.xib)
- fenêtre « À propos » (fichier Facture-Info.plist)


### Numéro de version

- fenêtre de l'application (fichier MainMenu.xib)
- fenêtre « À propos » (fichier Facture-Info.plist)


---

## Version 1.2.2

(Factorisation du code)
- Création d'un procédure « initialisationVariables » qui contient les instructions d'initialisation des variables
- Création d'un procédure « initialisationInterface » qui contient les instructions d'initialisation de l'interface
- Modification des messages contenuMessage1 et contenuMessage2 pour remplacer les sauts de ligne par "\n"
- Suppression des sauts de ligne pour la première ligne du fichier log.txt
- Formatage des commentaires de fonctions/procédures
- Harmonisation des écritures des fonctions/procédures :
applicationShouldTerminate_(sender) => applicationShouldTerminate:sender
applicationShouldTerminate_         => applicationShouldTerminate:
- Bug : quand la rue n'est pas renseignée, l'adresse n'est pas écrite dans la facture
- Remplissage des infos dans la fenêtre « À propos »

## Version 1.2.1

(Modification du code pour que le programme fonctionne sous macOS Sierra 10.12)
- Ajout de contraintes sur les éléments graphiques
- Ajout d'un caractère dans le nom temporaire du fichier PDF (cheminDossierTempFacture) :
"Macintosh HD:Users:bruno:Desktop:" => "Macintosh HD:Users:bruno:Desktop:_"


## Version 1.1.2

- Ajout d'un fichier "ReleaseNotes.md" pour décrire les modifications faites pour chaque version.
- Ajout d'un champ "prixHoraire" pour modifier le prix à l'heure de la facture.
- Modification de l'interface graphique pour afficher le prix horaire.
- Modification du numéro de version sur l'interface graphique.
- Modification de la date du copyright (2016).
