# Release Notes


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
