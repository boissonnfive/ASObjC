# Facture

## Description ##

Ce programme AppleScript permet de remplir une facture et de l'envoyer par mail.

Il a été codé en AppleScript pour tirer parti des applications suivantes :
- Contacts (base de données des clients)
- Excel (modèle de facture + export pdf)
- Mail (envoi de la facture au client)

Il a été codé en AppleScriptObjC afin de créer une interface utilisateur plus ergonomique.


## Prérequis ##

1. Un carnet d'adresse (application "Contacts") contenant les coordonnées des clients :
    - nom
    - prénom
    - adresse mail
    - adresse
    - type (entreprise ou particulier)
2. Excel
3. Un modèle de facture prédéfini au format excel.
4. Un dossier qui contiendra une copie des factures envoyées.
5. Un compte mail enregistré dans Mail (avec une signature)


## Algorithme de fonctionnement ##

1. Récupération des données de fonctionnement du programme :
	- chemin du dossier contenant le modèle excel,
	- nom du fichier modèle excel,
	- chemin du dossier contenant la copie des factures envoyées,
	- Nom du compte mail avec lequel envoyer la facture,
	- Contenu de l'objet du message,
	- Contenu du message à envoyer avec la facture,
	- Nom de la signature de ce compte mail.
2. Récupération des données utilisateur par une interface graphique :
	- durée d'intervention,
	- nom du client,
	- type d'intervention,
	- date d'intervention.
3. Appuie sur le bouton facturer.
4. Recherche du client à partir de son nom (si le nom n'est pas dans "Contacts", on termine le programme).
5. Récupération des informations du client dans "Contacts"
6. Création d'un nouveau numéro de facture (à partir du nom de la dernière facture dans le dossier qui contient la copie des factures envoyées).
7. Ouverture du modèle Excel.
8. Remplissage du modèle avec les informations obtenues :
	- numéro de facture,
	- date d'intervention,
	- nom complet du client,
	- adresse du client,
	- forme juridique du client,
	- type d'intervention,
	- durée de l'intervention,
	- coût de l'intervention.
9. Export du fichier au format PDF (sur le bureau de manière temporaire).
10. Ouverture de "Mail".
11. Création d'un nouveau message avec les informations obtenues :
	- adresse de l'expéditeur
	- adresse du destinataire
	- sujet du message
	- Contenu du message
	- signature
	- pièce jointe : la facture au format PDF
12. Déplacement de la facture dans le dossier d'archivage des factures.
13. Incrémentation du numéro de facture.


## Avant d'utiliser ce programme ##

### Introduction ###

Les prérequis impliquent que ce programme ne peut pas fonctionner sans une certaine préparation.


### Créer le bon contexte ###

1. S'assurer que "Contacts" contient bien les coordonnées des clients.
2. Installer Excel.
3. Créer (ou récupérer) un modèle de facture excel qui correspond au format défini plus bas.
4. Avoir un compte mail dans "Mail" et une signature qui y est associée.
5. Créer un dossier pour archiver les factures


### Modifier les variable du script ###

- cheminFichierModeleFactureExcel : le chemin du dossier contenant le modèle excel.
- cheminDossierFactures : le chemin du dossier d'archivage des factures envoyées au client.
- monAdresseCourrier : l'adresse utilisée pour envoyer le mail.
- maSignature : le nom de la signature utilisée pour envoyer le mail.
- monSujet : le sujet du mail.
- contenuMessage1 : contenu du message jusqu'à la date d'intervention.
- contenuMessage2 : fin du contenu du message.


### Format du nom de la facture ###

- son nom a le format suivant : "Facture_" + numéro à 5 chiffres précédés par des zéros + ".pdf" (exemple: Facture_00123.pdf)
- son numéro est incrémenté à chaque nouvelle facture
- le numéro de la facture est retrouvé à partir du nom de la dernière facture du dossier d'archivage (voir Prérequis point 4), et incrémenté de 1.


### Modèle de facture ###

Le modèle de facture doit être formaté de la manière suivante (*):

- Numéro de facture : cellules D4 et B11,
- date d'intervention : cellule B12,
- nom client : cellule B14,
- adresse du client : cellule B15,
- forme juridique : cellule B16,
- type d'intervention : cellule A21,
- durée d'intervention : cellule C21,
- coût de la facture : cellule D21.

(*) : voir fichier inclus dans le projet : Modele_Facture_v0_5.xltx.


