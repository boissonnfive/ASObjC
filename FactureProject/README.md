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


### Contruire le projet


### Changer le mode de Build en Release

1. Ouvrir XCode
2. Aller dans **Product > Scheme > Edit Scheme...**
3. Dans la colonne de gauche, sélectionner **Run**
4. Dans le panneau de droite, sélectionner l'onglet **Info**
5. Dans la liste déroulante **Build configuration**, sélectionner **Release**
6. Décocher **Debug executable**
7. Cliquer sur le bouton **Close**


### Créer une version "Release" du programme

1. Aller dans **Product > Clean**
2. Aller dans **Product > Build**


### Créer un fichier .dmg

1. Créez un dossier **Facture X.Y.Z** sur le bureau (avec X.Y.Z la dernière version du programme)
2. Dans XCode, faites un clic droit sur **Facture.app** dans la colonne de droite
3. Sélectionnez le menu **Show in Finder**
4. Copiez le fichier app **Facture** dans le dossier **Facture X.Y.Z**
5. Copiez un raccourci du dossier **Applications** dans le dossier **Facture X.Y.Z**
6. Avec le **Terminal**, créez un dossier caché **.background** dans le dossier **Facture X.Y.Z** :

        ($cd "~/Desktop/Facture 1.2.3" & mkdir .background & open .background)

7. Quittez le **Terminal** 
8. Mettez dans le dossier **.background** le fichier d'arrière-plan **Fond_DMG.png**
9. Ouvrez l'**Utilitaire de disque**
10. Allez dans **Fichier > Nouvelle image > Image d'un dossier...**
11. Sélectionnez le dossier **Facture X.Y.Z** sur le bureau et cliquez sur le bouton **Ouvrir**
12. Une nouvelle fenêtre s'ouvre
13. Dans le champ **Enregistrer sous : **, mettez **Facture 1.2.3**
14. Sélectionnez le dossier de destination (Documents par exemple)
15. Dans la liste déroulante **Chiffrement :**, laissez **aucun**
16. Dans la liste déroulante **Format d'image :**, sélectionnez **lecture/écriture**
17. Cliquez sur le bouton **Enregistrer**.
18. Allez dans le dossier **Documents** et montez l'image **Facture 1.2.3.dmg** en double-cliquant dessus
19. Sur le bureau, double-cliquez sur le disque **Facture 1.2.3**
20. Modifiez les options de présentations ( CMD + J) comme suit :

        - **Toujours présenter par icônes**       : coché
        - **Naviguer en présentation par icônes** : coché
        - **Organiser par**                       : Aucun
        - **Trier par**                           : Aucun
        - **Taille des icônes**                   : 144 x 144
        - **Espacement de la grille**             : maximum
        - **Taille du texte**                     : 12
        - **Position du texte**                   : En bas
        - **Afficher les informations**           : non coché
        - **Utiliser un aperçu comme icône**      : coché
        - **Arrière-plan**                        : Image (et faire glisser l'image du dossier .background de l'image DMG)
21. Pour l'image d'arrière-plan, ouvrez le **Terminal** et tapez :

        $ cd /Volumes/Facture\ 1.2.3/; open .background

22. Quittez le **Terminal**
23. Faites glisser l'image **Fond_DMG.png** sur le carré (en laissant la fenêtre Finder du disque **Facture 1.2.3** au premier plan)
24. Disposez les icônes à droite et à gauche du symbôle flêche.
25. Fermez la fenêtre Finder du disque **Facture 1.2.3**
26. Ouvrez l'**Utilitaire de disque**
27. Dans la colonne de droite, sélectionne l'image **Facture 1.2.3**
28. Allez dans **Fichier > Nouvelle image > Image de « Facture 1.2.3 »...**
29. Sélectionnez le dossier de destination (Documents par exemple)
30. Dans la liste déroulante **Format :**, sélectionnez **lecture/écriture**
31. Dans la liste déroulante **Chiffrement :**, laissez **aucun**
32. Cliquez sur le bouton **Enregistrer**.
33. Quittez l'**Utilitaire de disque**
34. Démontez le disque **Facture 1.2.3**
35. Allez dans le dossier **Documents** et montez l'image **Facture 1.2.3.dmg** en double-cliquant dessus
36. Sur le bureau, double-cliquez sur le disque **Facture 1.2.3**

