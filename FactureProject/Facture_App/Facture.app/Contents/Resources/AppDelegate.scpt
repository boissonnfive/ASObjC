FasdUAS 1.101.10   ��   ��    k             l     ��������  ��  ��        l     �� 	 
��   	    AppDelegate.applescript    
 �   2     A p p D e l e g a t e . a p p l e s c r i p t      l     ��  ��     	  Facture     �        F a c t u r e      l     ��������  ��  ��        l     ��  ��    2 ,  Created by Bruno Boissonnet on 07/09/2014.     �   X     C r e a t e d   b y   B r u n o   B o i s s o n n e t   o n   0 7 / 0 9 / 2 0 1 4 .      l     ��  ��    = 7  Copyright (c) 2014 BInfoService. All rights reserved.     �   n     C o p y r i g h t   ( c )   2 0 1 4   B I n f o S e r v i c e .   A l l   r i g h t s   r e s e r v e d .      l     ��������  ��  ��         l     ��������  ��  ��      ! " ! h     �� #�� 0 appdelegate AppDelegate # k       $ $  % & % j     �� '
�� 
pare ' 4     �� (
�� 
pcls ( m     ) ) � * *  N S O b j e c t &  + , + l     ��������  ��  ��   ,  - . - l     ��������  ��  ��   .  / 0 / l     �� 1 2��   1 P J--------------------------------------------------------------------------    2 � 3 3 � - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 0  4 5 4 l     ��������  ��  ��   5  6 7 6 l     �� 8 9��   8 2 ,                        VARIABLES A MODIFIER    9 � : : X                                                 V A R I A B L E S   A   M O D I F I E R 7  ; < ; l     ��������  ��  ��   <  = > = l     �� ? @��   ? P J--------------------------------------------------------------------------    @ � A A � - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - >  B C B l     ��������  ��  ��   C  D E D l     �� F G��   F   Prix horaire    G � H H    P r i x   h o r a i r e E  I J I j   	 �� K�� 0 prixhoraire prixHoraire K m   	 
 L L @9       J  M N M l     �� O P��   O ( " dossier contenant le mod�le excel    P � Q Q D   d o s s i e r   c o n t e n a n t   l e   m o d � l e   e x c e l N  R S R j    �� T�� B0 cheminfichiermodelefactureexcel cheminFichierModeleFactureExcel T m     U U � V V x / U s e r s / b r u n o / D o c u m e n t s / b i n f o s e r v i c e / M o d e l e _ F a c t u r e _ v 0 _ 5 . x l t x S  W X W l     �� Y Z��   Y ' ! dossier d'archivage des factures    Z � [ [ B   d o s s i e r   d ' a r c h i v a g e   d e s   f a c t u r e s X  \ ] \ j    �� ^�� .0 chemindossierfactures cheminDossierFactures ^ m     _ _ � ` ` r M a c i n t o s h   H D : U s e r s : b r u n o : D o c u m e n t s : b i n f o s e r v i c e : F a c t u r e s : ]  a b a l     �� c d��   c   mail    d � e e 
   m a i l b  f g f j    �� h�� (0 monadressecourrier monAdresseCourrier h m     i i � j j , b i n f o s e r v i c e @ g m a i l . c o m g  k l k j    �� m�� 0 masignature maSignature m m     n n � o o  S i g n a t u r e   n � 1 l  p q p j    �� r�� 0 monsujet monSujet r m     s s � t t ( F a c t u r e   B I n f o S e r v i c e q  u v u j    �� w�� "0 contenumessage1 contenuMessage1 w m     x x � y y � B o n j o u r , 
 
 v e u i l l e z   t r o u v e r   c i - j o i n t e   l a   f a c t u r e   d e   l ' i n t e r v e n t i o n   d u   v  z { z j    "�� |�� "0 contenumessage2 contenuMessage2 | m    ! } } � ~ ~ < . 
 
 C o r d i a l e m e n t . 
 B r u n o . 
 
 - - - 
 
 {   �  l     ��������  ��  ��   �  � � � l     �� � ���   � P J--------------------------------------------------------------------------    � � � � � - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - �  � � � l     ��������  ��  ��   �  � � � l     �� � ���   � 3 -                        VARIABLES MODIFIABLES    � � � � Z                                                 V A R I A B L E S   M O D I F I A B L E S �  � � � l     ��������  ��  ��   �  � � � l     �� � ���   � P J--------------------------------------------------------------------------    � � � � � - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - �  � � � l     ��������  ��  ��   �  � � � l     ��������  ��  ��   �  � � � l     �� � ���   � ( " Format du nom du fichier facture.    � � � � D   F o r m a t   d u   n o m   d u   f i c h i e r   f a c t u r e . �  � � � j   # '�� ��� 20 prefixnomfichierfacture prefixNomFichierFacture � m   # & � � � � �  F a c t u r e _ 0 0 �  � � � j   ( ,�� ��� 20 extensionfichierfacture extensionFichierFacture � m   ( + � � � � �  . p d f �  � � � l     ��������  ��  ��   �  � � � l     �� � ���   �   Fichier de log.    � � � �     F i c h i e r   d e   l o g . �  � � � j   - 7�� ��� 0 
fichierlog 
fichierLog � 4   - 6�� �
�� 
psxf � m   1 4 � � � � � 8 / U s e r s / b r u n o / D e s k t o p / l o g . t x t �  � � � j   8 <�� ��� 20 referenceversfichierlog referenceVersFichierLog � m   8 ; � � � � �   �  � � � l     ��������  ��  ��   �  � � � l     �� � ���   � [ U Dossier temporaire o� va �tre enregistr�e la facture avant d'�tre envoy�e au client.    � � � � �   D o s s i e r   t e m p o r a i r e   o �   v a   � t r e   e n r e g i s t r � e   l a   f a c t u r e   a v a n t   d ' � t r e   e n v o y � e   a u   c l i e n t . �  � � � l     �� � ���   � K E ATTENTION ! Chemin au format HFS car Excel ne comprend pas le POSIX.    � � � � �   A T T E N T I O N   !   C h e m i n   a u   f o r m a t   H F S   c a r   E x c e l   n e   c o m p r e n d   p a s   l e   P O S I X . �  � � � j   = A�� ��� 40 chemindossiertempfacture cheminDossierTempFacture � m   = @ � � � � � D M a c i n t o s h   H D : U s e r s : b r u n o : D e s k t o p : _ �  � � � l     ��������  ��  ��   �  � � � l     ��������  ��  ��   �  � � � l     ��������  ��  ��   �  � � � l     ��������  ��  ��   �  � � � l     �� � ���   � P J--------------------------------------------------------------------------    � � � � � - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - �  � � � l     ��������  ��  ��   �  � � � l     �� � ���   � 6 0                      VARIABLES A NE PAS TOUCHER    � � � � `                                             V A R I A B L E S   A   N E   P A S   T O U C H E R �  � � � l     ��������  ��  ��   �  � � � l     �� � ���   � P J--------------------------------------------------------------------------    � � � � � - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - �  � � � l     ��������  ��  ��   �  � � � l     ��������  ��  ��   �  � � � l     �� � ���   � ) # Variables de l'interface graphique    � � � � F   V a r i a b l e s   d e   l ' i n t e r f a c e   g r a p h i q u e �  � � � l     �� � ���   �  
 IBOutlets    � � � �    I B O u t l e t s �  � � � j   B F�� ��� 0 	thewindow 	theWindow � m   B E��
�� 
msng �  � � � j   G K�� ��� 0 
datepicker 
datePicker � m   G J��
�� 
msng �  � � � j   L R�� ��� 0 tfnomclient tfNomClient � m   L O��
�� 
msng �  � � � j   S Y�� ��� (0 tftypeintervention tfTypeIntervention � m   S V��
�� 
msng �  �  � j   Z `���� *0 tfdureeintervention tfDureeIntervention m   Z ]��
�� 
msng   j   a g���� &0 tfnumerodefacture tfNumeroDeFacture m   a d��
�� 
msng  j   h n���� 0 tfprixhoraire tfPrixHoraire m   h k��
�� 
msng 	 l     ����~��  �  �~  	 

 l     �}�}   : 4 Variables n�cessaire au fonctionnement du programme    � h   V a r i a b l e s   n � c e s s a i r e   a u   f o n c t i o n n e m e n t   d u   p r o g r a m m e  l     �|�|   "  Autres propri�t�s du script    � 8   A u t r e s   p r o p r i � t � s   d u   s c r i p t  j   o u�{�{ 0 numerofacture numeroFacture m   o r �    j   v |�z�z 40 cheminfichiertempfacture cheminFichierTempFacture m   v y �    j   } ��y �y &0 nomfichierfacture nomFichierFacture  m   } �!! �""   #$# j   � ��x%�x (0 cheminfacturefinal cheminFactureFinal% m   � �&& �''  $ ()( j   � ��w*�w ,0 dossierfacturesalias dossierFacturesAlias* m   � �++ �,,  ) -.- l     �v/0�v  / 4 .property dateDuJourCourte :                 ""   0 �11 \ p r o p e r t y   d a t e D u J o u r C o u r t e   :                                   " ". 232 j   � ��u4�u .0 adressecourrierclient adresseCourrierClient4 m   � �55 �66  3 787 j   � ��t9�t 0 	nomclient 	nomClient9 m   � �:: �;;  8 <=< l     �s>?�s  > 4 .property ficheClient :                      ""   ? �@@ \ p r o p e r t y   f i c h e C l i e n t   :                                             " "= ABA j   � ��rC�r $0 typeintervention typeInterventionC m   � �DD �EE  B FGF j   � ��qH�q &0 dureeintervention dureeInterventionH m   � �II �JJ  G KLK j   � ��pM�p $0 dateintervention dateInterventionM m   � �NN �OO  L PQP l     �o�n�m�o  �n  �m  Q RSR l     �l�k�j�l  �k  �j  S TUT l     �iVW�i  V P J--------------------------------------------------------------------------   W �XX � - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -U YZY l     �h�g�f�h  �g  �f  Z [\[ l     �e]^�e  ] : 4                      applicationWillFinishLaunching   ^ �__ h                                             a p p l i c a t i o n W i l l F i n i s h L a u n c h i n g\ `a` l     �d�c�b�d  �c  �b  a bcb l     �ade�a  d P J--------------------------------------------------------------------------   e �ff � - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -c ghg l     �`�_�^�`  �_  �^  h iji l      �]kl�]  k
     applicationWillFinishLaunching :   Fonction appel�e juste avant que la
                                        fen�tre n'appara�sse � l'�cran.
                                        Utile pour initialiser les variables.
     aNotification                  :   ???.
       l �mm( 
           a p p l i c a t i o n W i l l F i n i s h L a u n c h i n g   :       F o n c t i o n   a p p e l � e   j u s t e   a v a n t   q u e   l a 
                                                                                 f e n � t r e   n ' a p p a r a � s s e   �   l ' � c r a n . 
                                                                                 U t i l e   p o u r   i n i t i a l i s e r   l e s   v a r i a b l e s . 
           a N o t i f i c a t i o n                                     :       ? ? ? . 
        j non l     �\�[�Z�\  �[  �Z  o pqp i   � �rsr I      �Yt�X�Y B0 applicationwillfinishlaunching_ applicationWillFinishLaunching_t u�Wu o      �V�V 0 anotification aNotification�W  �X  s k    vv wxw l     �Uyz�U  y R L Insert code here to initialize your application before any files are opened   z �{{ �   I n s e r t   c o d e   h e r e   t o   i n i t i a l i z e   y o u r   a p p l i c a t i o n   b e f o r e   a n y   f i l e s   a r e   o p e n e dx |}| l     �T�S�R�T  �S  �R  } ~~ r     ��� I    �Q��
�Q .rdwropenshor       file� o     �P�P 0 
fichierlog 
fichierLog� �O��N
�O 
perm� m    �M
�M boovtrue�N  � o      �L�L 20 referenceversfichierlog referenceVersFichierLog ��� l   �K�J�I�K  �J  �I  � ��� n   ��� I    �H��G�H .0 logtofileandtoconsole logToFileAndToConsole� ��F� m    �� ��� X 
 - - - - - - - - - -   D � b u t   d u   p r o g r a m m e   - - - - - - - - - - - - 
�F  �G  �  f    � ��� l   �E�D�C�E  �D  �C  � ��� l   �B�A�@�B  �A  �@  � ��� l   �?���?  � L F Initialisation des variables du programme ---------------------------   � ��� �   I n i t i a l i s a t i o n   d e s   v a r i a b l e s   d u   p r o g r a m m e   - - - - - - - - - - - - - - - - - - - - - - - - - - -� ��� l   �>�=�<�>  �=  �<  � ��� n   ��� I    �;��:�; .0 logtofileandtoconsole logToFileAndToConsole� ��9� m    �� ��� D - - - - >   I n i t i a l i s a t i o n   d e s   v a r i a b l e s�9  �:  �  f    � ��� l     �8�7�6�8  �7  �6  � ��� l     �5�4�3�5  �4  �3  � ��� l      �2���2  � U O numeroFacture : r�cup�r� � partir du num�ro de fichier de la derni�re facture    � ��� �   n u m e r o F a c t u r e   :   r � c u p � r �   �   p a r t i r   d u   n u m � r o   d e   f i c h i e r   d e   l a   d e r n i � r e   f a c t u r e  � ��� l     �1�0�/�1  �0  �/  � ��� l     �.���.  � N H set dossierFacturesAlias to (POSIX file chemindossierFactures) as alias   � ��� �   s e t   d o s s i e r F a c t u r e s A l i a s   t o   ( P O S I X   f i l e   c h e m i n d o s s i e r F a c t u r e s )   a s   a l i a s� ��� r     -��� c     '��� o     %�-�- .0 chemindossierfactures cheminDossierFactures� m   % &�,
�, 
alis� o      �+�+ ,0 dossierfacturesalias dossierFacturesAlias� ��� r   . >��� n  . 8��� I   / 8�*��)�* $0 getnumerofacture getNumeroFacture� ��(� o   / 4�'�' ,0 dossierfacturesalias dossierFacturesAlias�(  �)  �  f   . /� o      �&�& 0 numerofacture numeroFacture� ��� n  ? K��� I   @ K�%��$�% .0 logtofileandtoconsole logToFileAndToConsole� ��#� b   @ G��� m   @ A�� ���  N u m � r o   :  � o   A F�"�" 0 numerofacture numeroFacture�#  �$  �  f   ? @� ��� l  L L�!� ��!  �   �  � ��� l  L L����  �  �  � ��� l   L L����  � Q K nomFichierFacture : nom du fichier facture de la forme "Facture_000X.pdf"    � ��� �   n o m F i c h i e r F a c t u r e   :   n o m   d u   f i c h i e r   f a c t u r e   d e   l a   f o r m e   " F a c t u r e _ 0 0 0 X . p d f "  � ��� l  L L����  �  �  � ��� r   L e��� b   L _��� b   L Y��� o   L Q�� 20 prefixnomfichierfacture prefixNomFichierFacture� l  Q X���� c   Q X��� o   Q V�� 0 numerofacture numeroFacture� m   V W�
� 
ctxt�  �  � o   Y ^�� 20 extensionfichierfacture extensionFichierFacture� o      �� &0 nomfichierfacture nomFichierFacture� ��� n  f r��� I   g r���� .0 logtofileandtoconsole logToFileAndToConsole� ��� b   g n��� m   g h�� ��� " N o m   d u   f i c h i e r   :  � o   h m�� &0 nomfichierfacture nomFichierFacture�  �  �  f   f g� ��� l  s s���
�  �  �
  � ��� l  s s�	���	  �  �  � ��� l   s s����  � ^ X cheminFichierTempFacture : Chemin vers le dossier temporaire de stockage de la facture    � ��� �   c h e m i n F i c h i e r T e m p F a c t u r e   :   C h e m i n   v e r s   l e   d o s s i e r   t e m p o r a i r e   d e   s t o c k a g e   d e   l a   f a c t u r e  � ��� l  s s����  �  �  � ��� r   s ���� b   s ���� b   s z��� o   s x�� 40 chemindossiertempfacture cheminDossierTempFacture� 1   x y�
� 
spac� o   z � �  &0 nomfichierfacture nomFichierFacture� o      ���� 40 cheminfichiertempfacture cheminFichierTempFacture� ��� n  � �� � I   � ������� .0 logtofileandtoconsole logToFileAndToConsole �� c   � � b   � � m   � � � ( C h e m i n   t e m p o r a i r e   :   l  � �	����	 n   � �

 1   � ���
�� 
psxp o   � ����� 40 cheminfichiertempfacture cheminFichierTempFacture��  ��   m   � ���
�� 
ctxt��  ��     f   � ��  l  � ���������  ��  ��    l  � ���������  ��  ��    l   � �����   S M cheminFactureFinal : Chemin vers le dossier final de stockage de la facture     � �   c h e m i n F a c t u r e F i n a l   :   C h e m i n   v e r s   l e   d o s s i e r   f i n a l   d e   s t o c k a g e   d e   l a   f a c t u r e    l  � ���������  ��  ��    r   � � b   � � o   � ����� .0 chemindossierfactures cheminDossierFactures o   � ����� &0 nomfichierfacture nomFichierFacture o      ���� (0 cheminfacturefinal cheminFactureFinal  n  � �  I   � ���!���� .0 logtofileandtoconsole logToFileAndToConsole! "��" c   � �#$# b   � �%&% m   � �'' �((  C h e m i n   f i n a l   :  & l  � �)����) n   � �*+* 1   � ���
�� 
psxp+ o   � ����� (0 cheminfacturefinal cheminFactureFinal��  ��  $ m   � ���
�� 
ctxt��  ��     f   � � ,-, l  � ���������  ��  ��  - ./. l   � ���01��  0 S M contenuMessage1 et contenuMessage2 : contenu du message � envoyer au client    1 �22 �   c o n t e n u M e s s a g e 1   e t   c o n t e n u M e s s a g e 2   :   c o n t e n u   d u   m e s s a g e   �   e n v o y e r   a u   c l i e n t  / 343 l  � ���������  ��  ��  4 565 l  � ���78��  7 L F ATTENTION ! Il ne faut pas de tabulations ou d'espaces dans le texte.   8 �99 �   A T T E N T I O N   !   I l   n e   f a u t   p a s   d e   t a b u l a t i o n s   o u   d ' e s p a c e s   d a n s   l e   t e x t e .6 :;: l  � ���<=��  < / ) Attention lors de l'indentation du code.   = �>> R   A t t e n t i o n   l o r s   d e   l ' i n d e n t a t i o n   d u   c o d e .; ?@? l   � ���AB��  A w q
        set contenuMessage1 to "Bonjour,

veuillez trouver ci-jointe la facture de l'intervention du "
            B �CC � 
                 s e t   c o n t e n u M e s s a g e 1   t o   " B o n j o u r , 
 
 v e u i l l e z   t r o u v e r   c i - j o i n t e   l a   f a c t u r e   d e   l ' i n t e r v e n t i o n   d u   " 
                  @ DED n  � �FGF I   � ���H���� .0 logtofileandtoconsole logToFileAndToConsoleH I��I b   � �JKJ m   � �LL �MM J C o n t e n u   d u   m e s s a g e   ( 1 � r e   p a r t i e   )   :   
K o   � ����� "0 contenumessage1 contenuMessage1��  ��  G  f   � �E NON l   � ���PQ��  P P J
        set contenuMessage2 to ".

Cordialement.
Bruno.

---

"
            Q �RR � 
                 s e t   c o n t e n u M e s s a g e 2   t o   " . 
 
 C o r d i a l e m e n t . 
 B r u n o . 
 
 - - - 
 
 " 
                  O STS n  � �UVU I   � ���W���� .0 logtofileandtoconsole logToFileAndToConsoleW X��X b   � �YZY m   � �[[ �\\ J C o n t e n u   d u   m e s s a g e   ( 2 � m e   p a r t i e   )   :   
Z o   � ����� "0 contenumessage2 contenuMessage2��  ��  V  f   � �T ]^] l  � ���������  ��  ��  ^ _`_ l   � ���ab��  a < 6 dateDuJourCourte : Date du jour au format JJ/MM/AAAA    b �cc l   d a t e D u J o u r C o u r t e   :   D a t e   d u   j o u r   a u   f o r m a t   J J / M M / A A A A  ` ded l  � ���������  ��  ��  e fgf l  � ���hi��  h V Pset dateDuJourCourte to short date string of (current date) -- format JJ/MM/AAAA   i �jj � s e t   d a t e D u J o u r C o u r t e   t o   s h o r t   d a t e   s t r i n g   o f   ( c u r r e n t   d a t e )   - -   f o r m a t   J J / M M / A A A Ag klk l  � ���mn��  m U Omy logToFileAndToConsole("Date du jour (forme JJ/MM/AA) : " & dateDuJourCourte)   n �oo � m y   l o g T o F i l e A n d T o C o n s o l e ( " D a t e   d u   j o u r   ( f o r m e   J J / M M / A A )   :   "   &   d a t e D u J o u r C o u r t e )l pqp l  � ���������  ��  ��  q rsr l  � ���tu��  t M G Initialisation de l'interface du programme ---------------------------   u �vv �   I n i t i a l i s a t i o n   d e   l ' i n t e r f a c e   d u   p r o g r a m m e   - - - - - - - - - - - - - - - - - - - - - - - - - - -s wxw l  � ���������  ��  ��  x yzy l   � ���{|��  { D > Initialise le calendrier de l'interface avec la date du jour    | �}} |   I n i t i a l i s e   l e   c a l e n d r i e r   d e   l ' i n t e r f a c e   a v e c   l a   d a t e   d u   j o u r  z ~~ l  � ���������  ��  ��   ��� l  � �������  � " set thisDate to current date   � ��� 8 s e t   t h i s D a t e   t o   c u r r e n t   d a t e� ��� l  � �������  � F @my logToFileAndToConsole("Date du jour : " & (thisDate as text))   � ��� � m y   l o g T o F i l e A n d T o C o n s o l e ( " D a t e   d u   j o u r   :   "   &   ( t h i s D a t e   a s   t e x t ) )� ��� l  � �������  � % datePicker's setDateAS:thisDate   � ��� > d a t e P i c k e r ' s   s e t D a t e A S : t h i s D a t e� ��� l  � �������  � * $datePicker's setDateAS(current date)   � ��� H d a t e P i c k e r ' s   s e t D a t e A S ( c u r r e n t   d a t e )� ��� n  � ���� I   � �������� 0 setdatevalue_ setDateValue_� ���� I  � �������
�� .misccurdldt    ��� null��  ��  ��  ��  � o   � ����� 0 
datepicker 
datePicker� ��� l  � ���������  ��  ��  � ��� n  � ���� I   � �������� "0 setstringvalue_ setStringValue_� ���� o   � ����� 0 numerofacture numeroFacture��  ��  � o   � ����� &0 tfnumerodefacture tfNumeroDeFacture� ��� l  � ���������  ��  ��  � ��� n  ���� I   �������� "0 setstringvalue_ setStringValue_� ���� o   � ����� 0 prixhoraire prixHoraire��  ��  � o   � ����� 0 tfprixhoraire tfPrixHoraire� ��� l ��������  ��  ��  � ��� n 
��� I  
������� .0 logtofileandtoconsole logToFileAndToConsole� ���� m  �� ��� D < - - - -   I n i t i a l i s a t i o n   d e s   v a r i a b l e s��  ��  �  f  � ���� l ��������  ��  ��  ��  q ��� l     ��������  ��  ��  � ��� l     ��������  ��  ��  � ��� l     ��������  ��  ��  � ��� l     ��������  ��  ��  � ��� l     ��������  ��  ��  � ��� l     ������  � P J--------------------------------------------------------------------------   � ��� � - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -� ��� l     ��������  ��  ��  � ��� l     ������  � 6 0                      applicationShouldTerminate   � ��� `                                             a p p l i c a t i o n S h o u l d T e r m i n a t e� ��� l     �������  ��  �  � ��� l     �~���~  � P J--------------------------------------------------------------------------   � ��� � - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -� ��� l     �}�|�{�}  �|  �{  � ��� l      �z���z  �%
     applicationShouldTerminate :   Fonction appel�e juste avant que le programme
                                    se termine. Utile pour fermer des fichiers et
                                    tout mettre en forme une derni�re fois.
     sender                     :   ???.
        � ���> 
           a p p l i c a t i o n S h o u l d T e r m i n a t e   :       F o n c t i o n   a p p e l � e   j u s t e   a v a n t   q u e   l e   p r o g r a m m e 
                                                                         s e   t e r m i n e .   U t i l e   p o u r   f e r m e r   d e s   f i c h i e r s   e t 
                                                                         t o u t   m e t t r e   e n   f o r m e   u n e   d e r n i � r e   f o i s . 
           s e n d e r                                           :       ? ? ? . 
          � ��� l     �y�x�w�y  �x  �w  � ��� i   � ���� I      �v��u�v :0 applicationshouldterminate_ applicationShouldTerminate_� ��t� o      �s�s 
0 sender  �t  �u  � k     �� ��� l     �r�q�p�r  �q  �p  � ��� l     �o���o  � L F Insert code here to do any housekeeping before your application quits   � ��� �   I n s e r t   c o d e   h e r e   t o   d o   a n y   h o u s e k e e p i n g   b e f o r e   y o u r   a p p l i c a t i o n   q u i t s� ��� n    ��� I    �n��m�n .0 logtofileandtoconsole logToFileAndToConsole� ��l� m    �� ��� X 
 - - - - - - - - - - -   F i n   d u   p r o g r a m m e   - - - - - - - - - - - - - 
�l  �m  �  f     � ��� l   �k�j�i�k  �j  �i  � ��� I   �h��g
�h .rdwrclosnull���     ****� o    �f�f 20 referenceversfichierlog referenceVersFichierLog�g  � ��� l   �e�d�c�e  �d  �c  � ��� L    �� n   ��� o    �b�b  0 nsterminatenow NSTerminateNow� m    �a
�a misccura� ��`� l   �_�^�]�_  �^  �]  �`  � ��� l     �\�[�Z�\  �[  �Z  �    l     �Y�X�W�Y  �X  �W    l     �V�U�T�V  �U  �T    l     �S�R�Q�S  �R  �Q    l     �P	�P   P J--------------------------------------------------------------------------   	 �

 � - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  l     �O�N�M�O  �N  �M    l     �L�L   / )                              clickButton    � R                                                             c l i c k B u t t o n  l     �K�J�I�K  �J  �I    l     �H�H   P J--------------------------------------------------------------------------    � � - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  l     �G�F�E�G  �F  �E    l     �D�C�B�D  �C  �B    l      �A �A   � �
     clickButton :   Fonction appel�e lorsque l'on clique sur le bouton "Facturer".
     sender      :   Le bouton "Facturer".
          �!!
 
           c l i c k B u t t o n   :       F o n c t i o n   a p p e l � e   l o r s q u e   l ' o n   c l i q u e   s u r   l e   b o u t o n   " F a c t u r e r " . 
           s e n d e r             :       L e   b o u t o n   " F a c t u r e r " . 
           "#" l     �@�?�>�@  �?  �>  # $%$ i   � �&'& I      �=(�<�= 0 clickbutton_ clickButton_( )�;) o      �:�: 
0 sender  �;  �<  ' k    �** +,+ l     �9�8�7�9  �8  �7  , -.- n    /0/ I    �61�5�6 .0 logtofileandtoconsole logToFileAndToConsole1 2�42 m    33 �44 . 
 - - - - >   N o u v e l l e   f a c t u r e�4  �5  0  f     . 565 l   �3�2�1�3  �2  �1  6 787 l   �0�/�.�0  �/  �.  8 9:9 l    �-;<�-  ; x r**************************** R�cup�ration des informations du client dans Contacts *******************************   < �== � * * * * * * * * * * * * * * * * * * * * * * * * * * * *   R � c u p � r a t i o n   d e s   i n f o r m a t i o n s   d u   c l i e n t   d a n s   C o n t a c t s   * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *: >?> l   �,�+�*�,  �+  �*  ? @A@ r    BCB c    DED l   F�)�(F n   GHG I    �'�&�%�' 0 stringvalue stringValue�&  �%  H o    �$�$ 0 tfnomclient tfNomClient�)  �(  E m    �#
�# 
ctxtC o      �"�" 0 	nomclient 	nomClientA IJI n   %KLK I    %�!M� �! .0 logtofileandtoconsole logToFileAndToConsoleM N�N b    !OPO m    QQ �RR & C l i e n t   r e c h e r c h �   :  P o     �� 0 	nomclient 	nomClient�  �   L  f    J STS l  & &����  �  �  T UVU l  & &����  �  �  V WXW l   & &�YZ�  Y 9 3 Cherche le client dans l'application � Contacts �    Z �[[ f   C h e r c h e   l e   c l i e n t   d a n s   l ' a p p l i c a t i o n   �   C o n t a c t s   �  X \]\ l  &�^_`^ O   &�aba k   *�cc ded l  * *����  �  �  e fgf l  * *�hi�  h c ]tell current application to my logToFileAndToConsole("On appelle l'application \"Contacts\"")   i �jj � t e l l   c u r r e n t   a p p l i c a t i o n   t o   m y   l o g T o F i l e A n d T o C o n s o l e ( " O n   a p p e l l e   l ' a p p l i c a t i o n   \ " C o n t a c t s \ " " )g klk l  * *����  �  �  l mnm l   * *�op�  o 7 1 R�cup�ration de la fiche client � partir du nom    p �qq b   R � c u p � r a t i o n   d e   l a   f i c h e   c l i e n t   �   p a r t i r   d u   n o m  n rsr l  * *����  �  �  s tut l  * *�vw�  v 7 1 Liste de toutes les personnes qui portent ce nom   w �xx b   L i s t e   d e   t o u t e s   l e s   p e r s o n n e s   q u i   p o r t e n t   c e   n o mu yzy r   * <{|{ 6  * :}~} 2   * -�

�
 
azf4~ E   . 9� 1   / 1�	
�	 
pnam� o   2 8�� 0 	nomclient 	nomClient| o      �� *0 ficheclienttrouvees ficheclientTrouveesz ��� l  = =����  �  �  � ��� Z   =,����� =   = D��� l  = B���� I  = B� ���
�  .corecnte****       ****� o   = >���� *0 ficheclienttrouvees ficheclientTrouvees��  �  �  � m   B C����  � k   G ��� ��� l  G G��������  ��  ��  � ��� O  G R��� n  K Q��� I   L Q������� .0 logtofileandtoconsole logToFileAndToConsole� ���� m   L M�� ��� @ A u c u n e   p e r s o n n e   n e   p o r t e   c e   n o m .��  ��  �  f   K L� m   G H��
�� misccura� ��� l   S S������  � @ : Aucune personne ne porte ce nom, on d�clenche une erreur	   � ��� t   A u c u n e   p e r s o n n e   n e   p o r t e   c e   n o m ,   o n   d � c l e n c h e   u n e   e r r e u r 	� ��� l  S S��������  ��  ��  � ��� I  S �����
�� .sysodlogaskr        TEXT� b   S f��� b   S b��� b   S `��� b   S ^��� b   S \��� b   S Z��� m   S T�� ��� T A u c u n   c o n t a c t   d o n t   l e   n o m   d e   f a m i l l e   e s t   "� l 	 T Y������ o   T Y���� 0 	nomclient 	nomClient��  ��  � m   Z [�� ��� " "   n ' a   � t �   t r o u v � .� o   \ ]��
�� 
ret � l 	 ^ _������ m   ^ _�� ��� f V e u i l l e z   c r � e r   l e   c o n t a c t   e t   r e l a n c e r   l e   p r o g r a m m e .��  ��  � o   ` a��
�� 
ret � l 	 b e������ m   b e�� ��� " F i n   d u   p r o g r a m m e .��  ��  � ����
�� 
appr� l 	 i l������ m   i l�� ���  F a c t u r e��  ��  � ����
�� 
btns� J   o t�� ���� m   o r�� ���  T e r m i n e r��  � ����
�� 
cbtn� m   w z�� ���  T e r m i n e r� �����
�� 
dflt� m   } ��� ���  T e r m i n e r��  � ���� l  � ���������  ��  ��  ��  � ��� =   � ���� l  � ������� I  � ������
�� .corecnte****       ****� o   � ����� *0 ficheclienttrouvees ficheclientTrouvees��  ��  ��  � m   � ����� � ��� k   � ��� ��� l  � ���������  ��  ��  � ��� l  � �������  � n htell current application to my logToFileAndToConsole("Une personne porte ce nom, on r�cup�re sa fiche.")   � ��� � t e l l   c u r r e n t   a p p l i c a t i o n   t o   m y   l o g T o F i l e A n d T o C o n s o l e ( " U n e   p e r s o n n e   p o r t e   c e   n o m ,   o n   r � c u p � r e   s a   f i c h e . " )� ��� l   � �������  � 8 2 Une personne porte ce nom, on r�cup�re sa fiche.    � ��� d   U n e   p e r s o n n e   p o r t e   c e   n o m ,   o n   r � c u p � r e   s a   f i c h e .  � ��� l  � ���������  ��  ��  � ��� r   � ���� n   � ���� 4  � ����
�� 
cobj� m   � ����� � o   � ����� *0 ficheclienttrouvees ficheclientTrouvees� o      ���� 0 ficheclient ficheClient� ���� l  � ���������  ��  ��  ��  � ��� ?   � ���� l  � ������� I  � ������
�� .corecnte****       ****� o   � ����� *0 ficheclienttrouvees ficheclientTrouvees��  ��  ��  � m   � ����� � ���� k   �(�� ��� l  � ���������  ��  ��  � ��� l  � �������  � a [tell current application to my logToFileAndToConsole("Plusieurs personnes portent ce nom.")   � ��� � t e l l   c u r r e n t   a p p l i c a t i o n   t o   m y   l o g T o F i l e A n d T o C o n s o l e ( " P l u s i e u r s   p e r s o n n e s   p o r t e n t   c e   n o m . " )� ��� l   � �������  � � �
                 Plusieurs personnes portent ce nom.
                 On met les noms complets (propri�t� name) dans une liste
                 et on demande � l'utilisateur de s�lectionner la personne recherch�e.
                    � �  � 
                                   P l u s i e u r s   p e r s o n n e s   p o r t e n t   c e   n o m . 
                                   O n   m e t   l e s   n o m s   c o m p l e t s   ( p r o p r i � t �   n a m e )   d a n s   u n e   l i s t e 
                                   e t   o n   d e m a n d e   �   l ' u t i l i s a t e u r   d e   s � l e c t i o n n e r   l a   p e r s o n n e   r e c h e r c h � e . 
                                  �  l  � ���������  ��  ��    r   � � J   � �����   o      ���� 0 nomsclients nomsClients  l  � ���������  ��  ��   	
	 X   � ��� r   � � n   � � 1   � ���
�� 
pnam o   � ����� 0 unclient unClient n        ;   � � o   � ����� 0 nomsclients nomsClients�� 0 unclient unClient o   � ����� *0 ficheclienttrouvees ficheclientTrouvees
  l  � ���������  ��  ��    r   � I  � ��
�� .gtqpchltns    @   @ ns   l 
 � ����� o   � ����� 0 nomsclients nomsClients��  ��   ��
�� 
appr l 	 � ����� m   � � �    F a c t u r e��  ��   ��!"
�� 
prmp! b   � �#$# b   � �%&% b   � �'(' m   � �)) �** T I l   y   a   p l u s   d ' u n   c o n t a c t   q u i   p o r t e   l e   n o m  ( o   � ����� 0 	nomclient 	nomClient& o   � ���
�� 
ret $ l 	 � �+����+ m   � �,, �-- p C h o i s i s s e z   d a n s   l a   l i s t e   l e   c o n t a c t   q u e   v o u s   r e c h e r c h e z .��  ��  " ��./
�� 
okbt. l 	 � �0����0 m   � �11 �22  O K��  ��  / ��34
�� 
cnbt3 l 	 � �5����5 m   � �66 �77  A n n u l e r��  ��  4 ��8��
�� 
mlsl8 m   � ���
�� boovfals��   o      ���� 0 listereponse listeReponse 9:9 l ��������  ��  ��  : ;<; r  =>= n  	?@? 4  	��A
�� 
cobjA m  ���� @ o  ���� 0 listereponse listeReponse> o      ���� 0 	nomclient 	nomClient< BCB l ��DE��  D  display dialog nomClient   E �FF 0 d i s p l a y   d i a l o g   n o m C l i e n tC GHG l ��������  ��  ��  H IJI l  ��KL��  K � �
                 De la position du nom dans la liste,
                 on trouve la fiche dans la liste des clients
                    L �MM 
                                   D e   l a   p o s i t i o n   d u   n o m   d a n s   l a   l i s t e , 
                                   o n   t r o u v e   l a   f i c h e   d a n s   l a   l i s t e   d e s   c l i e n t s 
                                  J NON r  PQP n RSR I  ��T���� 0 list_position  T UVU o  ���� 0 	nomclient 	nomClientV W��W o  ���� 0 nomsclients nomsClients��  ��  S  f  Q o      ���� 0 pos  O XYX l ��Z[��  Z  display dialog pos   [ �\\ $ d i s p l a y   d i a l o g   p o sY ]^] l ����~��  �  �~  ^ _`_ r  &aba n  $cdc 4  $�}e
�} 
cobje o  "#�|�| 0 pos  d o  �{�{ *0 ficheclienttrouvees ficheclientTrouveesb o      �z�z 0 ficheclient ficheClient` f�yf l ''�x�w�v�x  �w  �v  �y  ��  �  � ghg l --�u�t�s�u  �t  �s  h iji l --�r�q�p�r  �q  �p  j klk l  --�omn�o  m - ' R�cup�ration du nom complet du client    n �oo N   R � c u p � r a t i o n   d u   n o m   c o m p l e t   d u   c l i e n t  l pqp l --�n�m�l�n  �m  �l  q rsr r  -6tut n  -0vwv 1  .0�k
�k 
pnamw o  -.�j�j 0 ficheclient ficheClientu o      �i�i 0 	nomclient 	nomClients xyx O 7Jz{z n ;I|}| I  <I�h~�g�h .0 logtofileandtoconsole logToFileAndToConsole~ �f b  <E��� m  <?�� ���   N o m   d u   c l i e n t   :  � o  ?D�e�e 0 	nomclient 	nomClient�f  �g  }  f  ;<{ m  78�d
�d misccuray ��� l KK�c�b�a�c  �b  �a  � ��� l KK�`�_�^�`  �_  �^  � ��� l  KK�]���]  � * $ R�cup�ration de la forme juridique    � ��� H   R � c u p � r a t i o n   d e   l a   f o r m e   j u r i d i q u e  � ��� Z  Kb���\�� = KR��� n  KP��� 1  LP�[
�[ 
az51� o  KL�Z�Z 0 ficheclient ficheClient� m  PQ�Y
�Y boovtrue� r  UZ��� m  UX�� ���  E n t r e p r i s e� o      �X�X  0 formejuridique formeJuridique�\  � r  ]b��� m  ]`�� ���  P a r t i c u l i e r� o      �W�W  0 formejuridique formeJuridique� ��� l cc�V�U�T�V  �U  �T  � ��� O cr��� n gq��� I  hq�S��R�S .0 logtofileandtoconsole logToFileAndToConsole� ��Q� b  hm��� m  hk�� ��� 8 F o r m e   j u r i d i q u e   d u   c l i e n t   :  � o  kl�P�P  0 formejuridique formeJuridique�Q  �R  �  f  gh� m  cd�O
�O misccura� ��� l ss�N�M�L�N  �M  �L  � ��� l ss�K�J�I�K  �J  �I  � ��� l  ss�H���H  � B < R�cup�ration de l'adresse mail � partir de la fiche client    � ��� x   R � c u p � r a t i o n   d e   l ' a d r e s s e   m a i l   �   p a r t i r   d e   l a   f i c h e   c l i e n t  � ��� l ss�G�F�E�G  �F  �E  � ��� l ss�D���D  �  repeat   � ���  r e p e a t� ��� l ss�C�B�A�C  �B  �A  � ��� r  sz��� n  sx��� 2 tx�@
�@ 
az21� o  st�?�? 0 ficheclient ficheClient� o      �>�> 0 contactemail contactEmail� ��� l {{�=�<�;�=  �<  �;  � ��� Z  {����:�� >  {��� o  {|�9�9 0 contactemail contactEmail� J  |~�8�8  � k  ���� ��� l ���7�6�5�7  �6  �5  � ��� l ���4���4  �  exit repeat   � ���  e x i t   r e p e a t� ��3� l ���2�1�0�2  �1  �0  �3  �:  � k  ���� ��� l ���/�.�-�/  �.  �-  � ��� I ���,��
�, .sysodlogaskr        TEXT� b  ����� b  ����� b  ����� b  ����� b  ����� b  ����� m  ���� ��� ^ A u c u n e   a d r e s s e   m � l e   r e n s e i g n � e   p o u r   l e   c l i e n t   "� o  ���+�+ 0 	nomclient 	nomClient� m  ���� ���  " .� l 	����*�)� o  ���(
�( 
ret �*  �)  � l 	����'�&� m  ���� ��� x V e u i l l e z   c o m p l � t e r   l a   f i c h e   c l i e n t   e t   r e l a n c e r   l e   p r o g r a m m e .�'  �&  � o  ���%
�% 
ret � l 	����$�#� m  ���� ��� " F i n   d u   p r o g r a m m e .�$  �#  � �"��
�" 
appr� l 	����!� � m  ���� ���  F a c t u r e�!  �   � ���
� 
btns� J  ���� ��� m  ���� ���  T e r m i n e r�  � �� 
� 
cbtn� m  �� �  T e r m i n e r  ��
� 
dflt m  �� �  T e r m i n e r�  � � l ������  �  �  �  �  l ������  �  �   	
	 l ����    
end repeat    �  e n d   r e p e a t
  l ������  �  �    Z  �[� ?  �� l ���� I ����
� .corecnte****       **** o  ���
�
 0 contactemail contactEmail�  �  �   m  ���	�	  k  �8  l ������  �  �    r  �� J  ����   o      �� 0 	theemails 	theEmails  !  X  ��"�#" r  ��$%$ b  ��&'& b  ��()( n  ��*+* 1  ���
� 
az18+ o  ���� 0 anemail anEmail) m  ��,, �--  :  ' n  ��./. 1  ��� 
�  
az17/ o  ������ 0 anemail anEmail% n      010  ;  ��1 o  ������ 0 	theemails 	theEmails� 0 anemail anEmail# o  ������ 0 contactemail contactEmail! 232 r  �454 n  �676 4 ��8
�� 
cobj8 m  ���� 7 l �9����9 I ���:;
�� .gtqpchltns    @   @ ns  : o  ������ 0 	theemails 	theEmails; ��<��
�� 
prmp< m  == �>> b V e u i l l e z   s � l e c t i o n n e r   l ' a d r e s s e   m a i l   �   u t i l i s e r   :��  ��  ��  5 o      ���� 0 contactemail contactEmail3 ?@? r  ABA m  CC �DD  :  B n     EFE 1  ��
�� 
txdlF 1  ��
�� 
ascr@ GHG r  (IJI n  "KLK 4  "��M
�� 
citmM m   !������L o  ���� 0 contactemail contactEmailJ o      ���� .0 adressecourrierclient adresseCourrierClientH NON r  )6PQP J  ).RR S��S m  ),TT �UU  ��  Q n     VWV 1  15��
�� 
txdlW 1  .1��
�� 
ascrO XYX l 77��������  ��  ��  Y Z[Z l 77��\]��  \ z ttell current application to my logToFileAndToConsole("Adresse mail principale du client : " & adresseCourrierClient)   ] �^^ � t e l l   c u r r e n t   a p p l i c a t i o n   t o   m y   l o g T o F i l e A n d T o C o n s o l e ( " A d r e s s e   m a i l   p r i n c i p a l e   d u   c l i e n t   :   "   &   a d r e s s e C o u r r i e r C l i e n t )[ _��_ l 77��������  ��  ��  ��   `a` =  ;Bbcb l ;@d����d I ;@��e��
�� .corecnte****       ****e o  ;<���� 0 contactemail contactEmail��  ��  ��  c m  @A���� a f��f k  EWgg hih l EE��������  ��  ��  i jkj r  EUlml n  EOnon 1  KO��
�� 
az17o n  EKpqp 4 FK��r
�� 
cobjr m  IJ���� q o  EF���� 0 contactemail contactEmailm o      ���� .0 adressecourrierclient adresseCourrierClientk sts l VV��uv��  u o itell current application to my logToFileAndToConsole("Adresse mail du client : " & adresseCourrierClient)   v �ww � t e l l   c u r r e n t   a p p l i c a t i o n   t o   m y   l o g T o F i l e A n d T o C o n s o l e ( " A d r e s s e   m a i l   d u   c l i e n t   :   "   &   a d r e s s e C o u r r i e r C l i e n t )t x��x l VV��������  ��  ��  ��  ��  �   yzy l \\��������  ��  ��  z {|{ O \o}~} n `n� I  an������� .0 logtofileandtoconsole logToFileAndToConsole� ���� b  aj��� m  ad�� ��� 2 A d r e s s e   m a i l   d u   c l i e n t   :  � o  di���� .0 adressecourrierclient adresseCourrierClient��  ��  �  f  `a~ m  \]��
�� misccura| ��� l pp��������  ��  ��  � ��� l  pp������  � = 7 R�cup�ration de l'adresse � partir de la fiche client    � ��� n   R � c u p � r a t i o n   d e   l ' a d r e s s e   �   p a r t i r   d e   l a   f i c h e   c l i e n t  � ��� l pp��������  ��  ��  � ��� Q  p����� k  s��� ��� r  s��� n  s}��� 4 x}���
�� 
cobj� m  {|���� � n  sx��� m  tx��
�� 
az27� o  st���� 0 ficheclient ficheClient� o      ���� 0 adresseclient adresseClient� ��� O  ����� r  ����� b  ����� b  ����� b  ����� b  ����� 1  ����
�� 
az28� 1  ����
�� 
spac� 1  ����
�� 
az31� 1  ����
�� 
spac� 1  ����
�� 
az29� o      ���� .0 adresseclientformatee adresseClientFormatee� o  ������ 0 adresseclient adresseClient� ��� l ����������  ��  ��  � ��� O ����� n ����� I  ��������� .0 logtofileandtoconsole logToFileAndToConsole� ���� b  ����� m  ���� ��� ( A d r e s s e   d u   c l i e n t   :  � o  ������ .0 adresseclientformatee adresseClientFormatee��  ��  �  f  ��� m  ����
�� misccura� ���� l ����������  ��  ��  ��  � R      ������
�� .ascrerr ****      � ****��  ��  � k  ���� ��� l ����������  ��  ��  � ��� I ������
�� .sysodlogaskr        TEXT� b  ����� b  ����� b  ����� b  ����� b  ����� b  ����� m  ���� ��� T A u c u n e   a d r e s s e   r e n s e i g n � e   p o u r   l e   c l i e n t   "� o  ������ 0 	nomclient 	nomClient� m  ���� ���  " .� l 	�������� o  ����
�� 
ret ��  ��  � m  ���� ��� x V e u i l l e z   c o m p l � t e r   l a   f i c h e   c l i e n t   e t   r e l a n c e r   l e   p r o g r a m m e .� o  ����
�� 
ret � l 	�������� m  ���� ��� " F i n   d u   p r o g r a m m e .��  ��  � ����
�� 
appr� l 	�������� m  ���� ���  F a c t u r e��  ��  � ����
�� 
btns� J  ���� ���� m  ���� ���  T e r m i n e r��  � ����
�� 
cbtn� m  ���� ���  T e r m i n e r� �����
�� 
dflt� m  ���� ���  T e r m i n e r��  � ���� l ����������  ��  ��  ��  � ��� l ����������  ��  ��  � ��� l ����������  ��  ��  � ��� I ��������
�� .aevtquitnull��� ��� null��  ��  � ���� l ����������  ��  ��  ��  b m   & '���                                                                                  adrb  alis    V  Macintosh HD               �Z�H+   �=oContacts.app                                                    �ShӘϻ        ����  	                Applications    �Z�s      Ә��     �=o  'Macintosh HD:Applications: Contacts.app     C o n t a c t s . a p p    M a c i n t o s h   H D  Applications/Contacts.app   / ��  _   application "Contacts"   ` ��� .   a p p l i c a t i o n   " C o n t a c t s "] ��� l ������~��  �  �~  � ��� l ���}���}  �   display dialog nomClient   � ��� 2   d i s p l a y   d i a l o g   n o m C l i e n t� ��� l ���|�{�z�|  �{  �z  � ��� l ���y �y    ) # log "Nom du client : " & nomClient    � F   l o g   " N o m   d u   c l i e n t   :   "   &   n o m C l i e n t�  l ���x�x   > 8 log "Adresse mail du client : " & adresseCourrierClient    � p   l o g   " A d r e s s e   m a i l   d u   c l i e n t   :   "   &   a d r e s s e C o u r r i e r C l i e n t 	 l ���w
�w  
 9 3 log "Adresse du client : " & adresseClientFormatee    � f   l o g   " A d r e s s e   d u   c l i e n t   :   "   &   a d r e s s e C l i e n t F o r m a t e e	  l ���v�u�t�v  �u  �t    n � I  �s�r�s "0 setstringvalue_ setStringValue_ �q o  	�p�p 0 	nomclient 	nomClient�q  �r   o  ��o�o 0 tfnomclient tfNomClient  l �n�m�l�n  �m  �l    l �k�k   , &set theText to leTexte's stringValue()    � L s e t   t h e T e x t   t o   l e T e x t e ' s   s t r i n g V a l u e ( )  l �j�j   2 ,display alert "Vous avez �crit : " & theText    �   X d i s p l a y   a l e r t   " V o u s   a v e z   � c r i t   :   "   &   t h e T e x t !"! l �i�h�g�i  �h  �g  " #$# l �f%&�f  % / )leTexte's setStringValue_("rien du tout")   & �'' R l e T e x t e ' s   s e t S t r i n g V a l u e _ ( " r i e n   d u   t o u t " )$ ()( l �e*+�e  * 8 2unAutreTexte's setStringValue_("que dalle, oui !")   + �,, d u n A u t r e T e x t e ' s   s e t S t r i n g V a l u e _ ( " q u e   d a l l e ,   o u i   ! " )) -.- l �d�c�b�d  �c  �b  . /0/ l �a�`�_�a  �`  �_  0 121 l  �^34�^  3 \ V*********************** R�cup�ration des donn�es de l'intervention *******************   4 �55 � * * * * * * * * * * * * * * * * * * * * * * *   R � c u p � r a t i o n   d e s   d o n n � e s   d e   l ' i n t e r v e n t i o n   * * * * * * * * * * * * * * * * * * *2 676 l �]�\�[�]  �\  �[  7 898 l �Z�Y�X�Z  �Y  �X  9 :;: l  �W<=�W  < + % R�cup�ration du type d'intervention    = �>> J   R � c u p � r a t i o n   d u   t y p e   d ' i n t e r v e n t i o n  ; ?@? l �V�U�T�V  �U  �T  @ ABA r  CDC c  EFE l G�S�RG n HIH I  �Q�P�O�Q 0 stringvalue stringValue�P  �O  I o  �N�N (0 tftypeintervention tfTypeIntervention�S  �R  F m  �M
�M 
ctxtD o      �L�L $0 typeintervention typeInterventionB JKJ n  .LML I  !.�KN�J�K .0 logtofileandtoconsole logToFileAndToConsoleN O�IO b  !*PQP m  !$RR �SS : L e   t y p e   d ' i n t e r v e n t i o n   e s t   :  Q o  $)�H�H $0 typeintervention typeIntervention�I  �J  M  f   !K TUT l //�G�F�E�G  �F  �E  U VWV l //�D�C�B�D  �C  �B  W XYX l  //�AZ[�A  Z 2 , R�cup�ration de la dur�e de l'intervention    [ �\\ X   R � c u p � r a t i o n   d e   l a   d u r � e   d e   l ' i n t e r v e n t i o n  Y ]^] l //�@�?�>�@  �?  �>  ^ _`_ r  /@aba c  /:cdc l /8e�=�<e n /8fgf I  48�;�:�9�; 0 stringvalue stringValue�:  �9  g o  /4�8�8 *0 tfdureeintervention tfDureeIntervention�=  �<  d m  89�7
�7 
ctxtb o      �6�6 &0 dureeintervention dureeIntervention` hih n AOjkj I  BO�5l�4�5 .0 logtofileandtoconsole logToFileAndToConsolel m�3m b  BKnon m  BEpp �qq < L a   d u r � e   d ' i n t e r v e n t i o n   e s t   :  o o  EJ�2�2 &0 dureeintervention dureeIntervention�3  �4  k  f  ABi rsr l PP�1�0�/�1  �0  �/  s tut l PP�.�-�,�.  �-  �,  u vwv l  PP�+xy�+  x 1 + R�cup�ration de la date de l'intervention    y �zz V   R � c u p � r a t i o n   d e   l a   d a t e   d e   l ' i n t e r v e n t i o n  w {|{ r  Pc}~} l P]�*�) c  P]��� n PY��� I  UY�(�'�&�( 0 dateas dateAS�'  �&  � o  PU�%�% 0 
datepicker 
datePicker� m  Y\�$
�$ 
ldt �*  �)  ~ o      �#�# $0 dateintervention dateIntervention| ��� n dt��� I  et�"��!�" .0 logtofileandtoconsole logToFileAndToConsole� �� � c  ep��� b  en��� m  eh�� ��� @ L a   d a t e   d e   l ' i n t e r v e n t i o n   e s t   :  � o  hm�� $0 dateintervention dateIntervention� m  no�
� 
ctxt�   �!  �  f  de� ��� l uu����  �  �  � ��� l  uu����  �]W
         
         -- Si on est arriv� ici, c'est que l'on a toutes les donn�es de l'intervention
         display alert "Type d'intervention : " & typeIntervention & return & �
         "Dur�e de l'intervention : " & dureeIntervention & return & �
         "Date de l'intervention : " & dateIntervention
         --return
         
            � ���� 
                   
                   - -   S i   o n   e s t   a r r i v �   i c i ,   c ' e s t   q u e   l ' o n   a   t o u t e s   l e s   d o n n � e s   d e   l ' i n t e r v e n t i o n 
                   d i s p l a y   a l e r t   " T y p e   d ' i n t e r v e n t i o n   :   "   &   t y p e I n t e r v e n t i o n   &   r e t u r n   &   � 
                   " D u r � e   d e   l ' i n t e r v e n t i o n   :   "   &   d u r e e I n t e r v e n t i o n   &   r e t u r n   &   � 
                   " D a t e   d e   l ' i n t e r v e n t i o n   :   "   &   d a t e I n t e r v e n t i o n 
                   - - r e t u r n 
                   
                  � ��� l uu����  �  �  � ��� l uu����  �  �  � ��� l  uu����  � _ Y**************************** Cr�ation de la facture dans Excel **************************   � ��� � * * * * * * * * * * * * * * * * * * * * * * * * * * * *   C r � a t i o n   d e   l a   f a c t u r e   d a n s   E x c e l   * * * * * * * * * * * * * * * * * * * * * * * * * *� ��� l uu����  �  �  � ��� l uu����  �  �  � ��� l  uu����  �71 Cr�e un fichier excel � partir d'un mod�le, le modifie en fonction des donn�es entr�es et l'enregistre en pdf sur le bureau.
         Le fichier aura le nom: " Facture_00XXX.pdf" avec une espace devant ajout�e automatiquement par excel qui concat�ne le nom du fichier avec le nom de la feuille de calcul.   � ���b   C r � e   u n   f i c h i e r   e x c e l   �   p a r t i r   d ' u n   m o d � l e ,   l e   m o d i f i e   e n   f o n c t i o n   d e s   d o n n � e s   e n t r � e s   e t   l ' e n r e g i s t r e   e n   p d f   s u r   l e   b u r e a u . 
                   L e   f i c h i e r   a u r a   l e   n o m :   "   F a c t u r e _ 0 0 X X X . p d f "   a v e c   u n e   e s p a c e   d e v a n t   a j o u t � e   a u t o m a t i q u e m e n t   p a r   e x c e l   q u i   c o n c a t � n e   l e   n o m   d u   f i c h i e r   a v e c   l e   n o m   d e   l a   f e u i l l e   d e   c a l c u l .� ��� l u����� O  u���� k  {��� ��� l {{��
�	�  �
  �	  � ��� l {{����  � [ Utell current application to my logToFileAndToConsole("Appel de l'application Excel.")   � ��� � t e l l   c u r r e n t   a p p l i c a t i o n   t o   m y   l o g T o F i l e A n d T o C o n s o l e ( " A p p e l   d e   l ' a p p l i c a t i o n   E x c e l . " )� ��� l {{����  �  �  � ��� I {����
� .aevtodocnull  �    alis� c  {���� o  {��� B0 cheminfichiermodelefactureexcel cheminFichierModeleFactureExcel� m  ���
� 
psxf�  � ��� l ��� �����   ��  ��  � ��� l �G���� O  �G��� k  �F�� ��� l ����������  ��  ��  � ��� l ��������  � > 8 Mise � jour du num�ro de facture (dans les 2 cellules!)   � ��� p   M i s e   �   j o u r   d u   n u m � r o   d e   f a c t u r e   ( d a n s   l e s   2   c e l l u l e s ! )� ��� r  ����� o  ������ 0 numerofacture numeroFacture� n      ��� 1  ����
�� 
DPVu� 4  �����
�� 
ccel� m  ���� ���  D 4� ��� r  ����� o  ������ 0 numerofacture numeroFacture� n      ��� 1  ����
�� 
DPVu� 4  �����
�� 
ccel� m  ���� ���  B 1 1� ��� l ����������  ��  ��  � ��� l ��������  � , & Mise � jour de la date d'intervention   � ��� L   M i s e   �   j o u r   d e   l a   d a t e   d ' i n t e r v e n t i o n� ��� r  ����� n  ����� 1  ����
�� 
shdt� o  ������ $0 dateintervention dateIntervention� n      ��� 1  ����
�� 
DPVu� 4  �����
�� 
ccel� m  ���� ���  B 1 2� ��� l ����������  ��  ��  � ��� l ��������  � #  Mise � jour du nom du client   � ��� :   M i s e   �   j o u r   d u   n o m   d u   c l i e n t� ��� r  ����� o  ������ 0 	nomclient 	nomClient� n      � � 1  ����
�� 
DPVu  4  ����
�� 
ccel m  �� �  B 1 4�  l ����������  ��  ��    l ����	��   ) # Mise � jour de l'adresse du client   	 �

 F   M i s e   �   j o u r   d e   l ' a d r e s s e   d u   c l i e n t  r  �� o  ������ .0 adresseclientformatee adresseClientFormatee n       1  ����
�� 
DPVu 4  ����
�� 
ccel m  �� �  B 1 5  l ����������  ��  ��    l ������   2 , Mise � jour de la forme juridique du client    � X   M i s e   �   j o u r   d e   l a   f o r m e   j u r i d i q u e   d u   c l i e n t  r  � o  ������  0 formejuridique formeJuridique n        1   ��
�� 
DPVu  4  � ��!
�� 
ccel! m  ��"" �##  B 1 6 $%$ l ��������  ��  ��  % &'& l ��()��  ( ) # Mise � jour du type d'intervention   ) �** F   M i s e   �   j o u r   d u   t y p e   d ' i n t e r v e n t i o n' +,+ r  -.- o  ���� $0 typeintervention typeIntervention. n      /0/ 1  ��
�� 
DPVu0 4  ��1
�� 
ccel1 m  22 �33  A 2 1, 454 l ��������  ��  ��  5 676 l ��89��  8 0 * Mise � jour de la dur�e de l'intervention   9 �:: T   M i s e   �   j o u r   d e   l a   d u r � e   d e   l ' i n t e r v e n t i o n7 ;<; r  +=>= o  ���� &0 dureeintervention dureeIntervention> n      ?@? 1  &*��
�� 
DPVu@ 4  &��A
�� 
ccelA m  "%BB �CC  C 2 1< DED l ,,��������  ��  ��  E FGF l ,,��HI��  H , & Mise � jour du co�t de l'intervention   I �JJ L   M i s e   �   j o u r   d u   c o � t   d e   l ' i n t e r v e n t i o nG KLK r  ,DMNM ]  ,7OPO o  ,1���� &0 dureeintervention dureeInterventionP o  16���� 0 prixhoraire prixHoraireN n      QRQ 1  ?C��
�� 
DPVuR 4  7?��S
�� 
ccelS m  ;>TT �UU  D 2 1L V��V l EE��������  ��  ��  ��  � n  ��WXW 4  ����Y
�� 
XwSHY m  ��ZZ �[[  S h e e t 1X 1  ����
�� 
1172� + %worksheet "Sheet1" of active workbook   � �\\ J w o r k s h e e t   " S h e e t 1 "   o f   a c t i v e   w o r k b o o k� ]^] l HH��������  ��  ��  ^ _`_ l HH��ab��  a A ; Une fois la facture remplie, on l'enregistre au format PDF   b �cc v   U n e   f o i s   l a   f a c t u r e   r e m p l i e ,   o n   l ' e n r e g i s t r e   a u   f o r m a t   P D F` ded l HH��fg��  f ? 9 On donne un nom � la feuille (ex : "Facture_00" & "234")   g �hh r   O n   d o n n e   u n   n o m   �   l a   f e u i l l e   ( e x   :   " F a c t u r e _ 0 0 "   &   " 2 3 4 " )e iji r  H]klk l HUm����m b  HUnon o  HM���� 20 prefixnomfichierfacture prefixNomFichierFactureo l MTp����p c  MTqrq o  MR���� 0 numerofacture numeroFacturer m  RS��
�� 
ctxt��  ��  ��  ��  l n      sts 1  Z\��
�� 
pnamt 1  UZ��
�� 
1107j uvu l ^^��wx��  w P J On enregistre au format PDF en indiquant le chemin du fichier de la forme   x �yy �   O n   e n r e g i s t r e   a u   f o r m a t   P D F   e n   i n d i q u a n t   l e   c h e m i n   d u   f i c h i e r   d e   l a   f o r m ev z{z l ^^��|}��  | L F chemin + extension (ex : "Macintosh HD:Users:bruno:Desktop:" & ".pdf"   } �~~ �   c h e m i n   +   e x t e n s i o n   ( e x   :   " M a c i n t o s h   H D : U s e r s : b r u n o : D e s k t o p : "   &   " . p d f "{ � l ^^������  � ] W le fichier produit aura la forme : chemin + une espace + nom de la feuille + extension   � ��� �   l e   f i c h i e r   p r o d u i t   a u r a   l a   f o r m e   :   c h e m i n   +   u n e   e s p a c e   +   n o m   d e   l a   f e u i l l e   +   e x t e n s i o n� ��� l ^^������  � B < (ex : "Macintosh HD:Users:bruno:Desktop: Facture_00234.pdf"   � ��� x   ( e x   :   " M a c i n t o s h   H D : U s e r s : b r u n o : D e s k t o p :   F a c t u r e _ 0 0 2 3 4 . p d f "� ��� l ^^������  � G A Il faudra donc renommer le fichier pour enlever l'espace inutile   � ��� �   I l   f a u d r a   d o n c   r e n o m m e r   l e   f i c h i e r   p o u r   e n l e v e r   l ' e s p a c e   i n u t i l e� ��� O ^���� I f������
�� .smXL1659null���   7 X128��  � ����
�� 
5016� l ju������ b  ju��� o  jo���� 40 chemindossiertempfacture cheminDossierTempFacture� o  ot���� 20 extensionfichierfacture extensionFichierFacture��  ��  � �����
�� 
1813� m  x{��
�� e188� .��  � 1  ^c��
�� 
1107� ��� l ����������  ��  ��  � ��� l ��������  � C = On n'enregistre pas les modifications dans le fichier mod�le   � ��� z   O n   n ' e n r e g i s t r e   p a s   l e s   m o d i f i c a t i o n s   d a n s   l e   f i c h i e r   m o d � l e� ��� I �������
�� .aevtquitnull��� ��� null��  � �����
�� 
savo� m  ����
�� boovfals��  � ���� l ����������  ��  ��  ��  � m  ux��
                                                                                  XCEL  alis    �  Macintosh HD               �Z�H+   x�0Microsoft Excel.app                                             y!�Ț�U        ����  	                Microsoft Office 2011     �Z�s      Ț�5     x�0 �=o  EMacintosh HD:Applications: Microsoft Office 2011: Microsoft Excel.app   (  M i c r o s o f t   E x c e l . a p p    M a c i n t o s h   H D  6Applications/Microsoft Office 2011/Microsoft Excel.app  / ��  � # application "Microsoft Excel"   � ��� : a p p l i c a t i o n   " M i c r o s o f t   E x c e l "� ��� l ����������  ��  ��  � ��� l ����������  ��  ��  � ��� l ����������  ��  ��  � ��� l  ��������  � ` Z************************* Renommage du fichier facture avec Finder ***********************   � ��� � * * * * * * * * * * * * * * * * * * * * * * * * *   R e n o m m a g e   d u   f i c h i e r   f a c t u r e   a v e c   F i n d e r   * * * * * * * * * * * * * * * * * * * * * * *� ��� l ����������  ��  ��  � ��� l  ��������  � � �
         On renome le fichier facture correctement (Impossible de le faire depuis Excel)
         et on le d�place dans le dossier des factures.
            � ���6 
                   O n   r e n o m e   l e   f i c h i e r   f a c t u r e   c o r r e c t e m e n t   ( I m p o s s i b l e   d e   l e   f a i r e   d e p u i s   E x c e l ) 
                   e t   o n   l e   d � p l a c e   d a n s   l e   d o s s i e r   d e s   f a c t u r e s . 
                  � ��� O  ����� k  ���� ��� l ���������  ��  �  � ��� l ���~���~  � \ Vtell current application to my logToFileAndToConsole("Appel de l'application Finder.")   � ��� � t e l l   c u r r e n t   a p p l i c a t i o n   t o   m y   l o g T o F i l e A n d T o C o n s o l e ( " A p p e l   d e   l ' a p p l i c a t i o n   F i n d e r . " )� ��� l ���}�|�{�}  �|  �{  � ��� r  ����� c  ����� o  ���z�z 40 cheminfichiertempfacture cheminFichierTempFacture� m  ���y
�y 
alis� o      �x�x 20 fichiertempfacturealias fichierTempFactureAlias� ��� l ���w�v�u�w  �v  �u  � ��� l ���t���t  � 9 3 On renome le fichier pour enlever l'espace inutile   � ��� f   O n   r e n o m e   l e   f i c h i e r   p o u r   e n l e v e r   l ' e s p a c e   i n u t i l e� ��� r  ����� o  ���s�s &0 nomfichierfacture nomFichierFacture� l     ��r�q� n      ��� 1  ���p
�p 
pnam� o  ���o�o 20 fichiertempfacturealias fichierTempFactureAlias�r  �q  � ��� l ���n�m�l�n  �m  �l  � ��� r  ����� c  ����� o  ���k�k .0 chemindossierfactures cheminDossierFactures� m  ���j
�j 
alis� o      �i�i ,0 dossierfacturesalias dossierFacturesAlias� ��� l ���h�g�f�h  �g  �f  � ��� Q  ������ k  ���� ��� l ���e���e  � T N Si on d�place un fichier, il faut mettre � jour la variable qui pointe dessus   � ��� �   S i   o n   d � p l a c e   u n   f i c h i e r ,   i l   f a u t   m e t t r e   �   j o u r   l a   v a r i a b l e   q u i   p o i n t e   d e s s u s� ��� l ���d���d  � + % car sinon on ne peut plus l'utiliser   � ��� J   c a r   s i n o n   o n   n e   p e u t   p l u s   l ' u t i l i s e r� ��� r  ����� I ���c��
�c .coremovenull���     obj � o  ���b�b 20 fichiertempfacturealias fichierTempFactureAlias� �a��`
�a 
insh� o  ���_�_ ,0 dossierfacturesalias dossierFacturesAlias�`  � o      �^�^ 0 	mynewfile 	myNewFile� ��� l ���]�\�[�]  �\  �[  � ��� l ���Z �Z    � �tell current application to my logToFileAndToConsole("Copie de " & cheminFichierTempFacture & " vers " & cheminDossierFactures & ".")    �
 t e l l   c u r r e n t   a p p l i c a t i o n   t o   m y   l o g T o F i l e A n d T o C o n s o l e ( " C o p i e   d e   "   &   c h e m i n F i c h i e r T e m p F a c t u r e   &   "   v e r s   "   &   c h e m i n D o s s i e r F a c t u r e s   &   " . " )� �Y l ���X�W�V�X  �W  �V  �Y  � R      �U�T�S
�U .ascrerr ****      � ****�T  �S  � I ���R�Q
�R .sysodisAaleR        TEXT b  �� b  �� b  ��	
	 b  �� b  �� m  �� � B I m p o s s i b l e   d e   c o p i e r   l e   f i c h i e r :   l 	���P�O o  ���N�N 40 cheminfichiertempfacture cheminFichierTempFacture�P  �O   l 
���M�L o  ���K
�K 
ret �M  �L  
 m  �� � j U n   f i c h i e r   p o r t a n t   l e   m � m e   n o m   e x i s t e   p e u t - � t r e   d � j � . l 
���J�I o  ���H
�H 
ret �J  �I   m  �� � B O p � r a t i o n   d e   d � p l a c e m e n t   a n n u l � e .�Q  � �G l ���F�E�D�F  �E  �D  �G  � m  ���                                                                                  MACS  alis    t  Macintosh HD               �Z�H+   �=P
Finder.app                                                      ��W��[        ����  	                CoreServices    �Z�s      ��;     �=P �=O �=N  6Macintosh HD:System: Library: CoreServices: Finder.app   
 F i n d e r . a p p    M a c i n t o s h   H D  &System/Library/CoreServices/Finder.app  / ��  �  l ���C�B�A�C  �B  �A    l ���@�?�>�@  �?  �>    l ���=�<�;�=  �<  �;    !  l ���:"#�:  " Y S Transformation de la date dans le format : nomDuJour Num�roDuJour NomDuMois Ann�e.   # �$$ �   T r a n s f o r m a t i o n   d e   l a   d a t e   d a n s   l e   f o r m a t   :   n o m D u J o u r   N u m � r o D u J o u r   N o m D u M o i s   A n n � e .! %&% l ���9'(�9  ' ( " Exemple : lundi 29 septembre 2015   ( �)) D   E x e m p l e   :   l u n d i   2 9   s e p t e m b r e   2 0 1 5& *+* l ���8,-�8  , < 6 Pour utilisation dans le message � envoyer au client.   - �.. l   P o u r   u t i l i s a t i o n   d a n s   l e   m e s s a g e   �   e n v o y e r   a u   c l i e n t .+ /0/ r  �121 n ��343 I  ���75�6�7 0 getdate getDate5 6�56 o  ���4�4 $0 dateintervention dateIntervention�5  �6  4  f  ��2 o      �3�3 80 dateinterventionformatlong dateInterventionFormatLong0 787 l �29:�2  9 g amy logToFileAndToConsole("Date de l'intervention au format long : " & dateInterventionFormatLong)   : �;; � m y   l o g T o F i l e A n d T o C o n s o l e ( " D a t e   d e   l ' i n t e r v e n t i o n   a u   f o r m a t   l o n g   :   "   &   d a t e I n t e r v e n t i o n F o r m a t L o n g )8 <=< l �1�0�/�1  �0  �/  = >?> l �.�-�,�.  �-  �,  ? @A@ l �+�*�)�+  �*  �)  A BCB l  �(DE�(  D [ U**************************** Envoi de la facture avec Mail **************************   E �FF � * * * * * * * * * * * * * * * * * * * * * * * * * * * *   E n v o i   d e   l a   f a c t u r e   a v e c   M a i l   * * * * * * * * * * * * * * * * * * * * * * * * * *C GHG l �'�&�%�'  �&  �%  H IJI l  �$KL�$  K��
         Cr�e le message contenant la facture � envoyer au client.
         Impossible d'ajouter la signature avant la pi�ce jointe,
         sinon elle est supprim�e ou consid�r�e comme du texte
         et non plus comme une signature (elle est m�me soulign�e).
         
         Pour r�pondre � ce probl�me, j'ajoute la pi�ce jointe,
         j'attends 5 secondes et j'ajoute la signature � la fin
         (via l'interface graphique).
            L �MM� 
                   C r � e   l e   m e s s a g e   c o n t e n a n t   l a   f a c t u r e   �   e n v o y e r   a u   c l i e n t . 
                   I m p o s s i b l e   d ' a j o u t e r   l a   s i g n a t u r e   a v a n t   l a   p i � c e   j o i n t e , 
                   s i n o n   e l l e   e s t   s u p p r i m � e   o u   c o n s i d � r � e   c o m m e   d u   t e x t e 
                   e t   n o n   p l u s   c o m m e   u n e   s i g n a t u r e   ( e l l e   e s t   m � m e   s o u l i g n � e ) . 
                   
                   P o u r   r � p o n d r e   �   c e   p r o b l � m e ,   j ' a j o u t e   l a   p i � c e   j o i n t e , 
                   j ' a t t e n d s   5   s e c o n d e s   e t   j ' a j o u t e   l a   s i g n a t u r e   �   l a   f i n 
                   ( v i a   l ' i n t e r f a c e   g r a p h i q u e ) . 
                  J NON O  �PQP k  �RR STS l �#�"�!�#  �"  �!  T UVU I � ��
�  .miscactvnull��� ��� null�  �  V WXW l ����  �  �  X YZY l �[\�  [ O Itell current application to my logToFileAndToConsole("Affichage de mail")   \ �]] � t e l l   c u r r e n t   a p p l i c a t i o n   t o   m y   l o g T o F i l e A n d T o C o n s o l e ( " A f f i c h a g e   d e   m a i l " )Z ^_^ l ����  �  �  _ `a` r  bcb b  ded b  fgf o  �� "0 contenumessage1 contenuMessage1g o  �� 80 dateinterventionformatlong dateInterventionFormatLonge o  �� "0 contenumessage2 contenuMessage2c o      ��  0 contenumessage contenuMessagea hih r   Rjkj I  N��l
� .corecrel****      � null�  l �mn
� 
koclm m  $'�
� 
bcken �o�
� 
prdto K  *Hpp �qr
� 
sndrq o  -2�� (0 monadressecourrier monAdresseCourrierr �
st
�
 
subjs o  5:�	�	 0 monsujet monSujett �uv
� 
pvisu m  =>�
� boovtruev �w�
� 
ctntw o  AD��  0 contenumessage contenuMessage�  �  k o      ��  0 nouveaumessage nouveauMessagei xyx l SS��� �  �  �   y z{z O  S�|}| k  Y�~~ � l YY��������  ��  ��  � ��� I Y|�����
�� .corecrel****      � null��  � ����
�� 
kocl� m  ]`��
�� 
rcpt� ����
�� 
insh� n  ci���  ;  hi� 2 ch��
�� 
trcp� �����
�� 
prdt� K  lv�� �����
�� 
radd� o  ot���� .0 adressecourrierclient adresseCourrierClient��  ��  � ��� l }}��������  ��  ��  � ��� l }����� I }������
�� .corecrel****      � null��  � ����
�� 
kocl� m  ����
�� 
atts� �����
�� 
prdt� K  ���� �����
�� 
atfn� l �������� n  ����� 1  ����
�� 
psxp� o  ������ (0 cheminfacturefinal cheminFactureFinal��  ��  ��  ��  � !  ins�r� � la fin du message   � ��� 6   i n s � r �   �   l a   f i n   d u   m e s s a g e� ��� I �������
�� .sysodelanull��� ��� nmbr� m  ������ ��  � ���� l ����������  ��  ��  ��  } o  SV����  0 nouveaumessage nouveauMessage{ ��� l ����������  ��  ��  � ��� I ��������
�� .miscactvnull��� ��� null��  ��  � ��� l ����������  ��  ��  � ��� l  ��������  � - ' Ins�re la signature par GUI Scripting    � ��� N   I n s � r e   l a   s i g n a t u r e   p a r   G U I   S c r i p t i n g  � ���� O  ����� O  ����� k  ���� ��� l ��������  �  	delay 1.3   � ���  d e l a y   1 . 3� ��� O  ����� k  ���� ��� I �������
�� .prcsclicnull��� ��� uiel� 4  �����
�� 
popB� m  ������ ��  � ���� I �������
�� .prcsclicnull��� ��� uiel� n  ����� 4  �����
�� 
menI� m  ������ � n  ����� 4  �����
�� 
menE� m  ������ � 4  �����
�� 
popB� m  ������ ��  ��  � 4  �����
�� 
cwin� m  ������ � ���� l ����������  ��  ��  ��  � 4  �����
�� 
prcs� m  ���� ���  M a i l� m  �����                                                                                  sevs  alis    �  Macintosh HD               �Z�H+   �=PSystem Events.app                                               �9���)Q        ����  	                CoreServices    �Z�s      ��1     �=P �=O �=N  =Macintosh HD:System: Library: CoreServices: System Events.app   $  S y s t e m   E v e n t s . a p p    M a c i n t o s h   H D  -System/Library/CoreServices/System Events.app   / ��  ��  Q m  ���                                                                                  emal  alis    F  Macintosh HD               �Z�H+   �=oMail.app                                                        �z���        ����  	                Applications    �Z�s      ��j     �=o  #Macintosh HD:Applications: Mail.app     M a i l . a p p    M a c i n t o s h   H D  Applications/Mail.app   / ��  O ��� l ����������  ��  ��  � ��� l ����������  ��  ��  � ��� l ��������  � ' !display dialog "Fin du programme"   � ��� B d i s p l a y   d i a l o g   " F i n   d u   p r o g r a m m e "� ��� l ��������  �   create an alert   � ���     c r e a t e   a n   a l e r t� ��� l ��������  � | vset theAlert to current application's NSAlert's makeAlert:"An alert" buttons:{"Cancel", "OK"} |text|:(theDate as text)   � ��� � s e t   t h e A l e r t   t o   c u r r e n t   a p p l i c a t i o n ' s   N S A l e r t ' s   m a k e A l e r t : " A n   a l e r t "   b u t t o n s : { " C a n c e l " ,   " O K " }   | t e x t | : ( t h e D a t e   a s   t e x t )� ��� l ��������  � : 4theAlert's showModal() -- returns name of the button   � ��� h t h e A l e r t ' s   s h o w M o d a l ( )   - -   r e t u r n s   n a m e   o f   t h e   b u t t o n� ��� l ��������  �  
log result   � ���  l o g   r e s u l t� ��� l ����������  ��  ��  � ��� l ����������  ��  ��  � ��� l  ��������  � V P**************************** MAJ du num�ro de facture **************************   � ��� � * * * * * * * * * * * * * * * * * * * * * * * * * * * *   M A J   d u   n u m � r o   d e   f a c t u r e   * * * * * * * * * * * * * * * * * * * * * * * * * *� ��� l ����������  ��  ��  � ��� n ����� I  ��������� .0 logtofileandtoconsole logToFileAndToConsole� ���� m  ���� �	 	  J - - - - >   M A J   d e s   v a r i a b l e s   d e   l a   f a c t u r e��  ��  �  f  ��� 			 l ����������  ��  ��  	 			 l  ����		��  	 + % Incr�mentation du num�ro de facture    	 �		 J   I n c r � m e n t a t i o n   d u   n u m � r o   d e   f a c t u r e  	 				 l ����������  ��  ��  		 	
		
 r  �			 [  ��			 o  ������ 0 numerofacture numeroFacture	 m  ������ 	 o      ���� 0 numerofacture numeroFacture	 			 n 			 I  ��	���� .0 logtofileandtoconsole logToFileAndToConsole	 	��	 b  			 m  		 �		  N u m � r o   :  	 o  ���� 0 numerofacture numeroFacture��  ��  	  f  	 			 l !				 n !		 	 I  !��	!���� "0 setstringvalue_ setStringValue_	! 	"��	" o  ���� 0 numerofacture numeroFacture��  ��  	  o  ���� &0 tfnumerodefacture tfNumeroDeFacture	   MAJ de l'interface   	 �	#	# &   M A J   d e   l ' i n t e r f a c e	 	$	%	$ l ""��������  ��  ��  	% 	&	'	& l ""��������  ��  ��  	' 	(	)	( l  ""��	*	+��  	* * $ Cr�ation du nom du fichier facture    	+ �	,	, H   C r � a t i o n   d u   n o m   d u   f i c h i e r   f a c t u r e  	) 	-	.	- l ""��������  ��  ��  	. 	/	0	/ r  ";	1	2	1 b  "5	3	4	3 b  "/	5	6	5 o  "'���� 20 prefixnomfichierfacture prefixNomFichierFacture	6 l '.	7����	7 c  '.	8	9	8 o  ',���� 0 numerofacture numeroFacture	9 m  ,-��
�� 
ctxt��  ��  	4 o  /4���� 20 extensionfichierfacture extensionFichierFacture	2 o      �� &0 nomfichierfacture nomFichierFacture	0 	:	;	: n <J	<	=	< I  =J�~	>�}�~ .0 logtofileandtoconsole logToFileAndToConsole	> 	?�|	? b  =F	@	A	@ m  =@	B	B �	C	C " N o m   d u   f i c h i e r   :  	A o  @E�{�{ &0 nomfichierfacture nomFichierFacture�|  �}  	=  f  <=	; 	D	E	D l KK�z�y�x�z  �y  �x  	E 	F	G	F l KK�w�v�u�w  �v  �u  	G 	H	I	H l  KK�t	J	K�t  	J C = Chemin vers le dossier temporaire de stockage de la facture    	K �	L	L z   C h e m i n   v e r s   l e   d o s s i e r   t e m p o r a i r e   d e   s t o c k a g e   d e   l a   f a c t u r e  	I 	M	N	M l KK�s�r�q�s  �r  �q  	N 	O	P	O r  K`	Q	R	Q b  KZ	S	T	S b  KT	U	V	U o  KP�p�p 40 chemindossiertempfacture cheminDossierTempFacture	V 1  PS�o
�o 
spac	T o  TY�n�n &0 nomfichierfacture nomFichierFacture	R o      �m�m 40 cheminfichiertempfacture cheminFichierTempFacture	P 	W	X	W n au	Y	Z	Y I  bu�l	[�k�l .0 logtofileandtoconsole logToFileAndToConsole	[ 	\�j	\ c  bq	]	^	] b  bo	_	`	_ m  be	a	a �	b	b ( C h e m i n   t e m p o r a i r e   :  	` l en	c�i�h	c n  en	d	e	d 1  jn�g
�g 
psxp	e o  ej�f�f 40 cheminfichiertempfacture cheminFichierTempFacture�i  �h  	^ m  op�e
�e 
ctxt�j  �k  	Z  f  ab	X 	f	g	f l vv�d�c�b�d  �c  �b  	g 	h	i	h l vv�a�`�_�a  �`  �_  	i 	j	k	j l  vv�^	l	m�^  	l > 8 Chemin vers le dossier final de stockage de la facture    	m �	n	n p   C h e m i n   v e r s   l e   d o s s i e r   f i n a l   d e   s t o c k a g e   d e   l a   f a c t u r e  	k 	o	p	o l vv�]�\�[�]  �\  �[  	p 	q	r	q r  v�	s	t	s b  v�	u	v	u o  v{�Z�Z .0 chemindossierfactures cheminDossierFactures	v o  {��Y�Y &0 nomfichierfacture nomFichierFacture	t o      �X�X (0 cheminfacturefinal cheminFactureFinal	r 	w	x	w n ��	y	z	y I  ���W	{�V�W .0 logtofileandtoconsole logToFileAndToConsole	{ 	|�U	| c  ��	}	~	} b  ��		�	 m  ��	�	� �	�	�  C h e m i n   f i n a l   :  	� l ��	��T�S	� n  ��	�	�	� 1  ���R
�R 
psxp	� o  ���Q�Q (0 cheminfacturefinal cheminFactureFinal�T  �S  	~ m  ���P
�P 
ctxt�U  �V  	z  f  ��	x 	�	�	� l ���O�N�M�O  �N  �M  	� 	�	�	� l ���L�K�J�L  �K  �J  	� 	�	�	� n ��	�	�	� I  ���I	��H�I .0 logtofileandtoconsole logToFileAndToConsole	� 	��G	� m  ��	�	� �	�	� J < - - - -   M A J   d e s   v a r i a b l e s   d e   l a   f a c t u r e�G  �H  	�  f  ��	� 	�	�	� l ���F�E�D�F  �E  �D  	� 	�	�	� n ��	�	�	� I  ���C	��B�C .0 logtofileandtoconsole logToFileAndToConsole	� 	��A	� m  ��	�	� �	�	� . < - - - -   N o u v e l l e   f a c t u r e 
�A  �B  	�  f  ��	� 	��@	� l ���?�>�=�?  �>  �=  �@  % 	�	�	� l     �<�;�:�<  �;  �:  	� 	�	�	� l     �9�8�7�9  �8  �7  	� 	�	�	� l     �6�5�4�6  �5  �4  	� 	�	�	� l     �3	�	��3  	� P J--------------------------------------------------------------------------   	� �	�	� � - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -	� 	�	�	� l     �2�1�0�2  �1  �0  	� 	�	�	� l     �/	�	��/  	� 2 ,                              showAlertModal   	� �	�	� X                                                             s h o w A l e r t M o d a l	� 	�	�	� l     �.�-�,�.  �-  �,  	� 	�	�	� l     �+	�	��+  	� P J--------------------------------------------------------------------------   	� �	�	� � - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -	� 	�	�	� l     �*�)�(�*  �)  �(  	� 	�	�	� l     �'	�	��'  	�   shows modal alert   	� �	�	� $   s h o w s   m o d a l   a l e r t	� 	�	�	� i   � �	�	�	� I      �&	��%�& "0 showalertmodal_ showAlertModal_	� 	��$	� o      �#�# 
0 sender  �$  �%  	� k     	�	� 	�	�	� l     �"	�	��"  	�   create an alert   	� �	�	�     c r e a t e   a n   a l e r t	� 	�	�	� r     	�	�	� n    	�	�	� I    �!	�� �! 20 makealert_buttons_text_ makeAlert_buttons_text_	� 	�	�	� m    	�	� �	�	�  A n   a l e r t	� 	�	�	� J    	�	� 	�	�	� m    	�	� �	�	�  C a n c e l	� 	��	� m    	�	� �	�	�  O K�  	� 	��	� m    		�	� �	�	� & F u r t h e r   e x p l a n a t i o n�  �   	� n    	�	�	� o    �� 0 nsalert NSAlert	� m     �
� misccura	� o      �� 0 thealert theAlert	� 	�	�	� l   	�	�	�	� n   	�	�	� I    ���� 0 	showmodal 	showModal�  �  	� o    �� 0 thealert theAlert	� !  returns name of the button   	� �	�	� 6   r e t u r n s   n a m e   o f   t h e   b u t t o n	� 	��	� I   �	��
� .ascrcmnt****      � ****	� 1    �
� 
rslt�  �  	� 	�	�	� l     ����  �  �  	� 	�	�	� l     ����  �  �  	� 	�	�	� l     �	�	��  	� P J--------------------------------------------------------------------------   	� �	�	� � - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -	� 	�	�	� l     ��
�	�  �
  �	  	� 	�	�	� l     �	�	��  	� 1 +                              datePickerGet   	� �	�	� V                                                             d a t e P i c k e r G e t	� 	�	�	� l     ����  �  �  	� 	�
 	� l     �

�  
 P J--------------------------------------------------------------------------   
 �

 � - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
  


 l     ����  �  �  
 


 i   � �

	
 I      � 

���   0 datepickerget_ datePickerGet_

 
��
 o      ���� 
0 sender  ��  ��  
	 k     #

 


 l     ��������  ��  ��  
 


 r     


 l    
����
 c     


 n    	


 I    	�������� 0 dateas dateAS��  ��  
 o     ���� 0 
datepicker 
datePicker
 m   	 
��
�� 
ldt ��  ��  
 o      ���� 0 thedate theDate
 


 I   ��
��
�� .ascrcmnt****      � ****
 c    


 o    ���� 0 thedate theDate
 m    ��
�� 
ctxt��  
 


 I   ��
��
�� .ascrcmnt****      � ****
 m    
 
  �
!
!  # #��  
 
"
#
" I   !��
$��
�� .ascrcmnt****      � ****
$ m    
%
% �
&
&  # #��  
# 
'��
' l  " "��������  ��  ��  ��  
 
(
)
( l     ��������  ��  ��  
) 
*
+
* l     ��������  ��  ��  
+ 
,
-
, l     ��������  ��  ��  
- 
.
/
. l     ��
0
1��  
0 P J--------------------------------------------------------------------------   
1 �
2
2 � - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
/ 
3
4
3 l     ��������  ��  ��  
4 
5
6
5 l     ��
7
8��  
7 1 +                              list_position   
8 �
9
9 V                                                             l i s t _ p o s i t i o n
6 
:
;
: l     ��������  ��  ��  
; 
<
=
< l     ��
>
?��  
> P J--------------------------------------------------------------------------   
? �
@
@ � - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
= 
A
B
A l     ��������  ��  ��  
B 
C
D
C l      ��
E
F��  
E  �
     list_position : renvoie la position d'un �l�ment dans une liste.
     this_item : �l�ment � rechercher.
     this_list : liste dans laquelle on recherche.
     r�sultat : 0 si non trouv�, la position de l'�l�ment
     dans la liste sinon.
        
F �
G
G� 
           l i s t _ p o s i t i o n   :   r e n v o i e   l a   p o s i t i o n   d ' u n   � l � m e n t   d a n s   u n e   l i s t e . 
           t h i s _ i t e m   :   � l � m e n t   �   r e c h e r c h e r . 
           t h i s _ l i s t   :   l i s t e   d a n s   l a q u e l l e   o n   r e c h e r c h e . 
           r � s u l t a t   :   0   s i   n o n   t r o u v � ,   l a   p o s i t i o n   d e   l ' � l � m e n t 
           d a n s   l a   l i s t e   s i n o n . 
          
D 
H
I
H i   � �
J
K
J I      ��
L���� 0 list_position  
L 
M
N
M o      ���� 0 	this_item  
N 
O��
O o      ���� 0 	this_list  ��  ��  
K k     .
P
P 
Q
R
Q l     ��������  ��  ��  
R 
S
T
S q      
U
U ������ 0 position  ��  
T 
V
W
V r     
X
Y
X m     ����  
Y o      ���� 0 position  
W 
Z
[
Z l   ��������  ��  ��  
[ 
\
]
\ Y    )
^��
_
`��
^ k    $
a
a 
b
c
b l   ��������  ��  ��  
c 
d
e
d Z    "
f
g����
f =   
h
i
h n    
j
k
j 4    ��
l
�� 
cobj
l o    ���� 0 i  
k o    ���� 0 	this_list  
i o    ���� 0 	this_item  
g r    
m
n
m o    ���� 0 i  
n o      ���� 0 position  ��  ��  
e 
o��
o l  # #��������  ��  ��  ��  �� 0 i  
_ m    ���� 
` l   
p����
p I   ��
q��
�� .corecnte****       ****
q o    	���� 0 	this_list  ��  ��  ��  ��  
] 
r
s
r l  * *��������  ��  ��  
s 
t
u
t L   * ,
v
v o   * +���� 0 position  
u 
w��
w l  - -��������  ��  ��  ��  
I 
x
y
x l     ��������  ��  ��  
y 
z
{
z l     ��������  ��  ��  
{ 
|
}
| l     ��
~
��  
~ P J--------------------------------------------------------------------------   
 �
�
� � - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
} 
�
�
� l     ��������  ��  ��  
� 
�
�
� l     ��
�
���  
� + %                              getDate   
� �
�
� J                                                             g e t D a t e
� 
�
�
� l     ��������  ��  ��  
� 
�
�
� l     ��
�
���  
� P J--------------------------------------------------------------------------   
� �
�
� � - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
� 
�
�
� l     ��������  ��  ��  
� 
�
�
� l      ��
�
���  
� 
     getDate :      renvoie la date au format "lundi 31 mars 2014" � partir
                    d'une date compl�te.
     shortDate :    date compl�te (format "lundi 31 mars 2014 17:19").
     r�sultat :     date au format "jour n� jour mois ann�e".
        
� �
�
�  
           g e t D a t e   :             r e n v o i e   l a   d a t e   a u   f o r m a t   " l u n d i   3 1   m a r s   2 0 1 4 "   �   p a r t i r 
                                         d ' u n e   d a t e   c o m p l � t e . 
           s h o r t D a t e   :         d a t e   c o m p l � t e   ( f o r m a t   " l u n d i   3 1   m a r s   2 0 1 4   1 7 : 1 9 " ) . 
           r � s u l t a t   :           d a t e   a u   f o r m a t   " j o u r   n �   j o u r   m o i s   a n n � e " . 
          
� 
�
�
� i   � �
�
�
� I      ��
����� 0 getdate getDate
� 
���
� o      ���� 0 datecomplete dateComplete��  ��  
� k     T
�
� 
�
�
� l     ��������  ��  ��  
� 
�
�
� l     ��
�
���  
� b \tell current application to my logToFileAndToConsole("Entr�e dans la fonction \"getDate\".")   
� �
�
� � t e l l   c u r r e n t   a p p l i c a t i o n   t o   m y   l o g T o F i l e A n d T o C o n s o l e ( " E n t r � e   d a n s   l a   f o n c t i o n   \ " g e t D a t e \ " . " )
� 
�
�
� l     ��������  ��  ��  
� 
�
�
� O    
�
�
� I   ��
���
�� .ascrcmnt****      � ****
� c    
�
�
� o    ���� 0 datecomplete dateComplete
� m    ��
�� 
ctxt��  
� m     �
� misccura
� 
�
�
� l   �~�}�|�~  �}  �|  
� 
�
�
� r    
�
�
� c    
�
�
� o    �{�{ 0 datecomplete dateComplete
� m    �z
�z 
ctxt
� o      �y�y 0 
moninstant 
monInstant
� 
�
�
� r    
�
�
� n    
�
�
� 4  �x
�
�x 
cwor
� m    �w�w 
� o    �v�v 0 
moninstant 
monInstant
� o      �u�u 0 mon_jour_semaine_fr  
� 
�
�
� r     
�
�
� n    
�
�
� 4  �t
�
�t 
cwor
� m    �s�s 
� o    �r�r 0 
moninstant 
monInstant
� o      �q�q 0 mon_jour_fr  
� 
�
�
� r   ! '
�
�
� n   ! %
�
�
� 4 " %�p
�
�p 
cwor
� m   # $�o�o 
� o   ! "�n�n 0 
moninstant 
monInstant
� o      �m�m 0 mon_mois_fr  
� 
�
�
� r   ( .
�
�
� n   ( ,
�
�
� 4 ) ,�l
�
�l 
cwor
� m   * +�k�k 
� o   ( )�j�j 0 
moninstant 
monInstant
� o      �i�i 0 mon_annee_fr  
� 
�
�
� r   / 5
�
�
� n   / 3
�
�
� 4 0 3�h
�
�h 
cwor
� m   1 2�g�g 
� o   / 0�f�f 0 
moninstant 
monInstant
� o      �e�e 0 mon_heure_fr  
� 
�
�
� r   6 <
�
�
� n   6 :
�
�
� 4 7 :�d
�
�d 
cwor
� m   8 9�c�c 
� o   6 7�b�b 0 
moninstant 
monInstant
� o      �a�a 0 ma_minute_fr  
� 
�
�
� r   = C
�
�
� n   = A
�
�
� 4 > A�`
�
�` 
cwor
� m   ? @�_�_ 
� o   = >�^�^ 0 
moninstant 
monInstant
� o      �]�] 0 ma_seconde_fr  
� 
�
�
� L   D R
�
� b   D Q
�
�
� b   D O
�
�
� b   D M
�
�
� b   D K
�
�
� b   D I
�
�
� b   D G
�
�
� o   D E�\�\ 0 mon_jour_semaine_fr  
� m   E F
�
� �
�
�   
� o   G H�[�[ 0 mon_jour_fr  
� m   I J
�
� �
�
�   
� o   K L�Z�Z 0 mon_mois_fr  
� m   M N
�
� �
�
�   
� o   O P�Y�Y 0 mon_annee_fr  
� 
��X
� l  S S�W�V�U�W  �V  �U  �X  
� 
�
�
� l     �T�S�R�T  �S  �R  
� 
�
�
� l     �Q�P�O�Q  �P  �O  
�    l     �N�N   P J--------------------------------------------------------------------------    � � - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  l     �M�L�K�M  �L  �K    l     �J	
�J  	 4 .                              getNumeroFacture   
 � \                                                             g e t N u m e r o F a c t u r e  l     �I�H�G�I  �H  �G    l     �F�F   P J--------------------------------------------------------------------------    � � - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  l     �E�D�C�E  �D  �C    l      �B�B  %
     getNumeroFacture : renvoie le num�ro de facture � partir du nom du dernier fichier
     dans le dossier des factures (Facture_00212.pdf).
     dossierFactures :  dossier contenant les fichiers de facture.
     r�sultat :         le num�ro de facture � utiliser pour facturer.
         �> 
           g e t N u m e r o F a c t u r e   :   r e n v o i e   l e   n u m � r o   d e   f a c t u r e   �   p a r t i r   d u   n o m   d u   d e r n i e r   f i c h i e r 
           d a n s   l e   d o s s i e r   d e s   f a c t u r e s   ( F a c t u r e _ 0 0 2 1 2 . p d f ) . 
           d o s s i e r F a c t u r e s   :     d o s s i e r   c o n t e n a n t   l e s   f i c h i e r s   d e   f a c t u r e . 
           r � s u l t a t   :                   l e   n u m � r o   d e   f a c t u r e   �   u t i l i s e r   p o u r   f a c t u r e r . 
            i   � � I      �A�@�A $0 getnumerofacture getNumeroFacture �? o      �>�> "0 dossierfactures dossierFactures�?  �@   k     N   !"! l     �=�<�;�=  �<  �;  " #$# r     %&% J     �:�:  & o      �9�9 $0 listenomfichiers listeNomFichiers$ '(' l   �8�7�6�8  �7  �6  ( )*) O    L+,+ k   	 K-- ./. r   	 010 n   	 232 2    �5
�5 
file3 4   	 �44
�4 
cfol4 o    �3�3 "0 dossierfactures dossierFactures1 o      �2�2 "0 touslesfichiers tousLesFichiers/ 565 X    .7�187 k   " )99 :;: r   " '<=< l  " %>�0�/> n   " %?@? 1   # %�.
�. 
pnam@ o   " #�-�- 0 	unfichier 	unFichier�0  �/  = o      �,�, 0 
nomfichier 
nomFichier; A�+A l  ( (�*BC�*  B / )set end of listeNomFichiers to nomFichier   C �DD R s e t   e n d   o f   l i s t e N o m F i c h i e r s   t o   n o m F i c h i e r�+  �1 0 	unfichier 	unFichier8 o    �)�) "0 touslesfichiers tousLesFichiers6 EFE r   / @GHG n   / :IJI 7 0 :�(KL
�( 
ctxtK m   4 6�'�' L m   7 9�&�& J o   / 0�%�% 0 
nomfichier 
nomFichierH o      �$�$ 0 numerofacture numeroFactureF MNM l  A A�#�"�!�#  �"  �!  N O� O L   A KPP [   A JQRQ l  A HS��S c   A HTUT o   A F�� 0 numerofacture numeroFactureU m   F G�
� 
long�  �  R m   H I�� �   , m    VV�                                                                                  MACS  alis    t  Macintosh HD               �Z�H+   �=P
Finder.app                                                      ��W��[        ����  	                CoreServices    �Z�s      ��;     �=P �=O �=N  6Macintosh HD:System: Library: CoreServices: Finder.app   
 F i n d e r . a p p    M a c i n t o s h   H D  &System/Library/CoreServices/Finder.app  / ��  * WXW l  M M����  �  �  X Y�Y l  M M����  �  �  �   Z[Z l     ����  �  �  [ \]\ l     ����  �  �  ] ^_^ l     ����  �  �  _ `a` l     �
bc�
  b P J--------------------------------------------------------------------------   c �dd � - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -a efe l     �	���	  �  �  f ghg l     �ij�  i @ :           applicationShouldTerminateAfterLastWindowClosed   j �kk t                       a p p l i c a t i o n S h o u l d T e r m i n a t e A f t e r L a s t W i n d o w C l o s e dh lml l     ����  �  �  m non l     �pq�  p P J--------------------------------------------------------------------------   q �rr � - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -o sts l     �� ���  �   ��  t uvu l      ��wx��  w��
     applicationShouldTerminateAfterLastWindowClosed : Fonction � red�finir si
                                                       on veut que le programme
                                                       se termine quand on
                                                       clique sur le bouton de
                                                       fermeture de la fen�tre.
     nsApplication                                   : Application.
       x �yy� 
           a p p l i c a t i o n S h o u l d T e r m i n a t e A f t e r L a s t W i n d o w C l o s e d   :   F o n c t i o n   �   r e d � f i n i r   s i 
                                                                                                               o n   v e u t   q u e   l e   p r o g r a m m e 
                                                                                                               s e   t e r m i n e   q u a n d   o n 
                                                                                                               c l i q u e   s u r   l e   b o u t o n   d e 
                                                                                                               f e r m e t u r e   d e   l a   f e n � t r e . 
           n s A p p l i c a t i o n                                                                       :   A p p l i c a t i o n . 
        v z{z l     ��������  ��  ��  { |}| i   � �~~ I      ������� d0 0applicationshouldterminateafterlastwindowclosed_ 0applicationShouldTerminateAfterLastWindowClosed_� ���� o      ���� 0 nsapplication nsApplication��  ��   k     �� ��� l     ��������  ��  ��  � ��� O    ��� I    
������� 0 
terminate_  � ����  f    ��  ��  � o     ���� 0 nsapplication nsApplication� ���� l   ��������  ��  ��  ��  } ��� l     ��������  ��  ��  � ��� l     ��������  ��  ��  � ��� l     ��������  ��  ��  � ��� l     ��������  ��  ��  � ��� l     ������  � P J--------------------------------------------------------------------------   � ��� � - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -� ��� l     ��������  ��  ��  � ��� l     ������  � - '                              logToFile   � ��� N                                                             l o g T o F i l e� ��� l     ��������  ��  ��  � ��� l     ������  � P J--------------------------------------------------------------------------   � ��� � - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -� ��� l     ��������  ��  ��  � ��� l      ������  � � �
     logToFile :    �crit une ligne (en UTF-8) � la fin du fichier referenceVersFichierLog.
     ligne :        le texte � �crire dans le fichier.
     r�sultat :     Aucun.
        � ���h 
           l o g T o F i l e   :         � c r i t   u n e   l i g n e   ( e n   U T F - 8 )   �   l a   f i n   d u   f i c h i e r   r e f e r e n c e V e r s F i c h i e r L o g . 
           l i g n e   :                 l e   t e x t e   �   � c r i r e   d a n s   l e   f i c h i e r . 
           r � s u l t a t   :           A u c u n . 
          � ��� i   � ���� I      ������� 0 	logtofile 	logToFile� ���� o      ���� 	0 ligne  ��  ��  � k     +�� ��� l     ��������  ��  ��  � ��� r     ��� b     ��� o     ���� 	0 ligne  � m    �� ���  
� o      ���� 	0 ligne  � ��� l   ��������  ��  ��  � ��� Q    )���� I  	 ����
�� .rdwrwritnull���     ****� o   	 
���� 	0 ligne  � ����
�� 
refn� o    ���� 20 referenceversfichierlog referenceVersFichierLog� ����
�� 
wrat� m    ��
�� rdwreof � �����
�� 
as  � m    ��
�� 
utf8��  � R      ������
�� .ascrerr ****      � ****��  ��  � I    )�����
�� .rdwrclosnull���     ****� o     %���� 20 referenceversfichierlog referenceVersFichierLog��  � ���� l  * *��������  ��  ��  ��  � ��� l     ��������  ��  ��  � ��� l     ��������  ��  ��  � ��� l     ��������  ��  ��  � ��� i   � ���� I      ������� .0 logtofileandtoconsole logToFileAndToConsole� ���� o      ���� 	0 ligne  ��  ��  � k     �� ��� l     ��������  ��  ��  � ��� I     ������� 0 	logtofile 	logToFile� ���� o    ���� 	0 ligne  ��  ��  � ��� I   �����
�� .ascrcmnt****      � ****� l   ������ o    ���� 	0 ligne  ��  ��  ��  � ���� l   ��������  ��  ��  ��  � ��� l     ��������  ��  ��  � ���� l     ��������  ��  ��  ��   " ���� l     ��������  ��  ��  ��       ������  � ���� 0 appdelegate AppDelegate� �� #���� 0 appdelegate AppDelegate� �� �����
�� misccura
�� 
pcls� ���  N S O b j e c t� ,����� L U _ i n s x } � �� � �������������~!&+5:DIN�������������  � *�}�|�{�z�y�x�w�v�u�t�s�r�q�p�o�n�m�l�k�j�i�h�g�f�e�d�c�b�a�`�_�^�]�\�[�Z�Y�X�W�V�U�T
�} 
pare�| 0 prixhoraire prixHoraire�{ B0 cheminfichiermodelefactureexcel cheminFichierModeleFactureExcel�z .0 chemindossierfactures cheminDossierFactures�y (0 monadressecourrier monAdresseCourrier�x 0 masignature maSignature�w 0 monsujet monSujet�v "0 contenumessage1 contenuMessage1�u "0 contenumessage2 contenuMessage2�t 20 prefixnomfichierfacture prefixNomFichierFacture�s 20 extensionfichierfacture extensionFichierFacture�r 0 
fichierlog 
fichierLog�q 20 referenceversfichierlog referenceVersFichierLog�p 40 chemindossiertempfacture cheminDossierTempFacture�o 0 	thewindow 	theWindow�n 0 
datepicker 
datePicker�m 0 tfnomclient tfNomClient�l (0 tftypeintervention tfTypeIntervention�k *0 tfdureeintervention tfDureeIntervention�j &0 tfnumerodefacture tfNumeroDeFacture�i 0 tfprixhoraire tfPrixHoraire�h 0 numerofacture numeroFacture�g 40 cheminfichiertempfacture cheminFichierTempFacture�f &0 nomfichierfacture nomFichierFacture�e (0 cheminfacturefinal cheminFactureFinal�d ,0 dossierfacturesalias dossierFacturesAlias�c .0 adressecourrierclient adresseCourrierClient�b 0 	nomclient 	nomClient�a $0 typeintervention typeIntervention�` &0 dureeintervention dureeIntervention�_ $0 dateintervention dateIntervention�^ B0 applicationwillfinishlaunching_ applicationWillFinishLaunching_�] :0 applicationshouldterminate_ applicationShouldTerminate_�\ 0 clickbutton_ clickButton_�[ "0 showalertmodal_ showAlertModal_�Z  0 datepickerget_ datePickerGet_�Y 0 list_position  �X 0 getdate getDate�W $0 getnumerofacture getNumeroFacture�V d0 0applicationshouldterminateafterlastwindowclosed_ 0applicationShouldTerminateAfterLastWindowClosed_�U 0 	logtofile 	logToFile�T .0 logtofileandtoconsole logToFileAndToConsole��  � 'furlfile:///Users/bruno/Desktop/log.txt
�� 
msng
�� 
msng
�� 
msng
�� 
msng
�� 
msng
� 
msng
�~ 
msng� �Ss�R�Q� �P�S B0 applicationwillfinishlaunching_ applicationWillFinishLaunching_�R �O�O   �N�N 0 anotification aNotification�Q  � �M�M 0 anotification aNotification  �L�K��J��I�H��G��F�E'L[�D�C�B�
�L 
perm
�K .rdwropenshor       file�J .0 logtofileandtoconsole logToFileAndToConsole
�I 
alis�H $0 getnumerofacture getNumeroFacture
�G 
ctxt
�F 
spac
�E 
psxp
�D .misccurdldt    ��� null�C 0 setdatevalue_ setDateValue_�B "0 setstringvalue_ setStringValue_�Pb  �el Ec  O)�k+ O)�k+ Ob  �&Ec  O)b  k+ Ec  O)�b  %k+ Ob  	b  �&%b  
%Ec  O)�b  %k+ Ob  �%b  %Ec  O)�b  �,%�&k+ Ob  b  %Ec  O)�b  �,%�&k+ O)�b  %k+ O)�b  %k+ Ob  *j k+ Ob  b  k+ Ob  b  k+ O)a k+ OP� �A��@�?�>�A :0 applicationshouldterminate_ applicationShouldTerminate_�@ �=�=   �<�< 
0 sender  �?   �;�; 
0 sender   ��:�9�8�7�: .0 logtofileandtoconsole logToFileAndToConsole
�9 .rdwrclosnull���     ****
�8 misccura�7  0 nsterminatenow NSTerminateNow�> )�k+ Ob  j O��,EOP� �6'�5�4�3�6 0 clickbutton_ clickButton_�5 �2�2   �1�1 
0 sender  �4   �0�/�.�-�,�+�*�)�(�'�&�%�$�#�"�!� ��0 
0 sender  �/ *0 ficheclienttrouvees ficheclientTrouvees�. 0 ficheclient ficheClient�- 0 nomsclients nomsClients�, 0 unclient unClient�+ 0 listereponse listeReponse�* 0 pos  �)  0 formejuridique formeJuridique�( 0 contactemail contactEmail�' 0 	theemails 	theEmails�& 0 anemail anEmail�% 0 adresseclient adresseClient�$ .0 adresseclientformatee adresseClientFormatee�# 20 fichiertempfacturealias fichierTempFactureAlias�" 0 	mynewfile 	myNewFile�! 80 dateinterventionformatlong dateInterventionFormatLong�   0 contenumessage contenuMessage�  0 nouveaumessage nouveauMessage �3���Q������������������������),�1�6��
�	��������������,�=C���T�� �������������������������Rp��������������Z���������"2BT��������������������������������������������������������������������������		B	a	�	�	�� .0 logtofileandtoconsole logToFileAndToConsole� 0 stringvalue stringValue
� 
ctxt
� 
azf4  
� 
pnam
� .corecnte****       ****
� misccura
� 
ret 
� 
appr
� 
btns
� 
cbtn
� 
dflt� 
� .sysodlogaskr        TEXT
� 
cobj
� 
kocl
� 
prmp
� 
okbt
� 
cnbt
� 
mlsl�
 

�	 .gtqpchltns    @   @ ns  � 0 list_position  
� 
az51
� 
az21
� 
az18
� 
az17
� 
ascr
� 
txdl
� 
citm
�  
az27
�� 
az28
�� 
spac
�� 
az31
�� 
az29��  ��  
�� .aevtquitnull��� ��� null�� "0 setstringvalue_ setStringValue_�� 0 dateas dateAS
�� 
ldt 
�� 
psxf
�� .aevtodocnull  �    alis
�� 
1172
�� 
XwSH
�� 
ccel
�� 
DPVu
�� 
shdt
�� 
1107
�� 
5016
�� 
1813
�� e188� .�� 
�� .smXL1659null���   7 X128
�� 
savo
�� 
alis
�� 
insh
�� .coremovenull���     obj 
�� .sysodisAaleR        TEXT�� 0 getdate getDate
�� .miscactvnull��� ��� null
�� 
bcke
�� 
prdt
�� 
sndr
�� 
subj
�� 
pvis
�� 
ctnt
�� .corecrel****      � null
�� 
rcpt
�� 
trcp
�� 
radd�� 
�� 
atts
�� 
atfn
�� 
psxp�� 
�� .sysodelanull��� ��� nmbr
�� 
prcs
�� 
cwin
�� 
popB
�� .prcsclicnull��� ��� uiel
�� 
menE
�� 
menI�3�)�k+ Ob  j+ �&Ec  O)�b  %k+ O��*�-�[�,\Zb  @1E�O�j 	j  F� )�k+ UO�b  %�%�%�%�%a %a a a a kva a a a a  OPY ��j 	k  �a k/E�OPY ��j 	k �jvE�O �[a a l 	kh ��,�6F[OY��O�a a a a b  %�%a  %a !a "a #a $a %fa & 'E�O�a k/Ec  O)b  �l+ (E�O�a �/E�OPY hO��,Ec  O� )a )b  %k+ UO�a *,e  
a +E�Y a ,E�O� )a -�%k+ UO�a .-E�O�jv hY =a /b  %a 0%�%a 1%�%a 2%a a 3a a 4kva a 5a a 6a  OPO�j 	k qjvE�O *�[a a l 	kh 
�a 7,a 8%�a 9,%�6F[OY��O�a a :l 'a k/E�Oa ;_ <a =,FO�a >i/Ec  Oa ?kv_ <a =,FOPY "�j 	k  �a k/a 9,Ec  OPY hO� )a @b  %k+ UO D�a A,a k/E�O� *a B,_ C%*a D,%_ C%*a E,%E�UO� )a F�%k+ UOPW BX G Ha Ib  %a J%�%a K%�%a L%a a Ma a Nkva a Oa a Pa  OPO*j QOPUOb  b  k+ ROb  j+ �&Ec  O)a Sb  %k+ Ob  j+ �&Ec  O)a Tb  %k+ Ob  j+ Ua V&Ec  O)a Wb  %�&k+ Oa Xb  a Y&j ZO*a [,a \a ]/ �b  *a ^a _/a `,FOb  *a ^a a/a `,FOb  a b,*a ^a c/a `,FOb  *a ^a d/a `,FO�*a ^a e/a `,FO�*a ^a f/a `,FOb  *a ^a g/a `,FOb  *a ^a h/a `,FOb  b   *a ^a i/a `,FOPUOb  	b  �&%*a j,�,FO*a j, *a kb  b  
%a la ma n oUO*a pfl QOPUOa q _b  a r&E�Ob  ��,FOb  a r&Ec  O �a sb  l tE�OPW  X G Ha ub  %�%a v%�%a w%j xOPUO)b  k+ yE�Oa z �*j {Ob  �%b  %E^ O*a a |a }a ~b  a b  a �ea �] a a n �E^ O]  N*a a �a s*a �-6a }a �b  la � �O*a a �a }a �b  a �,la n �Oa �j �OPUO*j {Oa � 9*a �a �/ -*a �k/ !*a �k/j �O*a �k/a �k/a �m/j �UOPUUUO)a �k+ Ob  kEc  O)a �b  %k+ Ob  b  k+ ROb  	b  �&%b  
%Ec  O)a �b  %k+ Ob  _ C%b  %Ec  O)a �b  a �,%�&k+ Ob  b  %Ec  O)a �b  a �,%�&k+ O)a �k+ O)a �k+ OP� ��	�����	
���� "0 showalertmodal_ showAlertModal_�� ����   ���� 
0 sender  ��  	 ������ 
0 sender  �� 0 thealert theAlert
 
����	�	�	�	���������
�� misccura�� 0 nsalert NSAlert�� 20 makealert_buttons_text_ makeAlert_buttons_text_�� 0 	showmodal 	showModal
�� 
rslt
�� .ascrcmnt****      � ****�� ��,���lv�m+ E�O�j+ O�j 	� ��
	��������  0 datepickerget_ datePickerGet_�� ����   ���� 
0 sender  ��   ������ 
0 sender  �� 0 thedate theDate ��������
 
%�� 0 dateas dateAS
�� 
ldt 
�� 
ctxt
�� .ascrcmnt****      � ****�� $b  j+  �&E�O��&j O�j O�j OP� ��
K�������� 0 list_position  �� ����   ������ 0 	this_item  �� 0 	this_list  ��   ���������� 0 	this_item  �� 0 	this_list  �� 0 position  �� 0 i   ����
�� .corecnte****       ****
�� 
cobj�� /jE�O $k�j  kh ��/�  �E�Y hOP[OY��O�OP� ��
��������� 0 getdate getDate�� ����   ���� 0 datecomplete dateComplete��   	�������������������� 0 datecomplete dateComplete�� 0 
moninstant 
monInstant�� 0 mon_jour_semaine_fr  �� 0 mon_jour_fr  �� 0 mon_mois_fr  �� 0 mon_annee_fr  �� 0 mon_heure_fr  �� 0 ma_minute_fr  �� 0 ma_seconde_fr   ����������������
�
�
�
�� misccura
�� 
ctxt
�� .ascrcmnt****      � ****
�� 
cwor�� �� �� �� �� U� 	��&j UO��&E�O��k/E�O��l/E�O��m/E�O���/E�O���/E�O���/E�O���/E�O��%�%�%�%�%�%OP� ���������� $0 getnumerofacture getNumeroFacture�� ����   ���� "0 dossierfactures dossierFactures��   ������������ "0 dossierfactures dossierFactures�� $0 listenomfichiers listeNomFichiers�� "0 touslesfichiers tousLesFichiers�� 0 	unfichier 	unFichier�� 0 
nomfichier 
nomFichier V��������~�}�|�{�z�y
�� 
cfol
�� 
file
�� 
kocl
� 
cobj
�~ .corecnte****       ****
�} 
pnam
�| 
ctxt�{ �z 
�y 
long�� OjvE�O� D*�/�-E�O �[��l kh ��,E�OP[OY��O�[�\[Z�\Z�2Ec  Ob  �&kUOP� �x�w�v�u�x d0 0applicationshouldterminateafterlastwindowclosed_ 0applicationShouldTerminateAfterLastWindowClosed_�w �t�t   �s�s 0 nsapplication nsApplication�v   �r�r 0 nsapplication nsApplication �q�q 0 
terminate_  �u � *)k+  UOP� �p��o�n�m�p 0 	logtofile 	logToFile�o �l�l   �k�k 	0 ligne  �n   �j�j 	0 ligne   ��i�h�g�f�e�d�c�b�a�`
�i 
refn
�h 
wrat
�g rdwreof 
�f 
as  
�e 
utf8�d 
�c .rdwrwritnull���     ****�b  �a  
�` .rdwrclosnull���     ****�m ,��%E�O ��b  ����� W X  	b  j 
OP� �_��^�]�\�_ .0 logtofileandtoconsole logToFileAndToConsole�^ �[ �[    �Z�Z 	0 ligne  �]   �Y�Y 	0 ligne   �X�W�X 0 	logtofile 	logToFile
�W .ascrcmnt****      � ****�\ *�k+  O�j OPascr  ��ޭ