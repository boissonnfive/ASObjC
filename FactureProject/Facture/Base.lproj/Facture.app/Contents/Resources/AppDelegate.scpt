FasdUAS 1.101.10   ��   ��    k             l     ��������  ��  ��        l     �� 	 
��   	    AppDelegate.applescript    
 �   2     A p p D e l e g a t e . a p p l e s c r i p t      l     ��  ��     	  Facture     �        F a c t u r e      l     ��������  ��  ��        l     ��  ��    2 ,  Created by Bruno Boissonnet on 07/09/2014.     �   X     C r e a t e d   b y   B r u n o   B o i s s o n n e t   o n   0 7 / 0 9 / 2 0 1 4 .      l     ��  ��    = 7  Copyright (c) 2014 BInfoService. All rights reserved.     �   n     C o p y r i g h t   ( c )   2 0 1 4   B I n f o S e r v i c e .   A l l   r i g h t s   r e s e r v e d .      l     ��������  ��  ��         l     ��������  ��  ��      !�� ! h     �� "�� 0 appdelegate AppDelegate " k       # #  $ % $ j     �� &
�� 
pare & 4     �� '
�� 
pcls ' m     ( ( � ) )  N S O b j e c t %  * + * l     ��������  ��  ��   +  , - , l     �� . /��   .  
 IBOutlets    / � 0 0    I B O u t l e t s -  1 2 1 j   	 �� 3
�� 
cwin 3 m   	 
��
�� 
msng 2  4 5 4 j    �� 6�� 0 
datepicker 
datePicker 6 m    ��
�� 
msng 5  7 8 7 j    �� 9�� 0 tfnomclient tfNomClient 9 m    ��
�� 
msng 8  : ; : j    �� <�� (0 tftypeintervention tfTypeIntervention < m    ��
�� 
msng ;  = > = j    �� ?�� *0 tfdureeintervention tfDureeIntervention ? m    ��
�� 
msng >  @ A @ j    �� B�� &0 tfnumerodefacture tfNumeroDeFacture B m    ��
�� 
msng A  C D C l     ��������  ��  ��   D  E F E l     �� G H��   G ' ! Propri�t�s du script modifiables    H � I I B   P r o p r i � t � s   d u   s c r i p t   m o d i f i a b l e s F  J K J l     �� L M��   L   fichier facture    M � N N     f i c h i e r   f a c t u r e K  O P O j    �� Q�� 20 prefixnomfichierfacture prefixNomFichierFacture Q m     R R � S S  F a c t u r e _ 0 0 P  T U T j     �� V�� 20 extensionfichierfacture extensionFichierFacture V m     W W � X X  . p d f U  Y Z Y j   ! #�� [�� B0 cheminfichiermodelefactureexcel cheminFichierModeleFactureExcel [ m   ! " \ \ � ] ] x / U s e r s / b r u n o / D o c u m e n t s / b i n f o s e r v i c e / M o d e l e _ F a c t u r e _ v 0 _ 5 . x l t x Z  ^ _ ^ j   $ (�� `�� .0 chemindossierfactures cheminDossierFactures ` m   $ ' a a � b b r M a c i n t o s h   H D : U s e r s : b r u n o : D o c u m e n t s : b i n f o s e r v i c e : F a c t u r e s : _  c d c l      e f g e j   ) -�� h�� 40 chemindossiertempfacture cheminDossierTempFacture h m   ) , i i � j j B M a c i n t o s h   H D : U s e r s : b r u n o : D e s k t o p : f "  Excel ne prend pas de POSIX    g � k k 8   E x c e l   n e   p r e n d   p a s   d e   P O S I X d  l m l l     �� n o��   n   mail    o � p p 
   m a i l m  q r q j   . 2�� s�� (0 monadressecourrier monAdresseCourrier s m   . 1 t t � u u , b i n f o s e r v i c e @ g m a i l . c o m r  v w v j   3 7�� x�� 0 masignature maSignature x m   3 6 y y � z z  S i g n a t u r e   n � 1 w  { | { j   8 <�� }�� 0 monsujet monSujet } m   8 ; ~ ~ �   ( F a c t u r e   B I n f o S e r v i c e |  � � � j   = A�� ��� "0 contenumessage1 contenuMessage1 � m   = @ � � � � � � B o n j o u r , 
 
 v e u i l l e z   t r o u v e r   c i - j o i n t e   l a   f a c t u r e   d e   l ' i n t e r v e n t i o n   d u   �  � � � j   B H�� ��� "0 contenumessage2 contenuMessage2 � m   B E � � � � � < . 
 
 C o r d i a l e m e n t . 
 B r u n o . 
 
 - - - 
 
 �  � � � l     �� � ���   �   Fichier de log    � � � �    F i c h i e r   d e   l o g �  � � � j   I U�� ��� 0 
fichierlog 
fichierLog � 4   I R�� �
�� 
psxf � m   M P � � � � � 8 / U s e r s / b r u n o / D e s k t o p / l o g . t x t �  � � � j   V \�� ��� 20 referenceversfichierlog referenceVersFichierLog � m   V Y � � � � �   �  � � � l     ��������  ��  ��   �  � � � l     �� � ���   � "  Autres propri�t�s du script    � � � � 8   A u t r e s   p r o p r i � t � s   d u   s c r i p t �  � � � j   ] c�� ��� 0 numerofacture numeroFacture � m   ] ` � � � � �   �  � � � j   d j�� ��� 40 cheminfichiertempfacture cheminFichierTempFacture � m   d g � � � � �   �  � � � j   k q�� ��� &0 nomfichierfacture nomFichierFacture � m   k n � � � � �   �  � � � j   r x�� ��� (0 cheminfacturefinal cheminFactureFinal � m   r u � � � � �   �  � � � j   y �� ��� ,0 dossierfacturesalias dossierFacturesAlias � m   y | � � � � �   �  � � � l     �� � ���   � 4 .property dateDuJourCourte :                 ""    � � � � \ p r o p e r t y   d a t e D u J o u r C o u r t e   :                                   " " �  � � � j   � ��� ��� .0 adressecourrierclient adresseCourrierClient � m   � � � � � � �   �  � � � j   � ��� ��� 0 	nomclient 	nomClient � m   � � � � � � �   �  � � � l     �� � ���   � 4 .property ficheClient :                      ""    � � � � \ p r o p e r t y   f i c h e C l i e n t   :                                             " " �  � � � j   � ��� ��� $0 typeintervention typeIntervention � m   � � � � � � �   �  � � � j   � ��� ��� &0 dureeintervention dureeIntervention � m   � � � � � � �   �  � � � j   � ��� ��� $0 dateintervention dateIntervention � m   � � � � � � �   �  � � � l     ��������  ��  ��   �  � � � l     ��������  ��  ��   �  � � � l     �� � ���   � P J--------------------------------------------------------------------------    � � � � � - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - �  � � � l     ��������  ��  ��   �  � � � l     ��������  ��  ��   �  � � � l     ��������  ��  ��   �  � � � i   � � � � � I      �� ����� B0 applicationwillfinishlaunching_ applicationWillFinishLaunching_ �  ��� � o      ���� 0 anotification aNotification��  ��   � k    � � �  � � � l     �� � ���   � R L Insert code here to initialize your application before any files are opened    � � � � �   I n s e r t   c o d e   h e r e   t o   i n i t i a l i z e   y o u r   a p p l i c a t i o n   b e f o r e   a n y   f i l e s   a r e   o p e n e d �  � � � l     ��������  ��  ��   �  � � � r      � � � I    �� � �
�� .rdwropenshor       file � o     ���� 0 
fichierlog 
fichierLog � �� ��
�� 
perm  m    ��
�� boovtrue��   � o      ���� 20 referenceversfichierlog referenceVersFichierLog �  l   ��������  ��  ��    n    I    ������ 0 	logtofile 	logToFile �� m    		 �

 X 
 - - - - - - - - - -   D � b u t   d u   p r o g r a m m e   - - - - - - - - - - - - 
��  ��    f      I   ����
�� .ascrcmnt****      � **** l   ���� m     � X 
 - - - - - - - - - -   D � b u t   d u   p r o g r a m m e   - - - - - - - - - - - - 
��  ��  ��    l   ��������  ��  ��    l   ����   L F Initialisation des variables du programme ---------------------------    � �   I n i t i a l i s a t i o n   d e s   v a r i a b l e s   d u   p r o g r a m m e   - - - - - - - - - - - - - - - - - - - - - - - - - - -  l   ��������  ��  ��    n   % I     %������ 0 	logtofile 	logToFile �� m     !   �!! D - - - - >   I n i t i a l i s a t i o n   d e s   v a r i a b l e s��  ��    f      "#" I  & +��$��
�� .ascrcmnt****      � ****$ l  & '%����% m   & '&& �'' D - - - - >   I n i t i a l i s a t i o n   d e s   v a r i a b l e s��  ��  ��  # ()( l  , ,����~��  �  �~  ) *+* l  , ,�}�|�{�}  �|  �{  + ,-, l   , ,�z./�z  . U O numeroFacture : r�cup�r� � partir du num�ro de fichier de la derni�re facture    / �00 �   n u m e r o F a c t u r e   :   r � c u p � r �   �   p a r t i r   d u   n u m � r o   d e   f i c h i e r   d e   l a   d e r n i � r e   f a c t u r e  - 121 l  , ,�y�x�w�y  �x  �w  2 343 l  , ,�v56�v  5 N H set dossierFacturesAlias to (POSIX file chemindossierFactures) as alias   6 �77 �   s e t   d o s s i e r F a c t u r e s A l i a s   t o   ( P O S I X   f i l e   c h e m i n d o s s i e r F a c t u r e s )   a s   a l i a s4 898 r   , 9:;: c   , 3<=< o   , 1�u�u .0 chemindossierfactures cheminDossierFactures= m   1 2�t
�t 
alis; o      �s�s ,0 dossierfacturesalias dossierFacturesAlias9 >?> r   : J@A@ n  : DBCB I   ; D�rD�q�r $0 getnumerofacture getNumeroFactureD E�pE o   ; @�o�o ,0 dossierfacturesalias dossierFacturesAlias�p  �q  C  f   : ;A o      �n�n 0 numerofacture numeroFacture? FGF n  K WHIH I   L W�mJ�l�m 0 	logtofile 	logToFileJ K�kK b   L SLML m   L MNN �OO  N u m � r o   :  M o   M R�j�j 0 numerofacture numeroFacture�k  �l  I  f   K LG PQP I  X c�iR�h
�i .ascrcmnt****      � ****R l  X _S�g�fS b   X _TUT m   X YVV �WW  N u m � r o   :  U o   Y ^�e�e 0 numerofacture numeroFacture�g  �f  �h  Q XYX l  d d�d�c�b�d  �c  �b  Y Z[Z l  d d�a�`�_�a  �`  �_  [ \]\ l   d d�^^_�^  ^ Q K nomFichierFacture : nom du fichier facture de la forme "Facture_000X.pdf"    _ �`` �   n o m F i c h i e r F a c t u r e   :   n o m   d u   f i c h i e r   f a c t u r e   d e   l a   f o r m e   " F a c t u r e _ 0 0 0 X . p d f "  ] aba l  d d�]�\�[�]  �\  �[  b cdc r   d }efe b   d wghg b   d qiji o   d i�Z�Z 20 prefixnomfichierfacture prefixNomFichierFacturej l  i pk�Y�Xk c   i plml o   i n�W�W 0 numerofacture numeroFacturem m   n o�V
�V 
ctxt�Y  �X  h o   q v�U�U 20 extensionfichierfacture extensionFichierFacturef o      �T�T &0 nomfichierfacture nomFichierFactured non n  ~ �pqp I    ��Sr�R�S 0 	logtofile 	logToFiler s�Qs b    �tut m    �vv �ww " N o m   d u   f i c h i e r   :  u o   � ��P�P &0 nomfichierfacture nomFichierFacture�Q  �R  q  f   ~ o xyx I  � ��Oz�N
�O .ascrcmnt****      � ****z l  � �{�M�L{ b   � �|}| m   � �~~ � " N o m   d u   f i c h i e r   :  } o   � ��K�K &0 nomfichierfacture nomFichierFacture�M  �L  �N  y ��� l  � ��J�I�H�J  �I  �H  � ��� l  � ��G�F�E�G  �F  �E  � ��� l   � ��D���D  � ^ X cheminFichierTempFacture : Chemin vers le dossier temporaire de stockage de la facture    � ��� �   c h e m i n F i c h i e r T e m p F a c t u r e   :   C h e m i n   v e r s   l e   d o s s i e r   t e m p o r a i r e   d e   s t o c k a g e   d e   l a   f a c t u r e  � ��� l  � ��C�B�A�C  �B  �A  � ��� r   � ���� b   � ���� b   � ���� o   � ��@�@ 40 chemindossiertempfacture cheminDossierTempFacture� 1   � ��?
�? 
spac� o   � ��>�> &0 nomfichierfacture nomFichierFacture� o      �=�= 40 cheminfichiertempfacture cheminFichierTempFacture� ��� n  � ���� I   � ��<��;�< 0 	logtofile 	logToFile� ��:� c   � ���� b   � ���� m   � ��� ��� ( C h e m i n   t e m p o r a i r e   :  � l  � ���9�8� n   � ���� 1   � ��7
�7 
psxp� o   � ��6�6 40 cheminfichiertempfacture cheminFichierTempFacture�9  �8  � m   � ��5
�5 
ctxt�:  �;  �  f   � �� ��� I  � ��4��3
�4 .ascrcmnt****      � ****� l  � ���2�1� c   � ���� b   � ���� m   � ��� ��� ( C h e m i n   t e m p o r a i r e   :  � l  � ���0�/� n   � ���� 1   � ��.
�. 
psxp� o   � ��-�- 40 cheminfichiertempfacture cheminFichierTempFacture�0  �/  � m   � ��,
�, 
ctxt�2  �1  �3  � ��� l  � ��+�*�)�+  �*  �)  � ��� l  � ��(�'�&�(  �'  �&  � ��� l   � ��%���%  � S M cheminFactureFinal : Chemin vers le dossier final de stockage de la facture    � ��� �   c h e m i n F a c t u r e F i n a l   :   C h e m i n   v e r s   l e   d o s s i e r   f i n a l   d e   s t o c k a g e   d e   l a   f a c t u r e  � ��� l  � ��$�#�"�$  �#  �"  � ��� r   � ���� b   � ���� o   � ��!�! .0 chemindossierfactures cheminDossierFactures� o   � �� �  &0 nomfichierfacture nomFichierFacture� o      �� (0 cheminfacturefinal cheminFactureFinal� ��� n  � ���� I   � ����� 0 	logtofile 	logToFile� ��� c   � ���� b   � ���� m   � ��� ���  C h e m i n   f i n a l   :  � l  � ����� n   � ���� 1   � ��
� 
psxp� o   � ��� (0 cheminfacturefinal cheminFactureFinal�  �  � m   � ��
� 
ctxt�  �  �  f   � �� ��� I  ����
� .ascrcmnt****      � ****� l  �
���� c   �
��� b   ���� m   � ��� ���  C h e m i n   f i n a l   :  � l  ����� n   ���� 1  �
� 
psxp� o   ��� (0 cheminfacturefinal cheminFactureFinal�  �  � m  	�
� 
ctxt�  �  �  � ��� l ����  �  �  � ��� l  �
���
  � S M contenuMessage1 et contenuMessage2 : contenu du message � envoyer au client    � ��� �   c o n t e n u M e s s a g e 1   e t   c o n t e n u M e s s a g e 2   :   c o n t e n u   d u   m e s s a g e   �   e n v o y e r   a u   c l i e n t  � ��� l �	���	  �  �  � ��� l ����  � L F ATTENTION ! Il ne faut pas de tabulations ou d'espaces dans le texte.   � ��� �   A T T E N T I O N   !   I l   n e   f a u t   p a s   d e   t a b u l a t i o n s   o u   d ' e s p a c e s   d a n s   l e   t e x t e .� ��� l ����  � / ) Attention lors de l'indentation du code.   � ��� R   A t t e n t i o n   l o r s   d e   l ' i n d e n t a t i o n   d u   c o d e .� ��� r  ��� m  �� ��� � B o n j o u r , 
 
 v e u i l l e z   t r o u v e r   c i - j o i n t e   l a   f a c t u r e   d e   l ' i n t e r v e n t i o n   d u  � o      �� "0 contenumessage1 contenuMessage1� ��� n '��� I  '���� 0 	logtofile 	logToFile� ��� b  #��� m  �� ��� J C o n t e n u   d u   m e s s a g e   ( 1 � r e   p a r t i e   )   :   
� o  "� �  "0 contenumessage1 contenuMessage1�  �  �  f  � � � I (5����
�� .ascrcmnt****      � **** l (1���� b  (1 m  (+ � J C o n t e n u   d u   m e s s a g e   ( 1 � r e   p a r t i e   )   :   
 o  +0���� "0 contenumessage1 contenuMessage1��  ��  ��     r  6?	
	 m  69 � < . 
 
 C o r d i a l e m e n t . 
 B r u n o . 
 
 - - - 
 

 o      ���� "0 contenumessage2 contenuMessage2  n @N I  AN������ 0 	logtofile 	logToFile �� b  AJ m  AD � J C o n t e n u   d u   m e s s a g e   ( 2 � m e   p a r t i e   )   :   
 o  DI���� "0 contenumessage2 contenuMessage2��  ��    f  @A  I O\����
�� .ascrcmnt****      � **** l OX���� b  OX m  OR � J C o n t e n u   d u   m e s s a g e   ( 2 � m e   p a r t i e   )   :   
 o  RW���� "0 contenumessage2 contenuMessage2��  ��  ��     l ]]��������  ��  ��    !"! l  ]]��#$��  # < 6 dateDuJourCourte : Date du jour au format JJ/MM/AAAA    $ �%% l   d a t e D u J o u r C o u r t e   :   D a t e   d u   j o u r   a u   f o r m a t   J J / M M / A A A A  " &'& l ]]��������  ��  ��  ' ()( l ]]��*+��  * V Pset dateDuJourCourte to short date string of (current date) -- format JJ/MM/AAAA   + �,, � s e t   d a t e D u J o u r C o u r t e   t o   s h o r t   d a t e   s t r i n g   o f   ( c u r r e n t   d a t e )   - -   f o r m a t   J J / M M / A A A A) -.- l ]]��/0��  / I Cmy logToFile("Date du jour (forme JJ/MM/AA) : " & dateDuJourCourte)   0 �11 � m y   l o g T o F i l e ( " D a t e   d u   j o u r   ( f o r m e   J J / M M / A A )   :   "   &   d a t e D u J o u r C o u r t e ). 232 l ]]��45��  4 @ :log("Date du jour (forme JJ/MM/AA) : " & dateDuJourCourte)   5 �66 t l o g ( " D a t e   d u   j o u r   ( f o r m e   J J / M M / A A )   :   "   &   d a t e D u J o u r C o u r t e )3 787 l ]]��������  ��  ��  8 9:9 l ]]��������  ��  ��  : ;<; l ]]��=>��  = M G Initialisation de l'interface du programme ---------------------------   > �?? �   I n i t i a l i s a t i o n   d e   l ' i n t e r f a c e   d u   p r o g r a m m e   - - - - - - - - - - - - - - - - - - - - - - - - - - -< @A@ l ]]��������  ��  ��  A BCB l  ]]��DE��  D D > Initialise le calendrier de l'interface avec la date du jour    E �FF |   I n i t i a l i s e   l e   c a l e n d r i e r   d e   l ' i n t e r f a c e   a v e c   l a   d a t e   d u   j o u r  C GHG l ]]��������  ��  ��  H IJI l ]]��KL��  K " set thisDate to current date   L �MM 8 s e t   t h i s D a t e   t o   c u r r e n t   d a t eJ NON l ]]��PQ��  P : 4my logToFile("Date du jour : " & (thisDate as text))   Q �RR h m y   l o g T o F i l e ( " D a t e   d u   j o u r   :   "   &   ( t h i s D a t e   a s   t e x t ) )O STS l ]]��UV��  U 1 +log("Date du jour : " & (thisDate as text))   V �WW V l o g ( " D a t e   d u   j o u r   :   "   &   ( t h i s D a t e   a s   t e x t ) )T XYX l ]]��Z[��  Z % datePicker's setDateAS:thisDate   [ �\\ > d a t e P i c k e r ' s   s e t D a t e A S : t h i s D a t eY ]^] n ]k_`_ I  bk��a���� 0 
setdateas_ 
setDateAS_a b��b l bgc����c I bg������
�� .misccurdldt    ��� null��  ��  ��  ��  ��  ��  ` o  ]b���� 0 
datepicker 
datePicker^ ded l ll��������  ��  ��  e fgf n lzhih I  qz��j���� "0 setstringvalue_ setStringValue_j k��k o  qv���� 0 numerofacture numeroFacture��  ��  i o  lq���� &0 tfnumerodefacture tfNumeroDeFactureg lml l {{��������  ��  ��  m non n {�pqp I  |���r���� 0 	logtofile 	logToFiler s��s m  |tt �uu D < - - - -   I n i t i a l i s a t i o n   d e s   v a r i a b l e s��  ��  q  f  {|o vwv I ����x��
�� .ascrcmnt****      � ****x l ��y����y m  ��zz �{{ D < - - - -   I n i t i a l i s a t i o n   d e s   v a r i a b l e s��  ��  ��  w |��| l ����������  ��  ��  ��   � }~} l     ��������  ��  ��  ~ � l     ��������  ��  ��  � ��� l     ��������  ��  ��  � ��� l     ��������  ��  ��  � ��� i   � ���� I      ������� :0 applicationshouldterminate_ applicationShouldTerminate_� ���� o      ���� 
0 sender  ��  ��  � k     �� ��� l     ��������  ��  ��  � ��� l     ������  � L F Insert code here to do any housekeeping before your application quits   � ��� �   I n s e r t   c o d e   h e r e   t o   d o   a n y   h o u s e k e e p i n g   b e f o r e   y o u r   a p p l i c a t i o n   q u i t s� ��� n    ��� I    ������� 0 	logtofile 	logToFile� ���� m    �� ��� X 
 - - - - - - - - - - -   F i n   d u   p r o g r a m m e   - - - - - - - - - - - - - 
��  ��  �  f     � ��� I   �����
�� .ascrcmnt****      � ****� l   ������ m    �� ��� X 
 - - - - - - - - - - -   F i n   d u   p r o g r a m m e   - - - - - - - - - - - - - 
��  ��  ��  � ��� l   ��������  ��  ��  � ��� I   �����
�� .rdwrclosnull���     ****� o    ���� 20 referenceversfichierlog referenceVersFichierLog��  � ��� l   ��������  ��  ��  � ��� L    �� n   ��� o    ����  0 nsterminatenow NSTerminateNow� m    ��
�� misccura� ���� l   ��������  ��  ��  ��  � ��� l     ��������  ��  ��  � ��� l     ��������  ��  ��  � ��� l     ��������  ��  ��  � ��� l     ��������  ��  ��  � ��� i   � ���� I      ���~� 0 clickbutton_ clickButton_� ��}� o      �|�| 
0 sender  �}  �~  � k    ��� ��� l     �{�z�y�{  �z  �y  � ��� n    ��� I    �x��w�x 0 	logtofile 	logToFile� ��v� m    �� ��� . 
 - - - - >   N o u v e l l e   f a c t u r e�v  �w  �  f     � ��� I   �u��t
�u .ascrcmnt****      � ****� l   ��s�r� m    �� ��� . 
 - - - - >   N o u v e l l e   f a c t u r e�s  �r  �t  � ��� l   �q�p�o�q  �p  �o  � ��� l   �n�m�l�n  �m  �l  � ��� l    �k���k  � x r**************************** R�cup�ration des informations du client dans Contacts *******************************   � ��� � * * * * * * * * * * * * * * * * * * * * * * * * * * * *   R � c u p � r a t i o n   d e s   i n f o r m a t i o n s   d u   c l i e n t   d a n s   C o n t a c t s   * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *� ��� l   �j�i�h�j  �i  �h  � ��� r    ��� c    ��� l   ��g�f� n   ��� I    �e�d�c�e 0 stringvalue stringValue�d  �c  � o    �b�b 0 tfnomclient tfNomClient�g  �f  � m    �a
�a 
ctxt� o      �`�` 0 	nomclient 	nomClient� ��� n   +��� I     +�_��^�_ 0 	logtofile 	logToFile� ��]� b     '��� m     !�� ��� & C l i e n t   r e c h e r c h �   :  � o   ! &�\�\ 0 	nomclient 	nomClient�]  �^  �  f     � ��� I  , 7�[��Z
�[ .ascrcmnt****      � ****� l  , 3��Y�X� b   , 3��� m   , -�� ��� & C l i e n t   r e c h e r c h �   :  � o   - 2�W�W 0 	nomclient 	nomClient�Y  �X  �Z  � ��� l  8 8�V�U�T�V  �U  �T  � ��� l  8 8�S�R�Q�S  �R  �Q  � ��� l   8 8�P���P  � 9 3 Cherche le client dans l'application � Contacts �    � ��� f   C h e r c h e   l e   c l i e n t   d a n s   l ' a p p l i c a t i o n   �   C o n t a c t s   �  � ��� l  8s�� � O   8s k   <r  l  < <�O�N�M�O  �N  �M    l  < <�L	�L   W Qtell current application to my logToFile("On appelle l'application \"Contacts\"")   	 �

 � t e l l   c u r r e n t   a p p l i c a t i o n   t o   m y   l o g T o F i l e ( " O n   a p p e l l e   l ' a p p l i c a t i o n   \ " C o n t a c t s \ " " )  l  < <�K�K   N Htell current application to log("On appelle l'application \"Contacts\"")    � � t e l l   c u r r e n t   a p p l i c a t i o n   t o   l o g ( " O n   a p p e l l e   l ' a p p l i c a t i o n   \ " C o n t a c t s \ " " )  l  < <�J�I�H�J  �I  �H    l   < <�G�G   7 1 R�cup�ration de la fiche client � partir du nom     � b   R � c u p � r a t i o n   d e   l a   f i c h e   c l i e n t   �   p a r t i r   d u   n o m    l  < <�F�E�D�F  �E  �D    l  < <�C�C   7 1 Liste de toutes les personnes qui portent ce nom    � b   L i s t e   d e   t o u t e s   l e s   p e r s o n n e s   q u i   p o r t e n t   c e   n o m  r   < N !  6  < L"#" 2   < ?�B
�B 
azf4# E   @ K$%$ 1   A C�A
�A 
pnam% o   D J�@�@ 0 	nomclient 	nomClient! o      �?�? *0 ficheclienttrouvees ficheclientTrouvees &'& l  O O�>�=�<�>  �=  �<  ' ()( Z   OU*+,�;* =   O V-.- l  O T/�:�9/ I  O T�80�7
�8 .corecnte****       ****0 o   O P�6�6 *0 ficheclienttrouvees ficheclientTrouvees�7  �:  �9  . m   T U�5�5  + k   Y �11 232 l  Y Y�4�3�2�4  �3  �2  3 454 O  Y d676 n  ] c898 I   ^ c�1:�0�1 0 	logtofile 	logToFile: ;�/; m   ^ _<< �== @ A u c u n e   p e r s o n n e   n e   p o r t e   c e   n o m .�/  �0  9  f   ] ^7 m   Y Z�.
�. misccura5 >?> O  e o@A@ I  i n�-B�,
�- .ascrcmnt****      � ****B l  i jC�+�*C m   i jDD �EE @ A u c u n e   p e r s o n n e   n e   p o r t e   c e   n o m .�+  �*  �,  A m   e f�)
�) misccura? FGF l   p p�(HI�(  H @ : Aucune personne ne porte ce nom, on d�clenche une erreur	   I �JJ t   A u c u n e   p e r s o n n e   n e   p o r t e   c e   n o m ,   o n   d � c l e n c h e   u n e   e r r e u r 	G KLK l  p p�'�&�%�'  �&  �%  L MNM I  p ��$OP
�$ .sysodlogaskr        TEXTO b   p �QRQ b   p �STS b   p �UVU b   p �WXW b   p }YZY b   p y[\[ m   p s]] �^^ T A u c u n   c o n t a c t   d o n t   l e   n o m   d e   f a m i l l e   e s t   "\ l 	 s x_�#�"_ o   s x�!�! 0 	nomclient 	nomClient�#  �"  Z m   y |`` �aa " "   n ' a   � t �   t r o u v � .X o   } �� 
�  
ret V l 	 � �b��b m   � �cc �dd f V e u i l l e z   c r � e r   l e   c o n t a c t   e t   r e l a n c e r   l e   p r o g r a m m e .�  �  T o   � ��
� 
ret R l 	 � �e��e m   � �ff �gg " F i n   d u   p r o g r a m m e .�  �  P �hi
� 
apprh l 	 � �j��j m   � �kk �ll  F a c t u r e�  �  i �mn
� 
btnsm J   � �oo p�p m   � �qq �rr  T e r m i n e r�  n �st
� 
cbtns m   � �uu �vv  T e r m i n e rt �w�
� 
dfltw m   � �xx �yy  T e r m i n e r�  N z�z l  � �����  �  �  �  , {|{ =   � �}~} l  � ��� I  � ����
� .corecnte****       ****� o   � ��
�
 *0 ficheclienttrouvees ficheclientTrouvees�  �  �  ~ m   � ��	�	 | ��� k   � ��� ��� l  � �����  �  �  � ��� l  � �����  � b \tell current application to my logToFile("Une personne porte ce nom, on r�cup�re sa fiche.")   � ��� � t e l l   c u r r e n t   a p p l i c a t i o n   t o   m y   l o g T o F i l e ( " U n e   p e r s o n n e   p o r t e   c e   n o m ,   o n   r � c u p � r e   s a   f i c h e . " )� ��� l  � �����  � Y Stell current application to log("Une personne porte ce nom, on r�cup�re sa fiche.")   � ��� � t e l l   c u r r e n t   a p p l i c a t i o n   t o   l o g ( " U n e   p e r s o n n e   p o r t e   c e   n o m ,   o n   r � c u p � r e   s a   f i c h e . " )� ��� l   � �����  � 8 2 Une personne porte ce nom, on r�cup�re sa fiche.    � ��� d   U n e   p e r s o n n e   p o r t e   c e   n o m ,   o n   r � c u p � r e   s a   f i c h e .  � ��� l  � ���� �  �  �   � ��� r   � ���� n   � ���� 4  � ����
�� 
cobj� m   � ����� � o   � ����� *0 ficheclienttrouvees ficheclientTrouvees� o      ���� 0 ficheclient ficheClient� ���� l  � ���������  ��  ��  ��  � ��� ?   � ���� l  � ������� I  � ������
�� .corecnte****       ****� o   � ����� *0 ficheclienttrouvees ficheclientTrouvees��  ��  ��  � m   � ����� � ���� k   �Q�� ��� l  � ���������  ��  ��  � ��� l  � �������  � U Otell current application to my logToFile("Plusieurs personnes portent ce nom.")   � ��� � t e l l   c u r r e n t   a p p l i c a t i o n   t o   m y   l o g T o F i l e ( " P l u s i e u r s   p e r s o n n e s   p o r t e n t   c e   n o m . " )� ��� l  � �������  � L Ftell current application to log("Plusieurs personnes portent ce nom.")   � ��� � t e l l   c u r r e n t   a p p l i c a t i o n   t o   l o g ( " P l u s i e u r s   p e r s o n n e s   p o r t e n t   c e   n o m . " )� ��� l   � �������  � � �
                 Plusieurs personnes portent ce nom.
                 On met les noms complets (propri�t� name) dans une liste
                 et on demande � l'utilisateur de s�lectionner la personne recherch�e.
                    � ���� 
                                   P l u s i e u r s   p e r s o n n e s   p o r t e n t   c e   n o m . 
                                   O n   m e t   l e s   n o m s   c o m p l e t s   ( p r o p r i � t �   n a m e )   d a n s   u n e   l i s t e 
                                   e t   o n   d e m a n d e   �   l ' u t i l i s a t e u r   d e   s � l e c t i o n n e r   l a   p e r s o n n e   r e c h e r c h � e . 
                                  � ��� l  � ���������  ��  ��  � ��� r   � ���� J   � �����  � o      ���� 0 nomsclients nomsClients� ��� l  � ���������  ��  ��  � ��� X   � ������ r   � ���� n   � ���� 1   � ���
�� 
pnam� o   � ����� 0 unclient unClient� n      ���  ;   � �� o   � ����� 0 nomsclients nomsClients�� 0 unclient unClient� o   � ����� *0 ficheclienttrouvees ficheclientTrouvees� ��� l  � ���������  ��  ��  � ��� r   �+��� I  �)����
�� .gtqpchltns    @   @ ns  � l 
 � ������� o   � ����� 0 nomsclients nomsClients��  ��  � ����
�� 
appr� l 	 � ������� m   � ��� ���  F a c t u r e��  ��  � ����
�� 
prmp� b  ��� b  ��� b  ��� m  �� ��� T I l   y   a   p l u s   d ' u n   c o n t a c t   q u i   p o r t e   l e   n o m  � o  
���� 0 	nomclient 	nomClient� o  ��
�� 
ret � l 	������ m  �� ��� p C h o i s i s s e z   d a n s   l a   l i s t e   l e   c o n t a c t   q u e   v o u s   r e c h e r c h e z .��  ��  � ����
�� 
okbt� l 	������ m  �� ���  O K��  ��  � ����
�� 
cnbt� l 	������ m  �� ���  A n n u l e r��  ��  � �����
�� 
mlsl� m  "#��
�� boovfals��  � o      ���� 0 listereponse listeReponse� ��� l ,,��������  ��  ��  � ��� r  ,8��� n  ,2��� 4  -2���
�� 
cobj� m  01���� � o  ,-���� 0 listereponse listeReponse� o      ���� 0 	nomclient 	nomClient� ��� l 99������  �  display dialog nomClient   � ��� 0 d i s p l a y   d i a l o g   n o m C l i e n t� ��� l 99��������  ��  ��  �    l  99����   � �
                 De la position du nom dans la liste,
                 on trouve la fiche dans la liste des clients
                     � 
                                   D e   l a   p o s i t i o n   d u   n o m   d a n s   l a   l i s t e , 
                                   o n   t r o u v e   l a   f i c h e   d a n s   l a   l i s t e   d e s   c l i e n t s 
                                    r  9F n 9D	
	 I  :D������ 0 list_position    o  :?���� 0 	nomclient 	nomClient �� o  ?@���� 0 nomsclients nomsClients��  ��  
  f  9: o      ���� 0 pos    l GG����    display dialog pos    � $ d i s p l a y   d i a l o g   p o s  l GG��������  ��  ��    r  GO n  GM 4  HM��
�� 
cobj o  KL���� 0 pos   o  GH���� *0 ficheclienttrouvees ficheclientTrouvees o      ���� 0 ficheclient ficheClient �� l PP��������  ��  ��  ��  ��  �;  )  l VV��������  ��  ��    !  l VV��������  ��  ��  ! "#" l  VV��$%��  $ - ' R�cup�ration du nom complet du client    % �&& N   R � c u p � r a t i o n   d u   n o m   c o m p l e t   d u   c l i e n t  # '(' l VV��������  ��  ��  ( )*) r  V_+,+ n  VY-.- 1  WY��
�� 
pnam. o  VW���� 0 ficheclient ficheClient, o      ���� 0 	nomclient 	nomClient* /0/ O `s121 n dr343 I  er��5���� 0 	logtofile 	logToFile5 6��6 b  en787 m  eh99 �::   N o m   d u   c l i e n t   :  8 o  hm���� 0 	nomclient 	nomClient��  ��  4  f  de2 m  `a��
�� misccura0 ;<; O t�=>= I x���?��
�� .ascrcmnt****      � ****? l x�@����@ b  x�ABA m  x{CC �DD   N o m   d u   c l i e n t   :  B o  {����� 0 	nomclient 	nomClient��  ��  ��  > m  tu��
�� misccura< EFE l ����������  ��  ��  F GHG l ����������  ��  ��  H IJI l  ����KL��  K * $ R�cup�ration de la forme juridique    L �MM H   R � c u p � r a t i o n   d e   l a   f o r m e   j u r i d i q u e  J NON Z  ��PQ��RP = ��STS n  ��UVU 1  ����
�� 
az51V o  ������ 0 ficheclient ficheClientT m  ����
�� boovtrueQ r  ��WXW m  ��YY �ZZ  E n t r e p r i s eX o      ����  0 formejuridique formeJuridique��  R r  ��[\[ m  ��]] �^^  P a r t i c u l i e r\ o      ����  0 formejuridique formeJuridiqueO _`_ l ���������  ��  �  ` aba O ��cdc n ��efe I  ���~g�}�~ 0 	logtofile 	logToFileg h�|h b  ��iji m  ��kk �ll 8 F o r m e   j u r i d i q u e   d u   c l i e n t   :  j o  ���{�{  0 formejuridique formeJuridique�|  �}  f  f  ��d m  ���z
�z misccurab mnm O ��opo I ���yq�x
�y .ascrcmnt****      � ****q l ��r�w�vr b  ��sts m  ��uu �vv 8 F o r m e   j u r i d i q u e   d u   c l i e n t   :  t o  ���u�u  0 formejuridique formeJuridique�w  �v  �x  p m  ���t
�t misccuran wxw l ���s�r�q�s  �r  �q  x yzy l ���p�o�n�p  �o  �n  z {|{ l  ���m}~�m  } B < R�cup�ration de l'adresse mail � partir de la fiche client    ~ � x   R � c u p � r a t i o n   d e   l ' a d r e s s e   m a i l   �   p a r t i r   d e   l a   f i c h e   c l i e n t  | ��� l ���l�k�j�l  �k  �j  � ��� l ���i���i  �  repeat   � ���  r e p e a t� ��� l ���h�g�f�h  �g  �f  � ��� r  ����� n  ����� 2 ���e
�e 
az21� o  ���d�d 0 ficheclient ficheClient� o      �c�c 0 contactemail contactEmail� ��� l ���b�a�`�b  �a  �`  � ��� Z  ����_�� >  ����� o  ���^�^ 0 contactemail contactEmail� J  ���]�]  � k  ���� ��� l ���\�[�Z�\  �[  �Z  � ��� l ���Y���Y  �  exit repeat   � ���  e x i t   r e p e a t� ��X� l ���W�V�U�W  �V  �U  �X  �_  � k  ��� ��� l ���T�S�R�T  �S  �R  � ��� I ��Q��
�Q .sysodlogaskr        TEXT� b  ����� b  ����� b  ����� b  ����� b  ����� b  ����� m  ���� ��� ^ A u c u n e   a d r e s s e   m � l e   r e n s e i g n � e   p o u r   l e   c l i e n t   "� o  ���P�P 0 	nomclient 	nomClient� m  ���� ���  " .� l 	����O�N� o  ���M
�M 
ret �O  �N  � l 	����L�K� m  ���� ��� x V e u i l l e z   c o m p l � t e r   l a   f i c h e   c l i e n t   e t   r e l a n c e r   l e   p r o g r a m m e .�L  �K  � o  ���J
�J 
ret � l 	����I�H� m  ���� ��� " F i n   d u   p r o g r a m m e .�I  �H  � �G��
�G 
appr� l 	����F�E� m  ���� ���  F a c t u r e�F  �E  � �D��
�D 
btns� J  ���� ��C� m  ���� ���  T e r m i n e r�C  � �B��
�B 
cbtn� m  ��� ���  T e r m i n e r� �A��@
�A 
dflt� m  �� ���  T e r m i n e r�@  � ��?� l �>�=�<�>  �=  �<  �?  � ��� l �;�:�9�;  �:  �9  � ��� l �8���8  �  
end repeat   � ���  e n d   r e p e a t� ��� l �7�6�5�7  �6  �5  � ��� Z  �����4� ?  ��� l ��3�2� I �1��0
�1 .corecnte****       ****� o  �/�/ 0 contactemail contactEmail�0  �3  �2  � m  �.�. � k  ��� ��� l �-�,�+�-  �,  �+  � ��� r  ��� J  �*�*  � o      �)�) 0 	theemails 	theEmails� ��� X   K��(�� r  4F��� b  4C��� b  4=��� n  49��� 1  59�'
�' 
az18� o  45�&�& 0 anemail anEmail� m  9<�� ���  :  � n  =B��� 1  >B�%
�% 
az17� o  =>�$�$ 0 anemail anEmail� n      ���  ;  DE� o  CD�#�# 0 	theemails 	theEmails�( 0 anemail anEmail� o  #$�"�" 0 contactemail contactEmail� ��� r  L^� � n  L\ 4 W\�!
�! 
cobj m  Z[� �   l LW�� I LW�
� .gtqpchltns    @   @ ns   o  LM�� 0 	theemails 	theEmails ��
� 
prmp m  PS �		 b V e u i l l e z   s � l e c t i o n n e r   l ' a d r e s s e   m a i l   �   u t i l i s e r   :�  �  �    o      �� 0 contactemail contactEmail� 

 r  _j m  _b �  :   n      1  ei�
� 
txdl 1  be�
� 
ascr  r  kw n  kq 4  lq�
� 
citm m  op���� o  kl�� 0 contactemail contactEmail o      �� .0 adressecourrierclient adresseCourrierClient  r  x� J  x} � m  x{ �    �   n     !"! 1  ���
� 
txdl" 1  }��
� 
ascr #$# l ������  �  �  $ %&% l ���'(�  ' n htell current application to my logToFile("Adresse mail principale du client : " & adresseCourrierClient)   ( �)) � t e l l   c u r r e n t   a p p l i c a t i o n   t o   m y   l o g T o F i l e ( " A d r e s s e   m a i l   p r i n c i p a l e   d u   c l i e n t   :   "   &   a d r e s s e C o u r r i e r C l i e n t )& *+* l ���,-�  , e _tell current application to log("Adresse mail principale du client : " & adresseCourrierClient)   - �.. � t e l l   c u r r e n t   a p p l i c a t i o n   t o   l o g ( " A d r e s s e   m a i l   p r i n c i p a l e   d u   c l i e n t   :   "   &   a d r e s s e C o u r r i e r C l i e n t )+ /�
/ l ���	���	  �  �  �
  � 010 =  ��232 l ��4��4 I ���5�
� .corecnte****       ****5 o  ���� 0 contactemail contactEmail�  �  �  3 m  ���� 1 6� 6 k  ��77 898 l ����������  ��  ��  9 :;: r  ��<=< n  ��>?> 1  ����
�� 
az17? n  ��@A@ 4 ����B
�� 
cobjB m  ������ A o  ������ 0 contactemail contactEmail= o      ���� .0 adressecourrierclient adresseCourrierClient; CDC l ����EF��  E c ]tell current application to my logToFile("Adresse mail du client : " & adresseCourrierClient)   F �GG � t e l l   c u r r e n t   a p p l i c a t i o n   t o   m y   l o g T o F i l e ( " A d r e s s e   m a i l   d u   c l i e n t   :   "   &   a d r e s s e C o u r r i e r C l i e n t )D HIH l ����JK��  J Z Ttell current application to log("Adresse mail du client : " & adresseCourrierClient)   K �LL � t e l l   c u r r e n t   a p p l i c a t i o n   t o   l o g ( " A d r e s s e   m a i l   d u   c l i e n t   :   "   &   a d r e s s e C o u r r i e r C l i e n t )I M��M l ����������  ��  ��  ��  �   �4  � NON l ����������  ��  ��  O PQP O ��RSR n ��TUT I  ����V���� 0 	logtofile 	logToFileV W��W b  ��XYX m  ��ZZ �[[ 2 A d r e s s e   m a i l   d u   c l i e n t   :  Y o  ������ .0 adressecourrierclient adresseCourrierClient��  ��  U  f  ��S m  ����
�� misccuraQ \]\ O ��^_^ I ����`��
�� .ascrcmnt****      � ****` l ��a����a b  ��bcb m  ��dd �ee 2 A d r e s s e   m a i l   d u   c l i e n t   :  c o  ������ .0 adressecourrierclient adresseCourrierClient��  ��  ��  _ m  ����
�� misccura] fgf l ����������  ��  ��  g hih l  ����jk��  j = 7 R�cup�ration de l'adresse � partir de la fiche client    k �ll n   R � c u p � r a t i o n   d e   l ' a d r e s s e   �   p a r t i r   d e   l a   f i c h e   c l i e n t  i mnm l ����������  ��  ��  n opo Q  �jqrsq k  �#tt uvu r  ��wxw n  ��yzy 4 ����{
�� 
cobj{ m  ������ z n  ��|}| m  ����
�� 
az27} o  ������ 0 ficheclient ficheClientx o      ���� 0 adresseclient adresseClientv ~~ O  ���� r  ���� b  ����� b  ����� b  ����� b  ����� 1  ����
�� 
az28� 1  ����
�� 
spac� 1  ����
�� 
az31� 1  ����
�� 
spac� 1  ����
�� 
az29� o      ���� .0 adresseclientformatee adresseClientFormatee� o  ������ 0 adresseclient adresseClient ��� l ��������  ��  ��  � ��� O ��� n ��� I  ������� 0 	logtofile 	logToFile� ���� b  ��� m  �� ��� ( A d r e s s e   d u   c l i e n t   :  � o  ���� .0 adresseclientformatee adresseClientFormatee��  ��  �  f  � m  ��
�� misccura� ��� O !��� I  �����
�� .ascrcmnt****      � ****� l ������ b  ��� m  �� ��� ( A d r e s s e   d u   c l i e n t   :  � o  ���� .0 adresseclientformatee adresseClientFormatee��  ��  ��  � m  ��
�� misccura� ���� l ""��������  ��  ��  ��  r R      ������
�� .ascrerr ****      � ****��  ��  s k  +j�� ��� l ++��������  ��  ��  � ��� I +h����
�� .sysodlogaskr        TEXT� b  +H��� b  +D��� b  +@��� b  +<��� b  +8��� b  +4��� m  +.�� ��� T A u c u n e   a d r e s s e   r e n s e i g n � e   p o u r   l e   c l i e n t   "� o  .3���� 0 	nomclient 	nomClient� m  47�� ���  " .� l 	8;������ o  8;��
�� 
ret ��  ��  � m  <?�� ��� x V e u i l l e z   c o m p l � t e r   l a   f i c h e   c l i e n t   e t   r e l a n c e r   l e   p r o g r a m m e .� o  @C��
�� 
ret � l 	DG������ m  DG�� ��� " F i n   d u   p r o g r a m m e .��  ��  � ����
�� 
appr� l 	KN������ m  KN�� ���  F a c t u r e��  ��  � ����
�� 
btns� J  QV�� ���� m  QT�� ���  T e r m i n e r��  � ����
�� 
cbtn� m  Y\�� ���  T e r m i n e r� �����
�� 
dflt� m  _b�� ���  T e r m i n e r��  � ���� l ii��������  ��  ��  ��  p ��� l kk��������  ��  ��  � ��� l kk��������  ��  ��  � ��� I kp������
�� .aevtquitnull��� ��� null��  ��  � ���� l qq��������  ��  ��  ��   m   8 9���                                                                                  adrb  alis    V  Macintosh HD               �Z�H+     UContacts.app                                                     Ѿ�I        ����  	                Applications    �Z�s      � �)       U  'Macintosh HD:Applications: Contacts.app     C o n t a c t s . a p p    M a c i n t o s h   H D  Applications/Contacts.app   / ��  �   application "Contacts"     ��� .   a p p l i c a t i o n   " C o n t a c t s "� ��� l tt��������  ��  ��  � ��� l tt������  �   display dialog nomClient   � ��� 2   d i s p l a y   d i a l o g   n o m C l i e n t� ��� l tt��������  ��  ��  � ��� l tt������  � ) # log "Nom du client : " & nomClient   � ��� F   l o g   " N o m   d u   c l i e n t   :   "   &   n o m C l i e n t� ��� l tt������  � > 8 log "Adresse mail du client : " & adresseCourrierClient   � ��� p   l o g   " A d r e s s e   m a i l   d u   c l i e n t   :   "   &   a d r e s s e C o u r r i e r C l i e n t� ��� l tt������  � 9 3 log "Adresse du client : " & adresseClientFormatee   � ��� f   l o g   " A d r e s s e   d u   c l i e n t   :   "   &   a d r e s s e C l i e n t F o r m a t e e� ��� l tt��������  ��  ��  � ��� n t���� I  y�������� "0 setstringvalue_ setStringValue_� ���� o  y~���� 0 	nomclient 	nomClient��  ��  � o  ty���� 0 tfnomclient tfNomClient� ��� l ����������  ��  ��  �    l ������   , &set theText to leTexte's stringValue()    � L s e t   t h e T e x t   t o   l e T e x t e ' s   s t r i n g V a l u e ( )  l ������   2 ,display alert "Vous avez �crit : " & theText    �		 X d i s p l a y   a l e r t   " V o u s   a v e z   � c r i t   :   "   &   t h e T e x t 

 l ������~��  �  �~    l ���}�}   / )leTexte's setStringValue_("rien du tout")    � R l e T e x t e ' s   s e t S t r i n g V a l u e _ ( " r i e n   d u   t o u t " )  l ���|�|   8 2unAutreTexte's setStringValue_("que dalle, oui !")    � d u n A u t r e T e x t e ' s   s e t S t r i n g V a l u e _ ( " q u e   d a l l e ,   o u i   ! " )  l ���{�z�y�{  �z  �y    l ���x�w�v�x  �w  �v    l  ���u�u   \ V*********************** R�cup�ration des donn�es de l'intervention *******************    � � * * * * * * * * * * * * * * * * * * * * * * *   R � c u p � r a t i o n   d e s   d o n n � e s   d e   l ' i n t e r v e n t i o n   * * * * * * * * * * * * * * * * * * *   l ���t�s�r�t  �s  �r    !"! l ���q�p�o�q  �p  �o  " #$# l  ���n%&�n  % + % R�cup�ration du type d'intervention    & �'' J   R � c u p � r a t i o n   d u   t y p e   d ' i n t e r v e n t i o n  $ ()( l ���m�l�k�m  �l  �k  ) *+* r  ��,-, c  ��./. l ��0�j�i0 n ��121 I  ���h�g�f�h 0 stringvalue stringValue�g  �f  2 o  ���e�e (0 tftypeintervention tfTypeIntervention�j  �i  / m  ���d
�d 
ctxt- o      �c�c $0 typeintervention typeIntervention+ 343 n ��565 I  ���b7�a�b 0 	logtofile 	logToFile7 8�`8 b  ��9:9 m  ��;; �<< : L e   t y p e   d ' i n t e r v e n t i o n   e s t   :  : o  ���_�_ $0 typeintervention typeIntervention�`  �a  6  f  ��4 =>= I ���^?�]
�^ .ascrcmnt****      � ****? l ��@�\�[@ b  ��ABA m  ��CC �DD : L e   t y p e   d ' i n t e r v e n t i o n   e s t   :  B o  ���Z�Z $0 typeintervention typeIntervention�\  �[  �]  > EFE l ���Y�X�W�Y  �X  �W  F GHG l ���V�U�T�V  �U  �T  H IJI l  ���SKL�S  K 2 , R�cup�ration de la dur�e de l'intervention    L �MM X   R � c u p � r a t i o n   d e   l a   d u r � e   d e   l ' i n t e r v e n t i o n  J NON l ���R�Q�P�R  �Q  �P  O PQP r  ��RSR c  ��TUT l ��V�O�NV n ��WXW I  ���M�L�K�M 0 stringvalue stringValue�L  �K  X o  ���J�J *0 tfdureeintervention tfDureeIntervention�O  �N  U m  ���I
�I 
ctxtS o      �H�H &0 dureeintervention dureeInterventionQ YZY n ��[\[ I  ���G]�F�G 0 	logtofile 	logToFile] ^�E^ b  ��_`_ m  ��aa �bb < L a   d u r � e   d ' i n t e r v e n t i o n   e s t   :  ` o  ���D�D &0 dureeintervention dureeIntervention�E  �F  \  f  ��Z cdc I ���Ce�B
�C .ascrcmnt****      � ****e l ��f�A�@f b  ��ghg m  ��ii �jj < L a   d u r � e   d ' i n t e r v e n t i o n   e s t   :  h o  ���?�? &0 dureeintervention dureeIntervention�A  �@  �B  d klk l ���>�=�<�>  �=  �<  l mnm l ���;�:�9�;  �:  �9  n opo l  ���8qr�8  q 1 + R�cup�ration de la date de l'intervention    r �ss V   R � c u p � r a t i o n   d e   l a   d a t e   d e   l ' i n t e r v e n t i o n  p tut r  ��vwv l ��x�7�6x c  ��yzy n ��{|{ I  ���5�4�3�5 0 dateas dateAS�4  �3  | o  ���2�2 0 
datepicker 
datePickerz m  ���1
�1 
ldt �7  �6  w o      �0�0 $0 dateintervention dateInterventionu }~} n �� I  ��/��.�/ 0 	logtofile 	logToFile� ��-� c  ���� b  ����� m  ���� ��� @ L a   d a t e   d e   l ' i n t e r v e n t i o n   e s t   :  � o  ���,�, $0 dateintervention dateIntervention� m  � �+
�+ 
ctxt�-  �.  �  f  ��~ ��� I �*��)
�* .ascrcmnt****      � ****� l ��(�'� c  ��� b  ��� m  	�� ��� @ L a   d a t e   d e   l ' i n t e r v e n t i o n   e s t   :  � o  	�&�& $0 dateintervention dateIntervention� m  �%
�% 
ctxt�(  �'  �)  � ��� l �$�#�"�$  �#  �"  � ��� l  �!���!  �]W
         
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
                  � ��� l � ���   �  �  � ��� l ����  �  �  � ��� l  ����  � _ Y**************************** Cr�ation de la facture dans Excel **************************   � ��� � * * * * * * * * * * * * * * * * * * * * * * * * * * * *   C r � a t i o n   d e   l a   f a c t u r e   d a n s   E x c e l   * * * * * * * * * * * * * * * * * * * * * * * * * *� ��� l ����  �  �  � ��� l ����  �  �  � ��� l  ����  �71 Cr�e un fichier excel � partir d'un mod�le, le modifie en fonction des donn�es entr�es et l'enregistre en pdf sur le bureau.
         Le fichier aura le nom: " Facture_00XXX.pdf" avec une espace devant ajout�e automatiquement par excel qui concat�ne le nom du fichier avec le nom de la feuille de calcul.   � ���b   C r � e   u n   f i c h i e r   e x c e l   �   p a r t i r   d ' u n   m o d � l e ,   l e   m o d i f i e   e n   f o n c t i o n   d e s   d o n n � e s   e n t r � e s   e t   l ' e n r e g i s t r e   e n   p d f   s u r   l e   b u r e a u . 
                   L e   f i c h i e r   a u r a   l e   n o m :   "   F a c t u r e _ 0 0 X X X . p d f "   a v e c   u n e   e s p a c e   d e v a n t   a j o u t � e   a u t o m a t i q u e m e n t   p a r   e x c e l   q u i   c o n c a t � n e   l e   n o m   d u   f i c h i e r   a v e c   l e   n o m   d e   l a   f e u i l l e   d e   c a l c u l .� ��� l .���� O  .��� k  -�� ��� l ����  �  �  � ��� l ����  � O Itell current application to my logToFile("Appel de l'application Excel.")   � ��� � t e l l   c u r r e n t   a p p l i c a t i o n   t o   m y   l o g T o F i l e ( " A p p e l   d e   l ' a p p l i c a t i o n   E x c e l . " )� ��� l ����  � F @tell current application to log("Appel de l'application Excel.")   � ��� � t e l l   c u r r e n t   a p p l i c a t i o n   t o   l o g ( " A p p e l   d e   l ' a p p l i c a t i o n   E x c e l . " )� ��� l ����  �  �  � ��� I )�
��	
�
 .aevtodocnull  �    alis� c  %��� o  !�� B0 cheminfichiermodelefactureexcel cheminFichierModeleFactureExcel� m  !$�
� 
psxf�	  � ��� l **����  �  �  � ��� l *����� O  *���� k  9��� ��� l 99����  �  �  � ��� l 99� ���   � > 8 Mise � jour du num�ro de facture (dans les 2 cellules!)   � ��� p   M i s e   �   j o u r   d u   n u m � r o   d e   f a c t u r e   ( d a n s   l e s   2   c e l l u l e s ! )� ��� r  9K��� o  9>���� 0 numerofacture numeroFacture� n      ��� 1  FJ��
�� 
DPVu� 4  >F���
�� 
ccel� m  BE�� ���  D 4� ��� r  L^��� o  LQ���� 0 numerofacture numeroFacture� n      ��� 1  Y]��
�� 
DPVu� 4  QY���
�� 
ccel� m  UX�� ���  B 1 1� ��� l __��������  ��  ��  � ��� l __������  � , & Mise � jour de la date d'intervention   � ��� L   M i s e   �   j o u r   d e   l a   d a t e   d ' i n t e r v e n t i o n� ��� r  _u��� n  _h��� 1  dh��
�� 
shdt� o  _d���� $0 dateintervention dateIntervention� n      ��� 1  pt��
�� 
DPVu� 4  hp���
�� 
ccel� m  lo�� ���  B 1 2� ��� l vv��������  ��  ��  � ��� l vv�� ��    #  Mise � jour du nom du client    � :   M i s e   �   j o u r   d u   n o m   d u   c l i e n t�  r  v� o  v{���� 0 	nomclient 	nomClient n       1  ����
�� 
DPVu 4  {���	
�� 
ccel	 m  �

 �  B 1 4  l ����������  ��  ��    l ������   ) # Mise � jour de l'adresse du client    � F   M i s e   �   j o u r   d e   l ' a d r e s s e   d u   c l i e n t  r  �� o  ������ .0 adresseclientformatee adresseClientFormatee n       1  ����
�� 
DPVu 4  ����
�� 
ccel m  �� �  B 1 5  l ����������  ��  ��    l ���� !��    2 , Mise � jour de la forme juridique du client   ! �"" X   M i s e   �   j o u r   d e   l a   f o r m e   j u r i d i q u e   d u   c l i e n t #$# r  ��%&% o  ������  0 formejuridique formeJuridique& n      '(' 1  ����
�� 
DPVu( 4  ����)
�� 
ccel) m  ��** �++  B 1 6$ ,-, l ����������  ��  ��  - ./. l ����01��  0 ) # Mise � jour du type d'intervention   1 �22 F   M i s e   �   j o u r   d u   t y p e   d ' i n t e r v e n t i o n/ 343 r  ��565 o  ������ $0 typeintervention typeIntervention6 n      787 1  ����
�� 
DPVu8 4  ����9
�� 
ccel9 m  ��:: �;;  A 2 14 <=< l ����������  ��  ��  = >?> l ����@A��  @ 0 * Mise � jour de la dur�e de l'intervention   A �BB T   M i s e   �   j o u r   d e   l a   d u r � e   d e   l ' i n t e r v e n t i o n? CDC r  ��EFE o  ������ &0 dureeintervention dureeInterventionF n      GHG 1  ����
�� 
DPVuH 4  ����I
�� 
ccelI m  ��JJ �KK  C 2 1D LML l ����������  ��  ��  M NON l ����PQ��  P , & Mise � jour du co�t de l'intervention   Q �RR L   M i s e   �   j o u r   d u   c o � t   d e   l ' i n t e r v e n t i o nO STS r  ��UVU ]  ��WXW o  ������ &0 dureeintervention dureeInterventionX m  ��YY @4      V n      Z[Z 1  ����
�� 
DPVu[ 4  ����\
�� 
ccel\ m  ��]] �^^  D 2 1T _��_ l ����������  ��  ��  ��  � n  *6`a` 4  /6��b
�� 
XwSHb m  25cc �dd  S h e e t 1a 1  */��
�� 
1172� + %worksheet "Sheet1" of active workbook   � �ee J w o r k s h e e t   " S h e e t 1 "   o f   a c t i v e   w o r k b o o k� fgf l ����������  ��  ��  g hih l ����jk��  j A ; Une fois la facture remplie, on l'enregistre au format PDF   k �ll v   U n e   f o i s   l a   f a c t u r e   r e m p l i e ,   o n   l ' e n r e g i s t r e   a u   f o r m a t   P D Fi mnm l ����op��  o ? 9 On donne un nom � la feuille (ex : "Facture_00" & "234")   p �qq r   O n   d o n n e   u n   n o m   �   l a   f e u i l l e   ( e x   :   " F a c t u r e _ 0 0 "   &   " 2 3 4 " )n rsr r  ��tut l ��v����v b  ��wxw o  ������ 20 prefixnomfichierfacture prefixNomFichierFacturex l ��y����y c  ��z{z o  ������ 0 numerofacture numeroFacture{ m  ����
�� 
ctxt��  ��  ��  ��  u n      |}| 1  ����
�� 
pnam} 1  ����
�� 
1107s ~~ l ��������  � P J On enregistre au format PDF en indiquant le chemin du fichier de la forme   � ��� �   O n   e n r e g i s t r e   a u   f o r m a t   P D F   e n   i n d i q u a n t   l e   c h e m i n   d u   f i c h i e r   d e   l a   f o r m e ��� l ��������  � L F chemin + extension (ex : "Macintosh HD:Users:bruno:Desktop:" & ".pdf"   � ��� �   c h e m i n   +   e x t e n s i o n   ( e x   :   " M a c i n t o s h   H D : U s e r s : b r u n o : D e s k t o p : "   &   " . p d f "� ��� l ��������  � ] W le fichier produit aura la forme : chemin + une espace + nom de la feuille + extension   � ��� �   l e   f i c h i e r   p r o d u i t   a u r a   l a   f o r m e   :   c h e m i n   +   u n e   e s p a c e   +   n o m   d e   l a   f e u i l l e   +   e x t e n s i o n� ��� l ��������  � B < (ex : "Macintosh HD:Users:bruno:Desktop: Facture_00234.pdf"   � ��� x   ( e x   :   " M a c i n t o s h   H D : U s e r s : b r u n o : D e s k t o p :   F a c t u r e _ 0 0 2 3 4 . p d f "� ��� l ��������  � G A Il faudra donc renommer le fichier pour enlever l'espace inutile   � ��� �   I l   f a u d r a   d o n c   r e n o m m e r   l e   f i c h i e r   p o u r   e n l e v e r   l ' e s p a c e   i n u t i l e� ��� O �!��� I  �����
�� .smXL1659null���   7 X128��  � ����
�� 
5016� l 	������ b  	��� o  	���� 40 chemindossiertempfacture cheminDossierTempFacture� o  ���� 20 extensionfichierfacture extensionFichierFacture��  ��  � �����
�� 
1813� m  ��
�� e188� .��  � 1  ���
�� 
1107� ��� l ""��������  ��  ��  � ��� l ""������  � C = On n'enregistre pas les modifications dans le fichier mod�le   � ��� z   O n   n ' e n r e g i s t r e   p a s   l e s   m o d i f i c a t i o n s   d a n s   l e   f i c h i e r   m o d � l e� ��� I "+�����
�� .aevtquitnull��� ��� null��  � �����
�� 
savo� m  &'��
�� boovfals��  � ���� l ,,��������  ��  ��  ��  � m  ��
                                                                                  XCEL  alis    �  Macintosh HD               �Z�H+   	>Microsoft Excel.app                                             	IR�wi        ����  	                Microsoft Office 2008     �Z�s      �wY     	>   U  EMacintosh HD:Applications: Microsoft Office 2008: Microsoft Excel.app   (  M i c r o s o f t   E x c e l . a p p    M a c i n t o s h   H D  6Applications/Microsoft Office 2008/Microsoft Excel.app  / ��  � # application "Microsoft Excel"   � ��� : a p p l i c a t i o n   " M i c r o s o f t   E x c e l "� ��� l //��������  ��  ��  � ��� l //��������  ��  ��  � ��� l //��������  ��  ��  � ��� l  //������  � ` Z************************* Renommage du fichier facture avec Finder ***********************   � ��� � * * * * * * * * * * * * * * * * * * * * * * * * *   R e n o m m a g e   d u   f i c h i e r   f a c t u r e   a v e c   F i n d e r   * * * * * * * * * * * * * * * * * * * * * * *� ��� l //��������  ��  ��  � ��� l  //������  � � �
         On renome le fichier facture correctement (Impossible de le faire depuis Excel)
         et on le d�place dans le dossier des factures.
            � ���6 
                   O n   r e n o m e   l e   f i c h i e r   f a c t u r e   c o r r e c t e m e n t   ( I m p o s s i b l e   d e   l e   f a i r e   d e p u i s   E x c e l ) 
                   e t   o n   l e   d � p l a c e   d a n s   l e   d o s s i e r   d e s   f a c t u r e s . 
                  � ��� O  /���� k  5��� ��� l 55��������  ��  ��  � ��� l 55������  � P Jtell current application to my logToFile("Appel de l'application Finder.")   � ��� � t e l l   c u r r e n t   a p p l i c a t i o n   t o   m y   l o g T o F i l e ( " A p p e l   d e   l ' a p p l i c a t i o n   F i n d e r . " )� ��� l 55������  � G Atell current application to log("Appel de l'application Finder.")   � ��� � t e l l   c u r r e n t   a p p l i c a t i o n   t o   l o g ( " A p p e l   d e   l ' a p p l i c a t i o n   F i n d e r . " )� ��� l 55��������  ��  ��  � ��� r  5@��� c  5>��� o  5:���� 40 cheminfichiertempfacture cheminFichierTempFacture� m  :=�
� 
alis� o      �~�~ 20 fichiertempfacturealias fichierTempFactureAlias� ��� l AA�}�|�{�}  �|  �{  � ��� l AA�z���z  � 9 3 On renome le fichier pour enlever l'espace inutile   � ��� f   O n   r e n o m e   l e   f i c h i e r   p o u r   e n l e v e r   l ' e s p a c e   i n u t i l e� ��� r  AJ��� o  AF�y�y &0 nomfichierfacture nomFichierFacture� l     ��x�w� n      ��� 1  GI�v
�v 
pnam� o  FG�u�u 20 fichiertempfacturealias fichierTempFactureAlias�x  �w  � ��� l KK�t�s�r�t  �s  �r  � ��� r  KZ��� c  KT��� o  KP�q�q .0 chemindossierfactures cheminDossierFactures� m  PS�p
�p 
alis� o      �o�o ,0 dossierfacturesalias dossierFacturesAlias� ��� l [[�n�m�l�n  �m  �l  � ��� Q  [����� k  ^o�� ��� l ^^�k���k  � T N Si on d�place un fichier, il faut mettre � jour la variable qui pointe dessus   � ��� �   S i   o n   d � p l a c e   u n   f i c h i e r ,   i l   f a u t   m e t t r e   �   j o u r   l a   v a r i a b l e   q u i   p o i n t e   d e s s u s� ��� l ^^�j �j    + % car sinon on ne peut plus l'utiliser    � J   c a r   s i n o n   o n   n e   p e u t   p l u s   l ' u t i l i s e r�  r  ^m I ^k�i
�i .coremoveobj        obj  o  ^_�h�h 20 fichiertempfacturealias fichierTempFactureAlias �g	�f
�g 
insh	 o  bg�e�e ,0 dossierfacturesalias dossierFacturesAlias�f   o      �d�d 0 	mynewfile 	myNewFile 

 l nn�c�b�a�c  �b  �a    l nn�`�`    ytell current application to my logToFile("Copie de " & cheminFichierTempFacture & " vers " & cheminDossierFactures & ".")    � � t e l l   c u r r e n t   a p p l i c a t i o n   t o   m y   l o g T o F i l e ( " C o p i e   d e   "   &   c h e m i n F i c h i e r T e m p F a c t u r e   &   "   v e r s   "   &   c h e m i n D o s s i e r F a c t u r e s   &   " . " )  l nn�_�_   v ptell current application to log("Copie de " & cheminFichierTempFacture & " vers " & cheminDossierFactures & ".")    � � t e l l   c u r r e n t   a p p l i c a t i o n   t o   l o g ( " C o p i e   d e   "   &   c h e m i n F i c h i e r T e m p F a c t u r e   &   "   v e r s   "   &   c h e m i n D o s s i e r F a c t u r e s   &   " . " ) �^ l nn�]�\�[�]  �\  �[  �^  � R      �Z�Y�X
�Z .ascrerr ****      � ****�Y  �X  � I w��W�V
�W .sysodisAaleR        TEXT b  w� b  w� b  w� b  w� b  w� !  m  wz"" �## B I m p o s s i b l e   d e   c o p i e r   l e   f i c h i e r :  ! l 	z$�U�T$ o  z�S�S 40 cheminfichiertempfacture cheminFichierTempFacture�U  �T   l 
��%�R�Q% o  ���P
�P 
ret �R  �Q   m  ��&& �'' j U n   f i c h i e r   p o r t a n t   l e   m � m e   n o m   e x i s t e   p e u t - � t r e   d � j � . l 
��(�O�N( o  ���M
�M 
ret �O  �N   m  ��)) �** B O p � r a t i o n   d e   d � p l a c e m e n t   a n n u l � e .�V  � +�L+ l ���K�J�I�K  �J  �I  �L  � m  /2,,�                                                                                  MACS  alis    t  Macintosh HD               �Z�H+     3
Finder.app                                                      &���j        ����  	                CoreServices    �Z�s      ��\       3   0   /  6Macintosh HD:System: Library: CoreServices: Finder.app   
 F i n d e r . a p p    M a c i n t o s h   H D  &System/Library/CoreServices/Finder.app  / ��  � -.- l ���H�G�F�H  �G  �F  . /0/ r  ��121 n ��343 I  ���E5�D�E 0 getdate getDate5 6�C6 o  ���B�B $0 dateintervention dateIntervention�C  �D  4  f  ��2 o      �A�A 80 dateinterventionformatlong dateInterventionFormatLong0 787 l ���@9:�@  9 [ Umy logToFile("Date de l'intervention au format long : " & dateInterventionFormatLong)   : �;; � m y   l o g T o F i l e ( " D a t e   d e   l ' i n t e r v e n t i o n   a u   f o r m a t   l o n g   :   "   &   d a t e I n t e r v e n t i o n F o r m a t L o n g )8 <=< l ���?>?�?  > R Llog("Date de l'intervention au format long : " & dateInterventionFormatLong)   ? �@@ � l o g ( " D a t e   d e   l ' i n t e r v e n t i o n   a u   f o r m a t   l o n g   :   "   &   d a t e I n t e r v e n t i o n F o r m a t L o n g )= ABA l ���>�=�<�>  �=  �<  B CDC l ���;�:�9�;  �:  �9  D EFE l ���8�7�6�8  �7  �6  F GHG l  ���5IJ�5  I [ U**************************** Envoi de la facture avec Mail **************************   J �KK � * * * * * * * * * * * * * * * * * * * * * * * * * * * *   E n v o i   d e   l a   f a c t u r e   a v e c   M a i l   * * * * * * * * * * * * * * * * * * * * * * * * * *H LML l ���4�3�2�4  �3  �2  M NON l  ���1PQ�1  P��
         Cr�e le message contenant la facture � envoyer au client.
         Impossible d'ajouter la signature avant la pi�ce jointe,
         sinon elle est supprim�e ou consid�r�e comme du texte
         et non plus comme une signature (elle est m�me soulign�e).
         
         Pour r�pondre � ce probl�me, j'ajoute la pi�ce jointe,
         j'attends 5 secondes et j'ajoute la signature � la fin
         (via l'interface graphique).
            Q �RR� 
                   C r � e   l e   m e s s a g e   c o n t e n a n t   l a   f a c t u r e   �   e n v o y e r   a u   c l i e n t . 
                   I m p o s s i b l e   d ' a j o u t e r   l a   s i g n a t u r e   a v a n t   l a   p i � c e   j o i n t e , 
                   s i n o n   e l l e   e s t   s u p p r i m � e   o u   c o n s i d � r � e   c o m m e   d u   t e x t e 
                   e t   n o n   p l u s   c o m m e   u n e   s i g n a t u r e   ( e l l e   e s t   m � m e   s o u l i g n � e ) . 
                   
                   P o u r   r � p o n d r e   �   c e   p r o b l � m e ,   j ' a j o u t e   l a   p i � c e   j o i n t e , 
                   j ' a t t e n d s   5   s e c o n d e s   e t   j ' a j o u t e   l a   s i g n a t u r e   �   l a   f i n 
                   ( v i a   l ' i n t e r f a c e   g r a p h i q u e ) . 
                  O STS O  ��UVU k  ��WW XYX l ���0�/�.�0  �/  �.  Y Z[Z I ���-�,�+
�- .miscactvnull��� ��� null�,  �+  [ \]\ l ���*�)�(�*  �)  �(  ] ^_^ l ���'`a�'  ` C =tell current application to my logToFile("Affichage de mail")   a �bb z t e l l   c u r r e n t   a p p l i c a t i o n   t o   m y   l o g T o F i l e ( " A f f i c h a g e   d e   m a i l " )_ cdc l ���&ef�&  e : 4tell current application to log("Affichage de mail")   f �gg h t e l l   c u r r e n t   a p p l i c a t i o n   t o   l o g ( " A f f i c h a g e   d e   m a i l " )d hih l ���%�$�#�%  �$  �#  i jkj r  ��lml b  ��non b  ��pqp o  ���"�" "0 contenumessage1 contenuMessage1q o  ���!�! 80 dateinterventionformatlong dateInterventionFormatLongo o  ��� �  "0 contenumessage2 contenuMessage2m o      ��  0 contenumessage contenuMessagek rsr r  ��tut I ����v
� .corecrel****      � null�  v �wx
� 
koclw m  ���
� 
bckex �y�
� 
prdty K  ��zz �{|
� 
sndr{ o  ���� (0 monadressecourrier monAdresseCourrier| �}~
� 
subj} o  ���� 0 monsujet monSujet~ ��
� 
pvis m  ���
� boovtrue� ���
� 
ctnt� o  ����  0 contenumessage contenuMessage�  �  u o      ��  0 nouveaumessage nouveauMessages ��� l ������  �  �  � ��� O  �I��� k  �H�� ��� l ����
�	�  �
  �	  � ��� I ����
� .corecrel****      � null�  � ���
� 
kocl� m   �
� 
rcpt� ���
� 
insh� n  ���  ;  � 2 �
� 
trcp� ���
� 
prdt� K  �� � ���
�  
radd� o  ���� .0 adressecourrierclient adresseCourrierClient��  �  � ��� l   ��������  ��  ��  � ��� l  >���� I  >�����
�� .corecrel****      � null��  � ����
�� 
kocl� m  $'��
�� 
atts� �����
�� 
prdt� K  *8�� �����
�� 
atfn� l -6������ n  -6��� 1  26��
�� 
psxp� o  -2���� (0 cheminfacturefinal cheminFactureFinal��  ��  ��  ��  � !  ins�r� � la fin du message   � ��� 6   i n s � r �   �   l a   f i n   d u   m e s s a g e� ��� I ?F�����
�� .sysodelanull��� ��� nmbr� m  ?B���� ��  � ���� l GG��������  ��  ��  ��  � o  ������  0 nouveaumessage nouveauMessage� ��� l JJ��������  ��  ��  � ��� I JO������
�� .miscactvnull��� ��� null��  ��  � ��� l PP��������  ��  ��  � ��� l  PP������  � - ' Ins�re la signature par GUI Scripting    � ��� N   I n s � r e   l a   s i g n a t u r e   p a r   G U I   S c r i p t i n g  � ���� O  P���� O  V���� k  a��� ��� l aa������  �  	delay 1.3   � ���  d e l a y   1 . 3� ��� O  a���� k  j��� ��� I jt�����
�� .prcsclicnull��� ��� uiel� 4  jp���
�� 
popB� m  no���� ��  � ���� I u������
�� .prcsclicnull��� ��� uiel� n  u���� 4  �����
�� 
menI� m  ������ � n  u���� 4  {����
�� 
menE� m  ~���� � 4  u{���
�� 
popB� m  yz���� ��  ��  � 4  ag���
�� 
cwin� m  ef���� � ���� l ����������  ��  ��  ��  � 4  V^���
�� 
prcs� m  Z]�� ���  M a i l� m  PS���                                                                                  sevs  alis    �  Macintosh HD               �Z�H+     3System Events.app                                               Ɉ���        ����  	                CoreServices    �Z�s      ���       3   0   /  =Macintosh HD:System: Library: CoreServices: System Events.app   $  S y s t e m   E v e n t s . a p p    M a c i n t o s h   H D  -System/Library/CoreServices/System Events.app   / ��  ��  V m  �����                                                                                  emal  alis    F  Macintosh HD               �Z�H+     UMail.app                                                        ���'�        ����  	                Applications    �Z�s      ��       U  #Macintosh HD:Applications: Mail.app     M a i l . a p p    M a c i n t o s h   H D  Applications/Mail.app   / ��  T ��� l ����������  ��  ��  � ��� l ����������  ��  ��  � ��� l ��������  � ' !display dialog "Fin du programme"   � ��� B d i s p l a y   d i a l o g   " F i n   d u   p r o g r a m m e "� ��� l ��������  �   create an alert   � ���     c r e a t e   a n   a l e r t� ��� l ��������  � | vset theAlert to current application's NSAlert's makeAlert:"An alert" buttons:{"Cancel", "OK"} |text|:(theDate as text)   � ��� � s e t   t h e A l e r t   t o   c u r r e n t   a p p l i c a t i o n ' s   N S A l e r t ' s   m a k e A l e r t : " A n   a l e r t "   b u t t o n s : { " C a n c e l " ,   " O K " }   | t e x t | : ( t h e D a t e   a s   t e x t )� ��� l ��������  � : 4theAlert's showModal() -- returns name of the button   � ��� h t h e A l e r t ' s   s h o w M o d a l ( )   - -   r e t u r n s   n a m e   o f   t h e   b u t t o n� ��� l ��������  �  
log result   � ���  l o g   r e s u l t� ��� l ����������  ��  ��  � ��� l ����������  ��  ��  � ��� l  ��������  � V P**************************** MAJ du num�ro de facture **************************   � �	 	  � * * * * * * * * * * * * * * * * * * * * * * * * * * * *   M A J   d u   n u m � r o   d e   f a c t u r e   * * * * * * * * * * * * * * * * * * * * * * * * * *� 			 l ����������  ��  ��  	 			 n ��			 I  ����	���� 0 	logtofile 	logToFile	 	��	 m  ��				 �	
	
 J - - - - >   M A J   d e s   v a r i a b l e s   d e   l a   f a c t u r e��  ��  	  f  ��	 			 I ����	��
�� .ascrcmnt****      � ****	 l ��	����	 m  ��		 �		 J - - - - >   M A J   d e s   v a r i a b l e s   d e   l a   f a c t u r e��  ��  ��  	 			 l ����������  ��  ��  	 			 l  ����		��  	 + % Incr�mentation du num�ro de facture    	 �		 J   I n c r � m e n t a t i o n   d u   n u m � r o   d e   f a c t u r e  	 			 l ����������  ��  ��  	 			 r  ��			 [  ��			 o  ������ 0 numerofacture numeroFacture	 m  ������ 	 o      ���� 0 numerofacture numeroFacture	 	 	!	  n ��	"	#	" I  ����	$���� 0 	logtofile 	logToFile	$ 	%��	% b  ��	&	'	& m  ��	(	( �	)	)  N u m � r o   :  	' o  ������ 0 numerofacture numeroFacture��  ��  	#  f  ��	! 	*	+	* I ����	,��
�� .ascrcmnt****      � ****	, l ��	-����	- b  ��	.	/	. m  ��	0	0 �	1	1  N u m � r o   :  	/ o  ������ 0 numerofacture numeroFacture��  ��  ��  	+ 	2	3	2 l ��	4	5	6	4 n ��	7	8	7 I  ����	9���� "0 setstringvalue_ setStringValue_	9 	:��	: o  ������ 0 numerofacture numeroFacture��  ��  	8 o  ������ &0 tfnumerodefacture tfNumeroDeFacture	5   MAJ de l'interface   	6 �	;	; &   M A J   d e   l ' i n t e r f a c e	3 	<	=	< l ����������  ��  ��  	= 	>	?	> l ����������  ��  ��  	? 	@	A	@ l  ����	B	C��  	B * $ Cr�ation du nom du fichier facture    	C �	D	D H   C r � a t i o n   d u   n o m   d u   f i c h i e r   f a c t u r e  	A 	E	F	E l ����������  ��  ��  	F 	G	H	G r  ��	I	J	I b  ��	K	L	K b  ��	M	N	M o  ������ 20 prefixnomfichierfacture prefixNomFichierFacture	N l ��	O����	O c  ��	P	Q	P o  ������ 0 numerofacture numeroFacture	Q m  ����
�� 
ctxt��  ��  	L o  ������ 20 extensionfichierfacture extensionFichierFacture	J o      ���� &0 nomfichierfacture nomFichierFacture	H 	R	S	R n �	T	U	T I  ���	V���� 0 	logtofile 	logToFile	V 	W�	W b  ��	X	Y	X m  ��	Z	Z �	[	[ " N o m   d u   f i c h i e r   :  	Y o  ���~�~ &0 nomfichierfacture nomFichierFacture�  ��  	U  f  ��	S 	\	]	\ I �}	^�|
�} .ascrcmnt****      � ****	^ l 	_�{�z	_ b  	`	a	` m  	b	b �	c	c " N o m   d u   f i c h i e r   :  	a o  �y�y &0 nomfichierfacture nomFichierFacture�{  �z  �|  	] 	d	e	d l �x�w�v�x  �w  �v  	e 	f	g	f l �u�t�s�u  �t  �s  	g 	h	i	h l  �r	j	k�r  	j C = Chemin vers le dossier temporaire de stockage de la facture    	k �	l	l z   C h e m i n   v e r s   l e   d o s s i e r   t e m p o r a i r e   d e   s t o c k a g e   d e   l a   f a c t u r e  	i 	m	n	m l �q�p�o�q  �p  �o  	n 	o	p	o r  '	q	r	q b  !	s	t	s b  	u	v	u o  �n�n 40 chemindossiertempfacture cheminDossierTempFacture	v 1  �m
�m 
spac	t o   �l�l &0 nomfichierfacture nomFichierFacture	r o      �k�k 40 cheminfichiertempfacture cheminFichierTempFacture	p 	w	x	w n (<	y	z	y I  )<�j	{�i�j 0 	logtofile 	logToFile	{ 	|�h	| c  )8	}	~	} b  )6		�	 m  ),	�	� �	�	� ( C h e m i n   t e m p o r a i r e   :  	� l ,5	��g�f	� n  ,5	�	�	� 1  15�e
�e 
psxp	� o  ,1�d�d 40 cheminfichiertempfacture cheminFichierTempFacture�g  �f  	~ m  67�c
�c 
ctxt�h  �i  	z  f  ()	x 	�	�	� I =P�b	��a
�b .ascrcmnt****      � ****	� l =L	��`�_	� c  =L	�	�	� b  =J	�	�	� m  =@	�	� �	�	� ( C h e m i n   t e m p o r a i r e   :  	� l @I	��^�]	� n  @I	�	�	� 1  EI�\
�\ 
psxp	� o  @E�[�[ 40 cheminfichiertempfacture cheminFichierTempFacture�^  �]  	� m  JK�Z
�Z 
ctxt�`  �_  �a  	� 	�	�	� l QQ�Y�X�W�Y  �X  �W  	� 	�	�	� l QQ�V�U�T�V  �U  �T  	� 	�	�	� l  QQ�S	�	��S  	� > 8 Chemin vers le dossier final de stockage de la facture    	� �	�	� p   C h e m i n   v e r s   l e   d o s s i e r   f i n a l   d e   s t o c k a g e   d e   l a   f a c t u r e  	� 	�	�	� l QQ�R�Q�P�R  �Q  �P  	� 	�	�	� r  Qb	�	�	� b  Q\	�	�	� o  QV�O�O .0 chemindossierfactures cheminDossierFactures	� o  V[�N�N &0 nomfichierfacture nomFichierFacture	� o      �M�M (0 cheminfacturefinal cheminFactureFinal	� 	�	�	� n cw	�	�	� I  dw�L	��K�L 0 	logtofile 	logToFile	� 	��J	� c  ds	�	�	� b  dq	�	�	� m  dg	�	� �	�	�  C h e m i n   f i n a l   :  	� l gp	��I�H	� n  gp	�	�	� 1  lp�G
�G 
psxp	� o  gl�F�F (0 cheminfacturefinal cheminFactureFinal�I  �H  	� m  qr�E
�E 
ctxt�J  �K  	�  f  cd	� 	�	�	� I x��D	��C
�D .ascrcmnt****      � ****	� l x�	��B�A	� c  x�	�	�	� b  x�	�	�	� m  x{	�	� �	�	�  C h e m i n   f i n a l   :  	� l {�	��@�?	� n  {�	�	�	� 1  ���>
�> 
psxp	� o  {��=�= (0 cheminfacturefinal cheminFactureFinal�@  �?  	� m  ���<
�< 
ctxt�B  �A  �C  	� 	�	�	� l ���;�:�9�;  �:  �9  	� 	�	�	� l ���8�7�6�8  �7  �6  	� 	�	�	� n ��	�	�	� I  ���5	��4�5 0 	logtofile 	logToFile	� 	��3	� m  ��	�	� �	�	� J < - - - -   M A J   d e s   v a r i a b l e s   d e   l a   f a c t u r e�3  �4  	�  f  ��	� 	�	�	� I ���2	��1
�2 .ascrcmnt****      � ****	� l ��	��0�/	� m  ��	�	� �	�	� J < - - - -   M A J   d e s   v a r i a b l e s   d e   l a   f a c t u r e�0  �/  �1  	� 	�	�	� l ���.�-�,�.  �-  �,  	� 	�	�	� n ��	�	�	� I  ���+	��*�+ 0 	logtofile 	logToFile	� 	��)	� m  ��	�	� �	�	� . < - - - -   N o u v e l l e   f a c t u r e 
�)  �*  	�  f  ��	� 	�	�	� I ���(	��'
�( .ascrcmnt****      � ****	� l ��	��&�%	� m  ��	�	� �	�	� . < - - - -   N o u v e l l e   f a c t u r e 
�&  �%  �'  	� 	��$	� l ���#�"�!�#  �"  �!  �$  � 	�	�	� l     � ���   �  �  	� 	�	�	� l     ����  �  �  	� 	�	�	� l     ����  �  �  	� 	�	�	� l     ����  �  �  	� 	�	�	� l     ����  �  �  	� 	�	�	� l     �	�	��  	�   shows modal alert   	� �	�	� $   s h o w s   m o d a l   a l e r t	� 	�	�	� i   � �	�	�	� I      �	��� "0 showalertmodal_ showAlertModal_	� 	��	� o      �� 
0 sender  �  �  	� k     	�	� 	�	�	� l     �	�	��  	�   create an alert   	� �	�	�     c r e a t e   a n   a l e r t	� 	�	�	� r     
 

  n    


 I    �
�
� 20 makealert_buttons_text_ makeAlert_buttons_text_
 


 m    

 �

  A n   a l e r t
 
	


	 J    

 


 m    

 �

  C a n c e l
 
�	
 m    

 �

  O K�	  

 
�
 m    	

 �

 & F u r t h e r   e x p l a n a t i o n�  �
  
 n    


 o    �� 0 nsalert NSAlert
 m     �
� misccura
 o      �� 0 thealert theAlert	� 


 l   



 n   


 I    ���� 0 	showmodal 	showModal�  �  
 o    �� 0 thealert theAlert
 !  returns name of the button   
 �

 6   r e t u r n s   n a m e   o f   t h e   b u t t o n
 
 � 
  I   ��
!��
�� .ascrcmnt****      � ****
! 1    ��
�� 
rslt��  �   	� 
"
#
" l     ��������  ��  ��  
# 
$
%
$ l     ��������  ��  ��  
% 
&
'
& l     ��������  ��  ��  
' 
(
)
( l     ��������  ��  ��  
) 
*
+
* i   � �
,
-
, I      ��
.����  0 datepickerget_ datePickerGet_
. 
/��
/ o      ���� 
0 sender  ��  ��  
- k     #
0
0 
1
2
1 l     ��������  ��  ��  
2 
3
4
3 r     
5
6
5 l    
7����
7 c     
8
9
8 n    	
:
;
: I    	�������� 0 dateas dateAS��  ��  
; o     ���� 0 
datepicker 
datePicker
9 m   	 
��
�� 
ldt ��  ��  
6 o      ���� 0 thedate theDate
4 
<
=
< I   ��
>��
�� .ascrcmnt****      � ****
> c    
?
@
? o    ���� 0 thedate theDate
@ m    ��
�� 
ctxt��  
= 
A
B
A I   ��
C��
�� .ascrcmnt****      � ****
C m    
D
D �
E
E  # #��  
B 
F
G
F I   !��
H��
�� .ascrcmnt****      � ****
H m    
I
I �
J
J  # #��  
G 
K��
K l  " "��������  ��  ��  ��  
+ 
L
M
L l     ��������  ��  ��  
M 
N
O
N l     ��������  ��  ��  
O 
P
Q
P l     ��������  ��  ��  
Q 
R
S
R l     ��������  ��  ��  
S 
T
U
T l     ��������  ��  ��  
U 
V
W
V l      ��
X
Y��  
X  �
     list_position : renvoie la position d'un �l�ment dans une liste.
     this_item : �l�ment � rechercher.
     this_list : liste dans laquelle on recherche.
     r�sultat : 0 si non trouv�, la position de l'�l�ment
     dans la liste sinon.
        
Y �
Z
Z� 
           l i s t _ p o s i t i o n   :   r e n v o i e   l a   p o s i t i o n   d ' u n   � l � m e n t   d a n s   u n e   l i s t e . 
           t h i s _ i t e m   :   � l � m e n t   �   r e c h e r c h e r . 
           t h i s _ l i s t   :   l i s t e   d a n s   l a q u e l l e   o n   r e c h e r c h e . 
           r � s u l t a t   :   0   s i   n o n   t r o u v � ,   l a   p o s i t i o n   d e   l ' � l � m e n t 
           d a n s   l a   l i s t e   s i n o n . 
          
W 
[
\
[ i   � �
]
^
] I      ��
_���� 0 list_position  
_ 
`
a
` o      ���� 0 	this_item  
a 
b��
b o      ���� 0 	this_list  ��  ��  
^ k     .
c
c 
d
e
d l     ��������  ��  ��  
e 
f
g
f q      
h
h ������ 0 position  ��  
g 
i
j
i r     
k
l
k m     ����  
l o      ���� 0 position  
j 
m
n
m l   ��������  ��  ��  
n 
o
p
o Y    )
q��
r
s��
q k    $
t
t 
u
v
u l   ��������  ��  ��  
v 
w
x
w Z    "
y
z����
y =   
{
|
{ n    
}
~
} 4    ��

�� 
cobj
 o    ���� 0 i  
~ o    ���� 0 	this_list  
| o    ���� 0 	this_item  
z r    
�
�
� o    ���� 0 i  
� o      ���� 0 position  ��  ��  
x 
���
� l  # #��������  ��  ��  ��  �� 0 i  
r m    ���� 
s l   
�����
� I   ��
���
�� .corecnte****       ****
� o    	���� 0 	this_list  ��  ��  ��  ��  
p 
�
�
� l  * *��������  ��  ��  
� 
�
�
� L   * ,
�
� o   * +���� 0 position  
� 
���
� l  - -��������  ��  ��  ��  
\ 
�
�
� l     ��������  ��  ��  
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
� V Ptell current application to my logToFile("Entr�e dans la fonction \"getDate\".")   
� �
�
� � t e l l   c u r r e n t   a p p l i c a t i o n   t o   m y   l o g T o F i l e ( " E n t r � e   d a n s   l a   f o n c t i o n   \ " g e t D a t e \ " . " )
� 
�
�
� l     ��
�
���  
� M Gtell current application to log("Entr�e dans la fonction \"getDate\".")   
� �
�
� � t e l l   c u r r e n t   a p p l i c a t i o n   t o   l o g ( " E n t r � e   d a n s   l a   f o n c t i o n   \ " g e t D a t e \ " . " )
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
� m     ��
�� misccura
� 
�
�
� l   ��~�}�  �~  �}  
� 
�
�
� r    
�
�
� c    
�
�
� o    �|�| 0 datecomplete dateComplete
� m    �{
�{ 
ctxt
� o      �z�z 0 
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
� 4  �y
�
�y 
cwor
� m    �x�x 
� o    �w�w 0 
moninstant 
monInstant
� o      �v�v 0 mon_jour_semaine_fr  
� 
�
�
� r     
�
�
� n    
�
�
� 4  �u
�
�u 
cwor
� m    �t�t 
� o    �s�s 0 
moninstant 
monInstant
� o      �r�r 0 mon_jour_fr  
� 
�
�
� r   ! '
�
�
� n   ! %
�
�
� 4 " %�q
�
�q 
cwor
� m   # $�p�p 
� o   ! "�o�o 0 
moninstant 
monInstant
� o      �n�n 0 mon_mois_fr  
� 
�
�
� r   ( .
�
�
� n   ( ,
�
�
� 4 ) ,�m
�
�m 
cwor
� m   * +�l�l 
� o   ( )�k�k 0 
moninstant 
monInstant
� o      �j�j 0 mon_annee_fr  
� 
�
�
� r   / 5
�
�
� n   / 3
�
�
� 4 0 3�i
�
�i 
cwor
� m   1 2�h�h 
� o   / 0�g�g 0 
moninstant 
monInstant
� o      �f�f 0 mon_heure_fr  
� 
�
�
� r   6 <
�
�
� n   6 :
�
�
� 4 7 :�e
�
�e 
cwor
� m   8 9�d�d 
� o   6 7�c�c 0 
moninstant 
monInstant
� o      �b�b 0 ma_minute_fr  
� 
�
�
� r   = C
�
�
� n   = A
�
�
� 4 > A�a
�
�a 
cwor
� m   ? @�`�` 
� o   = >�_�_ 0 
moninstant 
monInstant
� o      �^�^ 0 ma_seconde_fr  
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
� o   D E�]�] 0 mon_jour_semaine_fr  
� m   E F
�
� �
�
�   
� o   G H�\�\ 0 mon_jour_fr  
� m   I J
�
� �
�
�   
� o   K L�[�[ 0 mon_mois_fr  
� m   M N
�
� �
�
�   
� o   O P�Z�Z 0 mon_annee_fr  
� 
��Y
� l  S S�X�W�V�X  �W  �V  �Y  
� 
� 
� l     �U�T�S�U  �T  �S     l     �R�Q�P�R  �Q  �P    l     �O�N�M�O  �N  �M    l      �L�L  %
     getNumeroFacture : renvoie le num�ro de facture � partir du nom du dernier fichier
     dans le dossier des factures (Facture_00212.pdf).
     dossierFactures :  dossier contenant les fichiers de facture.
     r�sultat :         le num�ro de facture � utiliser pour facturer.
         �		> 
           g e t N u m e r o F a c t u r e   :   r e n v o i e   l e   n u m � r o   d e   f a c t u r e   �   p a r t i r   d u   n o m   d u   d e r n i e r   f i c h i e r 
           d a n s   l e   d o s s i e r   d e s   f a c t u r e s   ( F a c t u r e _ 0 0 2 1 2 . p d f ) . 
           d o s s i e r F a c t u r e s   :     d o s s i e r   c o n t e n a n t   l e s   f i c h i e r s   d e   f a c t u r e . 
           r � s u l t a t   :                   l e   n u m � r o   d e   f a c t u r e   �   u t i l i s e r   p o u r   f a c t u r e r . 
           

 i   � � I      �K�J�K $0 getnumerofacture getNumeroFacture �I o      �H�H "0 dossierfactures dossierFactures�I  �J   k     N  l     �G�F�E�G  �F  �E    r      J     �D�D   o      �C�C $0 listenomfichiers listeNomFichiers  l   �B�A�@�B  �A  �@    O    L k   	 K  r   	  !  n   	 "#" 2    �?
�? 
file# 4   	 �>$
�> 
cfol$ o    �=�= "0 dossierfactures dossierFactures! o      �<�< "0 touslesfichiers tousLesFichiers %&% X    .'�;(' k   " ))) *+* r   " ',-, l  " %.�:�9. n   " %/0/ 1   # %�8
�8 
pnam0 o   " #�7�7 0 	unfichier 	unFichier�:  �9  - o      �6�6 0 
nomfichier 
nomFichier+ 1�51 l  ( (�423�4  2 / )set end of listeNomFichiers to nomFichier   3 �44 R s e t   e n d   o f   l i s t e N o m F i c h i e r s   t o   n o m F i c h i e r�5  �; 0 	unfichier 	unFichier( o    �3�3 "0 touslesfichiers tousLesFichiers& 565 r   / @787 n   / :9:9 7 0 :�2;<
�2 
ctxt; m   4 6�1�1 < m   7 9�0�0 : o   / 0�/�/ 0 
nomfichier 
nomFichier8 o      �.�. 0 numerofacture numeroFacture6 =>= l  A A�-�,�+�-  �,  �+  > ?�*? L   A K@@ [   A JABA l  A HC�)�(C c   A HDED o   A F�'�' 0 numerofacture numeroFactureE m   F G�&
�& 
long�)  �(  B m   H I�%�% �*   m    FF�                                                                                  MACS  alis    t  Macintosh HD               �Z�H+     3
Finder.app                                                      &���j        ����  	                CoreServices    �Z�s      ��\       3   0   /  6Macintosh HD:System: Library: CoreServices: Finder.app   
 F i n d e r . a p p    M a c i n t o s h   H D  &System/Library/CoreServices/Finder.app  / ��   GHG l  M M�$�#�"�$  �#  �"  H I�!I l  M M� ���   �  �  �!   JKJ l     ����  �  �  K LML l     �NO�  N C =-------------------------------------------------------------   O �PP z - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -M QRQ l     ����  �  �  R STS l     ����  �  �  T UVU i   � �WXW I      �Y�� d0 0applicationshouldterminateafterlastwindowclosed_ 0applicationShouldTerminateAfterLastWindowClosed_Y Z�Z o      �� 0 nsapplication nsApplication�  �  X k     [[ \]\ l     ����  �  �  ] ^_^ O    `a` I    
�b�� 0 
terminate_  b c�
c  f    �
  �  a o     �	�	 0 nsapplication nsApplication_ d�d l   ����  �  �  �  V efe l     ����  �  �  f ghg l     �� ���  �   ��  h iji l      ��kl��  k � �
     logToFile :    �crit une ligne (en UTF-8) � la fin du fichier referenceVersFichierLog.
     ligne :        le texte � �crire dans le fichier.
     r�sultat :     Aucun.
        l �mmh 
           l o g T o F i l e   :         � c r i t   u n e   l i g n e   ( e n   U T F - 8 )   �   l a   f i n   d u   f i c h i e r   r e f e r e n c e V e r s F i c h i e r L o g . 
           l i g n e   :                 l e   t e x t e   �   � c r i r e   d a n s   l e   f i c h i e r . 
           r � s u l t a t   :           A u c u n . 
          j non i   � �pqp I      ��r���� 0 	logtofile 	logToFiler s��s o      ���� 	0 ligne  ��  ��  q k     +tt uvu l     ��������  ��  ��  v wxw r     yzy b     {|{ o     ���� 	0 ligne  | m    }} �~~  
z o      ���� 	0 ligne  x � l   ��������  ��  ��  � ��� Q    )���� I  	 ����
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
�� .rdwrclosnull���     ****� o     %���� 20 referenceversfichierlog referenceVersFichierLog��  � ���� l  * *��������  ��  ��  ��  o ��� l     ��������  ��  ��  � ���� l     ��������  ��  ��  ��  ��       ������  � ���� 0 appdelegate AppDelegate� �� "���� 0 appdelegate AppDelegate� �� �����
�� misccura
�� 
pcls� ���  N S O b j e c t� )����������������� R W \ a i t y ~ � �� � � � � � � � � � � �������������  � '������������������������������������������������������������������������������
�� 
pare
�� 
cwin�� 0 
datepicker 
datePicker�� 0 tfnomclient tfNomClient�� (0 tftypeintervention tfTypeIntervention�� *0 tfdureeintervention tfDureeIntervention�� &0 tfnumerodefacture tfNumeroDeFacture�� 20 prefixnomfichierfacture prefixNomFichierFacture�� 20 extensionfichierfacture extensionFichierFacture�� B0 cheminfichiermodelefactureexcel cheminFichierModeleFactureExcel�� .0 chemindossierfactures cheminDossierFactures�� 40 chemindossiertempfacture cheminDossierTempFacture�� (0 monadressecourrier monAdresseCourrier�� 0 masignature maSignature�� 0 monsujet monSujet�� "0 contenumessage1 contenuMessage1�� "0 contenumessage2 contenuMessage2�� 0 
fichierlog 
fichierLog�� 20 referenceversfichierlog referenceVersFichierLog�� 0 numerofacture numeroFacture�� 40 cheminfichiertempfacture cheminFichierTempFacture�� &0 nomfichierfacture nomFichierFacture�� (0 cheminfacturefinal cheminFactureFinal�� ,0 dossierfacturesalias dossierFacturesAlias�� .0 adressecourrierclient adresseCourrierClient�� 0 	nomclient 	nomClient�� $0 typeintervention typeIntervention�� &0 dureeintervention dureeIntervention�� $0 dateintervention dateIntervention�� B0 applicationwillfinishlaunching_ applicationWillFinishLaunching_�� :0 applicationshouldterminate_ applicationShouldTerminate_�� 0 clickbutton_ clickButton_�� "0 showalertmodal_ showAlertModal_��  0 datepickerget_ datePickerGet_�� 0 list_position  �� 0 getdate getDate�� $0 getnumerofacture getNumeroFacture�� d0 0applicationshouldterminateafterlastwindowclosed_ 0applicationShouldTerminateAfterLastWindowClosed_�� 0 	logtofile 	logToFile��  
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
�� 
msng� 'furlfile:///Users/bruno/Desktop/log.txt� �� ����������� B0 applicationwillfinishlaunching_ applicationWillFinishLaunching_�� ����� �  ���� 0 anotification aNotification��  � ���� 0 anotification aNotification�  ����	���� &����NV��v~����������������tz
�� 
perm
�� .rdwropenshor       file�� 0 	logtofile 	logToFile
�� .ascrcmnt****      � ****
�� 
alis�� $0 getnumerofacture getNumeroFacture
�� 
ctxt
�� 
spac
�� 
psxp
�� .misccurdldt    ��� null�� 0 
setdateas_ 
setDateAS_�� "0 setstringvalue_ setStringValue_���b  �el Ec  O)�k+ O�j O)�k+ O�j Ob  
�&Ec  O)b  k+ 	Ec  O)�b  %k+ O�b  %j Ob  b  �&%b  %Ec  O)�b  %k+ O�b  %j Ob  �%b  %Ec  O)a b  a ,%�&k+ Oa b  a ,%�&j Ob  
b  %Ec  O)a b  a ,%�&k+ Oa b  a ,%�&j Oa Ec  O)a b  %k+ Oa b  %j Oa Ec  O)a b  %k+ Oa b  %j Ob  *j k+ Ob  b  k+ O)a k+ Oa j OP� ������������� :0 applicationshouldterminate_ applicationShouldTerminate_�� ����� �  ���� 
0 sender  ��  � ���� 
0 sender  � �������������� 0 	logtofile 	logToFile
�� .ascrcmnt****      � ****
�� .rdwrclosnull���     ****
�� misccura��  0 nsterminatenow NSTerminateNow�� )�k+ O�j Ob  j O��,EOP� ������������� 0 clickbutton_ clickButton_�� ����� �  �� 
0 sender  ��  � �~�}�|�{�z�y�x�w�v�u�t�s�r�q�p�o�n�m�~ 
0 sender  �} *0 ficheclienttrouvees ficheclientTrouvees�| 0 ficheclient ficheClient�{ 0 nomsclients nomsClients�z 0 unclient unClient�y 0 listereponse listeReponse�x 0 pos  �w  0 formejuridique formeJuridique�v 0 contactemail contactEmail�u 0 	theemails 	theEmails�t 0 anemail anEmail�s 0 adresseclient adresseClient�r .0 adresseclientformatee adresseClientFormatee�q 20 fichiertempfacturealias fichierTempFactureAlias�p 0 	mynewfile 	myNewFile�o 80 dateinterventionformatlong dateInterventionFormatLong�n  0 contenumessage contenuMessage�m  0 nouveaumessage nouveauMessage� ���l��k�j�i����h��g�f�e<D]`�dcf�ck�bq�au�`x�_�^�]�\��[���Z��Y��X�W�V�U9C�TY]ku�S���������R��Q�P�O�NZd�M�L�K�J�I���H�G���������F�E;Cai�D�C����B�A�@�?c�>��=��<�
*:JY]�;�:�9�8�7�6�5,�4�3�2"&)�1�0��/�.�-�,�+�*�)�(�'�&�%�$�#�"�!� ���������				(	0	Z	b	�	�	�	�	�	�	�	��l 0 	logtofile 	logToFile
�k .ascrcmnt****      � ****�j 0 stringvalue stringValue
�i 
ctxt
�h 
azf4�  
�g 
pnam
�f .corecnte****       ****
�e misccura
�d 
ret 
�c 
appr
�b 
btns
�a 
cbtn
�` 
dflt�_ 
�^ .sysodlogaskr        TEXT
�] 
cobj
�\ 
kocl
�[ 
prmp
�Z 
okbt
�Y 
cnbt
�X 
mlsl�W 

�V .gtqpchltns    @   @ ns  �U 0 list_position  
�T 
az51
�S 
az21
�R 
az18
�Q 
az17
�P 
ascr
�O 
txdl
�N 
citm
�M 
az27
�L 
az28
�K 
spac
�J 
az31
�I 
az29�H  �G  
�F .aevtquitnull��� ��� null�E "0 setstringvalue_ setStringValue_�D 0 dateas dateAS
�C 
ldt 
�B 
psxf
�A .aevtodocnull  �    alis
�@ 
1172
�? 
XwSH
�> 
ccel
�= 
DPVu
�< 
shdt
�; 
1107
�: 
5016
�9 
1813
�8 e188� .�7 
�6 .smXL1659null���   7 X128
�5 
savo
�4 
alis
�3 
insh
�2 .coremoveobj        obj 
�1 .sysodisAaleR        TEXT�0 0 getdate getDate
�/ .miscactvnull��� ��� null
�. 
bcke
�- 
prdt
�, 
sndr
�+ 
subj
�* 
pvis
�) 
ctnt
�( .corecrel****      � null
�' 
rcpt
�& 
trcp
�% 
radd�$ 
�# 
atts
�" 
atfn
�! 
psxp�  
� .sysodelanull��� ��� nmbr
� 
prcs
� 
cwin
� 
popB
� .prcsclicnull��� ��� uiel
� 
menE
� 
menI���)�k+ O�j Ob  j+ �&Ec  O)�b  %k+ O�b  %j O�8*�-�[�,\Zb  @1E�O�j j  [� )�k+ UO� �j UOa b  %a %_ %a %_ %a %a a a a kva a a a a  OPY ��j k  �a k/E�OPY ��j k �jvE�O �[a  a l kh ��,�6F[OY��O�a a !a "a #b  %_ %a $%a %a &a 'a (a )fa * +E�O�a k/Ec  O)b  �l+ ,E�O�a �/E�OPY hO��,Ec  O� )a -b  %k+ UO� a .b  %j UO�a /,e  
a 0E�Y a 1E�O� )a 2�%k+ UO� a 3�%j UO�a 4-E�O�jv hY Aa 5b  %a 6%_ %a 7%_ %a 8%a a 9a a :kva a ;a a <a  OPO�j k qjvE�O *�[a  a l kh 
�a =,a >%�a ?,%�6F[OY��O�a "a @l +a k/E�Oa A_ Ba C,FO�a Di/Ec  Oa Ekv_ Ba C,FOPY "�j k  �a k/a ?,Ec  OPY hO� )a Fb  %k+ UO� a Gb  %j UO S�a H,a k/E�O� *a I,_ J%*a K,%_ J%*a L,%E�UO� )a M�%k+ UO� a N�%j UOPW FX O Pa Qb  %a R%_ %a S%_ %a T%a a Ua a Vkva a Wa a Xa  OPO*j YOPUOb  b  k+ ZOb  j+ �&Ec  O)a [b  %k+ Oa \b  %j Ob  j+ �&Ec  O)a ]b  %k+ Oa ^b  %j Ob  j+ _a `&Ec  O)a ab  %�&k+ Oa bb  %�&j Oa cb  	a d&j eO*a f,a ga h/ �b  *a ia j/a k,FOb  *a ia l/a k,FOb  a m,*a ia n/a k,FOb  *a ia o/a k,FO�*a ia p/a k,FO�*a ia q/a k,FOb  *a ia r/a k,FOb  *a ia s/a k,FOb  a t *a ia u/a k,FOPUOb  b  �&%*a v,�,FO*a v, *a wb  b  %a xa ya z {UO*a |fl YOPUOa } cb  a ~&E�Ob  ��,FOb  
a ~&Ec  O �a b  l �E�OPW $X O Pa �b  %_ %a �%_ %a �%j �OPUO)b  k+ �E�Oa � �*j �Ob  �%b  %E^ O*a  a �a �a �b  a �b  a �ea �] a a z �E^ O]  N*a  a �a *a �-6a �a �b  la � �O*a  a �a �a �b  a �,la z �Oa �j �OPUO*j �Oa � 9*a �a �/ -*a �k/ !*a �k/j �O*a �k/a �k/a �m/j �UOPUUUO)a �k+ Oa �j Ob  kEc  O)a �b  %k+ Oa �b  %j Ob  b  k+ ZOb  b  �&%b  %Ec  O)a �b  %k+ Oa �b  %j Ob  _ J%b  %Ec  O)a �b  a �,%�&k+ Oa �b  a �,%�&j Ob  
b  %Ec  O)a �b  a �,%�&k+ Oa �b  a �,%�&j O)a �k+ Oa �j O)a �k+ Oa �j OP� �	������� "0 showalertmodal_ showAlertModal_� ��� �  �� 
0 sender  �  � ��� 
0 sender  � 0 thealert theAlert� 
��



����
� misccura� 0 nsalert NSAlert� 20 makealert_buttons_text_ makeAlert_buttons_text_� 0 	showmodal 	showModal
� 
rslt
� .ascrcmnt****      � ****� ��,���lv�m+ E�O�j+ O�j 	� �

-�	�����
  0 datepickerget_ datePickerGet_�	 ��� �  �� 
0 sender  �  � ��� 
0 sender  � 0 thedate theDate� ��� ��
D
I� 0 dateas dateAS
� 
ldt 
�  
ctxt
�� .ascrcmnt****      � ****� $b  j+  �&E�O��&j O�j O�j OP� ��
^���������� 0 list_position  �� ����� �  ������ 0 	this_item  �� 0 	this_list  ��  � ���������� 0 	this_item  �� 0 	this_list  �� 0 position  �� 0 i  � ����
�� .corecnte****       ****
�� 
cobj�� /jE�O $k�j  kh ��/�  �E�Y hOP[OY��O�OP� ��
����������� 0 getdate getDate�� ����� �  ���� 0 datecomplete dateComplete��  � 	�������������������� 0 datecomplete dateComplete�� 0 
moninstant 
monInstant�� 0 mon_jour_semaine_fr  �� 0 mon_jour_fr  �� 0 mon_mois_fr  �� 0 mon_annee_fr  �� 0 mon_heure_fr  �� 0 ma_minute_fr  �� 0 ma_seconde_fr  � ����������������
�
�
�
�� misccura
�� 
ctxt
�� .ascrcmnt****      � ****
�� 
cwor�� �� �� �� �� U� 	��&j UO��&E�O��k/E�O��l/E�O��m/E�O���/E�O���/E�O���/E�O���/E�O��%�%�%�%�%�%OP� ������������ $0 getnumerofacture getNumeroFacture�� ����� �  ���� "0 dossierfactures dossierFactures��  � ������������ "0 dossierfactures dossierFactures�� $0 listenomfichiers listeNomFichiers�� "0 touslesfichiers tousLesFichiers�� 0 	unfichier 	unFichier�� 0 
nomfichier 
nomFichier� F��������������������
�� 
cfol
�� 
file
�� 
kocl
�� 
cobj
�� .corecnte****       ****
�� 
pnam
�� 
ctxt�� �� 
�� 
long�� OjvE�O� D*�/�-E�O �[��l kh ��,E�OP[OY��O�[�\[Z�\Z�2Ec  Ob  �&kUOP� ��X���������� d0 0applicationshouldterminateafterlastwindowclosed_ 0applicationShouldTerminateAfterLastWindowClosed_�� ����� �  ���� 0 nsapplication nsApplication��  � ���� 0 nsapplication nsApplication� ���� 0 
terminate_  �� � *)k+  UOP� ��q���������� 0 	logtofile 	logToFile�� ����� �  ���� 	0 ligne  ��  � ���� 	0 ligne  � }��������������������
�� 
refn
�� 
wrat
�� rdwreof 
�� 
as  
�� 
utf8�� 
�� .rdwrwritnull���     ****��  ��  
�� .rdwrclosnull���     ****�� ,��%E�O ��b  ����� W X  	b  j 
OPascr  ��ޭ