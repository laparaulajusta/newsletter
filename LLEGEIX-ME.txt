============================================================
  NEWSLETTERS HTML - PROCEDIMENT
============================================================

No cal instal·lar res. Només cal: Windows i l'Outlook
d'escriptori (o, si no el tens, Edge i Outlook web).

Carpetes:
  generar.bat       -> el que fas servir per generar
  newsletters\      -> una .txt per cada newsletter (contingut)
  plantilla\        -> disseny (NO tocar, només el responsable)


------------------------------------------------------------
  FER UNA NEWSLETTER NOVA
------------------------------------------------------------

1. A la carpeta "newsletters", copia una newsletter (.txt)
   semblant a la que vols fer i posa-li un nom nou,
   per exemple: 2027-01-assemblea.txt

2. Obre-la amb el Bloc de notes i canvia l'assumpte, els
   textos, els enllaços i les imatges. Pots afegir, treure o
   reordenar blocs. Desa.

3. Arrossega el fitxer .txt a sobre de "generar.bat".

4. S'obre un correu nou a l'Outlook amb la newsletter a dins
   i l'assumpte posat. Si a la finestra negra surten AVISOS,
   tanca el correu, corregeix el .txt i torna al pas 3.

5. Envia-te-la primer a tu mateix i mira-la a l'Outlook i al
   mòbil. Si està bé, torna a generar-la, posa-hi els
   destinataris i envia-la.

Si l'ordinador no té l'Outlook d'escriptori, al pas 4 s'obre
una finestra nova d'Edge amb dues pestanyes: la newsletter i
un correu nou d'Outlook web. A la newsletter fes Ctrl+A i
Ctrl+C; al correu, clic al cos del missatge i Ctrl+V.


------------------------------------------------------------
  COM S'ESCRIU EL FITXER .TXT
------------------------------------------------------------

- Els camps van en MAJÚSCULES seguits de dos punts:
      TITOL: Benvolgudes i benvolguts,
- Un text pot ocupar diverses línies. Una línia en blanc fa
  un salt de paràgraf.
- Cada bloc comença amb una línia que només conté:  ---
- Les línies que comencen per # són notes i no surten.
- Negreta: <strong>text en negreta</strong>
- No comencis mai una línia de text amb una paraula en
  MAJÚSCULES seguida de dos punts (es confondria amb un camp).
- Un camp opcional que no vols, simplement no l'escriguis
  (o deixa'l buit): aquella part no surt.

Capçalera (abans del primer ---):
  ASSUMPTE    assumpte del correu
  PER         destinataris (opcional, separats per ;)

El logo, l'eslògan, el botó "Més notícies" i el peu hi són
sempre: no cal escriure'ls.

Blocs (línia BLOC:) i els seus camps. (opc.) = opcional

  text           TITOL (opc.), TEXT
                 Títol vermell i text justificat.

  text-centrat   TITOL (opc.), TEXT
                 Igual, però centrat.

  text-doble     TITOL1 (opc.), TEXT1, TITOL2 (opc.), TEXT2
                 Dues columnes de text.

  imatge         CAPCALERA (opc.), IMATGE, ENLLAC (opc.),
                 DESCRIPCIO (opc.)
                 Fons vermell, una imatge gran. CAPCALERA és
                 el text blanc de sobre (p. ex. "Clica per
                 veure el vídeo"). ENLLAC fa la imatge clicable.

  imatge-doble   CAPCALERA (opc.),
                 IMATGE1, ENLLAC1 (opc.), DESCRIPCIO1 (opc.),
                 TEXT1 (opc.),
                 IMATGE2, ENLLAC2 (opc.), DESCRIPCIO2 (opc.),
                 TEXT2 (opc.)
                 Fons vermell, dues imatges una al costat de
                 l'altra, amb text blanc opcional a sota.

  boto           TEXT, ENLLAC
                 Un botó vermell addicional.

  eslogan	 Imatge amb l'eslògan de campanya.

DESCRIPCIO és el text que surt si la imatge no es carrega.
Hi ha exemples complets a la carpeta "newsletters".


------------------------------------------------------------
  IMATGES (IMGUR)
------------------------------------------------------------

- Puja-les a imgur amb el compte de l'equip (les imatges
  pujades sense compte es poden acabar esborrant).
- Fes clic dret a la imatge > "Copia l'adreça de la imatge".
  L'enllaç ha de començar per https://i.imgur.com/ i acabar
  en .jpg, .jpeg o .png. L'enllaç de la pàgina
  (https://imgur.com/xxxx) NO serveix.
- RETALLA-LES A AQUESTA MIDA ABANS DE PUJAR-LES. L'Outlook
  d'escriptori no retalla imatges: si la proporció no és
  aquesta, la imatge sortirà més alta o més baixa.
      bloc imatge         1000 x 700 px
      bloc imatge-doble    460 x 700 px (les dues iguals)


------------------------------------------------------------
  DIFERÈNCIES QUE SÓN NORMALS
------------------------------------------------------------

- A l'Outlook d'escriptori les cantonades dels botons, de les
  fotos i dels cercles de les icones del peu surten rectes.
  A Outlook web, Gmail i el mòbil surten arrodonides.
- Mode fosc: cada programa (Outlook, Outlook web, Gmail,
  mòbil) canvia els colors a la seva manera a qui el té
  activat. No es pot controlar des del correu. Per revisar el
  disseny real, mira'l amb el mode clar (a l'Outlook, botó
  del sol/lluna "Canvia el fons" del missatge).


------------------------------------------------------------
  PROBLEMES HABITUALS
------------------------------------------------------------

- ERROR "el bloc ... no existeix": la línia BLOC: està mal
  escrita. El missatge diu quins tipus hi ha.
- AVÍS "falta el camp ...": és un camp obligatori i no hi és
  o està mal escrit; a la newsletter quedarà buit.
- AVÍS "el camp ... no existeix en aquest bloc": el nom del
  camp està mal escrit o no és d'aquest tipus de bloc.
- Els accents surten malament: desa el .txt en UTF-8
  (Bloc de notes > Desa com > Codificació: UTF-8).
- Windows no deixa executar generar.bat: fes clic dret al
  .zip original > Propietats > marca "Desbloqueja" > D'acord,
  i torna'l a descomprimir.
- Surt "No s'ha pogut obrir l'Outlook d'escriptori": no està
  instal·lat o no respon. Tanca'l i torna-ho a provar, o fes
  servir l'opció d'Outlook web que s'obre automàticament.


------------------------------------------------------------
  PER AL RESPONSABLE (canvis de disseny)
------------------------------------------------------------

- generar.ps1, a dalt de tot: $ObrirAmb = 'escriptori' o 'web'
  ($UrlNouCorreu és l'adreça d'Outlook web).
- plantilla\base.html: logo, eslògan, botó "Més notícies" i
  peu. Ha de contenir la marca {{CONTINGUT}}, on van els blocs.
- plantilla\blocs\*.html: un fitxer per tipus de bloc. El nom
  del fitxer és el que s'escriu a BLOC:. Per crear-ne un de
  nou, copia'n un i canvia'l. No cal tocar l'script.
- Marques a les plantilles (majúscules, sense accents):
    {{CAMP}}     camp obligatori (avisa si falta)
    {{CAMP?}}    camp opcional (queda buit si falta)
    <!--SI:CAMP--> ... <!--/SI:CAMP-->
                 tros que només surt si CAMP té valor
- HTML "antic": taules, estils en línia, amplades com a
  atributs, Arial primer a la llista de fonts, i els estils
  del text repetits en un <span> dins de la cel·la.
  No escriguis marques {{...}} dins de comentaris.
