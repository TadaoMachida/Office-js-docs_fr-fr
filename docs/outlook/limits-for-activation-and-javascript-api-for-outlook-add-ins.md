
# Limites d’activation et d’API JavaScript des compléments Outlook

Pour offrir une expérience satisfaisante aux utilisateurs de compléments Outlook, il convient de connaître certaines recommandations relatives à l’activation et à l’utilisation de l’API afin d’implémenter vos compléments tout en respectant ces limites. Ces recommandations permettent de s’assurer que le traitement par Exchange Server ou Outlook des règles d’activation ou des appels à l’interface API JavaScript pour Office n’est pas anormalement long pour un complément particulier, ce qui aurait une incidence sur l’expérience globale de l’utilisateur dans Outlook et d’autres compléments. Ces limites s’appliquent à la conception de règles d’activation dans le manifeste du complément, ainsi qu’à l’utilisation de propriétés personnalisées, de paramètres d’itinérance, de destinataires, de demandes et réponses de Services Web Exchange (EWS) et d’appels asynchrones. 

 >**Remarque** Si votre complément s’exécute sur un client riche Outlook, vous devez également vérifier que le complément s’exécute dans certaines limites d’utilisation des ressources au moment de l’exécution. 


## Limites pour les règles d’activation


Suivez les recommandations ci-dessous lors de la création de règles d’activation pour les compléments Outlook :


- Limitez la taille du manifeste à 256 Ko. Vous ne pouvez pas installer le complément Outlook pour une boîte aux lettres Exchange si vous dépassez cette limite.

- Spécifiez jusqu’à quinze règles d’activation pour le complément. Vous ne pouvez pas installer le complément si vous dépassez cette limite.
    
- Si vous appliquez une règle [ItemHasKnownEntity](http://msdn.microsoft.com/en-us/library/87e10fd2-eab4-c8aa-bec3-dff92d004d39%28Office.15%29.aspx) au corps de l’élément sélectionné, attendez-vous à ce qu’un client riche Outlook n’applique la règle qu’au premier mégaoctet du corps et non au reste du corps compris au-delà de cette limite. Votre complément ne sera pas activé s’il existe uniquement des correspondances comprises au-delà du premier mégaoctet du corps. Si ce scénario vous semble probable, redéfinissez vos conditions relatives à l’activation.
    
- Si vous utilisez des expressions régulières dans les règles  **ItemHasKnownEntity** ou [ItemHasRegularExpressionMatch](http://msdn.microsoft.com/en-us/library/bfb726cd-81b0-a8d5-644f-2ca90a5273fc%28Office.15%29.aspx), tenez compte des recommandations et limites suivantes qui s’appliquent généralement à n’importe quel hôte Outlook, et celles décrites dans les tableaux 1, 2 et 3 qui diffèrent en fonction de l’hôte :
    
      - Spécifiez au maximum cinq expressions régulières dans les règles d’activation d’un complément. Vous ne pouvez pas installer un complément si vous dépassez cette limite.
    
  - Spécifiez des expressions régulières de sorte que les résultats que vous prévoyez d’obtenir soient renvoyés par l’appel de la méthode  **getRegExMatches** dans les 50 premières correspondances.
    
  - Peut spécifier des assertions avant dans les expressions régulières, mais pas d’assertions arrière, (?<=text), ni d’assertions arrière négatives, (?<!text).
    

Le tableau 1 répertorie les limites et décrit les différences de prise en charge des expressions régulières entre un client riche Outlook et Outlook Web App ou OWA pour les appareils. La prise en charge est indépendante de tout type spécifique d’appareil et de corps d’élément.


 **Tableau 1. Différences générales dans la prise en charge des expressions régulières**


|**Client riche Outlook**|**Outlook Web App ou OWA pour périphériques**|
|:-----|:-----|
|Utilise le moteur d’expression régulière C++ fourni dans le cadre de la bibliothèque de modèles standard Visual Studio. Ce moteur est conforme aux normes ECMAScript 5. |Utilise l’évaluation d’expression régulière incluse dans JavaScript ; celle-ci est fournie par le navigateur et prend en charge un sur-ensemble d’ECMAScript 5.|
|Étant donné que les moteurs regex sont différents, une expression régulière qui inclut une classe de caractères personnalisée basée sur des classes de caractères prédéfinies peut renvoyer des résultats différents dans le client riche Outlook par rapport à ceux d’Outlook Web App ou d’OWA pour les appareils.<br/><br/>Par exemple, l’expression régulière « [\s\S]{0,100} » donne un nombre, compris entre 0 et 100, de caractères uniques correspondant à un espace blanc ou à un espace non blanc. Dans un client riche Outlook, cette expression régulière renvoie des résultats différents de ceux obtenus dans Outlook Web App et OWA pour les appareils. Vous devez réécrire l’expression régulière « (\s|\S){0,100} » pour contourner ce problème. Cette expression régulière correspond à un nombre, compris entre 0 et 100, d’espace blanc ou d’espace non blanc.<br/><br/>Vous devez tester minutieusement chaque expression régulière sur chaque hôte Outlook et, si une expression régulière renvoie des résultats différents, la réécrire. |Vous devez tester minutieusement chaque expression régulière sur chaque hôte Outlook et, si une expression régulière renvoie des résultats différents, la réécrire.|
|Par défaut, limite l’évaluation de toutes les expressions régulières pour un complément à 1 seconde. Le dépassement de cette limite engendre une réévaluation à 3 reprises au maximum. Au-delà de la limite de réévaluation, un client riche Outlook désactive l’exécution du complément pour la même boîte aux lettres dans n’importe lequel des hôtes Outlook.<br/><br/>Les administrateurs peuvent remplacer ces limites d’évaluation à l’aide des clés de Registre  **OutlookActivationAlertThreshold** et **OutlookActivationManagerRetryLimit**.|Ne prend pas en charge les mêmes paramètres de surveillance des ressources ou du Registre que dans un client riche Outlook. Toutefois, les compléments avec expressions régulières qui requièrent un temps d’évaluation trop long sur un client riche Outlook sont désactivés pour la même boîte aux lettres sur tous les hôtes Outlook.|

Le tableau 2 répertorie les limites et décrit les différences dans la partie du corps d’élément auquel chaque hôte Outlook applique une expression régulière. Certaines de ces limites dépendent du type d’appareil et de corps d’élément, si l’expression régulière est appliquée sur le corps d’élément.

**Tableau 2. Limites de la taille du corps d’élément évalué**


||**Client riche Outlook**|**Outlook Web App, OWA pour périphériques,OWA pour iPad ou OWA pour iPhone**|**Outlook Web App**|
|:-----|:-----|:-----|:-----|
|Facteur de forme|Tout appareil pris en charge|Smartphones Android, iPad ou iPhone|Tout appareil pris en charge autre que les smartphones Android, l’iPad et l’iPhone|
|Corps d’élément en texte brut|Applique le regex sur le premier mégaoctet des données du corps, mais pas sur le reste du corps au-delà de cette limite.|Active le complément uniquement si le corps < 16 000 caractères.|Active le complément uniquement si le corps < 500 000 caractères.|
|Corps d’élément HTML|Applique le regex sur les premiers 512 Ko des données du corps, mais pas sur le reste du corps au-delà de cette limite. (Le nombre réel de caractères dépend de l’encodage qui peut varier de 1 à 4 octets par caractère.)|Applique l’expression régulière sur les 64 000 premiers caractères (y compris les caractères de balises HTML), mais pas sur le reste du corps au-delà de cette limite.|Active le complément uniquement si le corps < 500 000 caractères.|

Le tableau 3 répertorie les limites et décrit les différences dans les correspondances que chacun des hôtes Outlook renvoie après avoir évalué une expression régulière. La prise en charge est indépendante du type spécifique d’appareil, mais peut dépendre du type de corps d’élément, si l’expression régulière est appliquée sur le corps d’élément.

**Tableau 3. Limites sur les correspondances retournées**


||**Client riche Outlook**|**Outlook Web App ou OWA pour périphériques**|
|:-----|:-----|:-----|
|Ordre des correspondances renvoyées|Supposez que  **getRegExMatches** renvoie des correspondances pour la même expression régulière appliquée au même élément et que celles-ci sont différentes dans un client riche Outlook par rapport à dans Outlook Web App ou OWA pour périphériques.|Supposez que  **getRegExMatches** renvoie des correspondances dans un ordre différent dans un client riche Outlook par rapport à dans Outlook Web App ou OWA pour périphériques.|
|Corps d’élément en texte brut|**getRegExMatches** renvoie les correspondances comprenant 1 536 caractères maximum (1,5 Ko), pour un maximum de 50 correspondances.<br/><br/>**Remarque** : **getRegExMatches** ne renvoie pas de correspondances dans un ordre spécifique dans le tableau renvoyé. En général, partez du principe que l’ordre des correspondances dans un client riche Outlook pour la même expression régulière appliquée au même élément est différent de celui dans Outlook Web App et OWA pour les appareils.|**getRegExMatches** renvoie toute correspondance de 3 072 (3 Ko) caractères au maximum, pour un nombre maximal de 50 correspondances.|
|Corps d’élément HTML|**getRegExMatches** renvoie les correspondances comprenant 3 072 caractères maximum (3 Ko), pour un maximum de 50 correspondances.<br/> <br/> **Remarque** : **getRegExMatches** ne renvoie pas de correspondances dans un ordre spécifique dans le tableau renvoyé. En général, partez du principe que l’ordre des correspondances dans un client riche Outlook pour la même expression régulière appliquée au même élément est différent de celui dans Outlook Web App et OWA pour les appareils.|**getRegExMatches** renvoie toute correspondance de 3 072 (3 Ko) caractères au maximum, pour un nombre maximal de 50 correspondances.|

## Limites pour l’API JavaScript


À part les recommandations précédentes relatives aux règles d’activation, chacun des hôtes Outlook applique certaines limites dans le modèle objet JavaScript, comme indiqué dans le tableau 4 :


**Tableau 4 : Limites relatives à l’obtention ou à la définition de certaines données à l’aide de l’API JavaScript pour Office**


|**Fonctionnalité**|**Limite**|**API associées**|**Description**|
|:-----|:-----|:-----|:-----|
|Propriétés personnalisées|2 500 caractères|Objet [CustomProperties](../../reference/outlook/CustomProperties.md)<br/> <br/>Méthode [item.loadCustomPropertiesAsync](../../reference/outlook/Office.context.mailbox.item.md)|Limite pour toutes les propriétés personnalisées d’un élément de rendez-vous ou de message. Tous les hôtes Outlook renvoient une erreur si la taille totale de toutes les propriétés personnalisées d’un complément dépasse cette limite.|
|Paramètres d’itinérance|32 Ko de caractères|Objet [RoamingSettings](../../reference/outlook/RoamingSettings.md)<br/><br/> Propriété [context.roamingSettings](../../reference/outlook/Office.context.md)|Limite pour tous les paramètres d’itinérance du complément. Tous les hôtes Outlook renvoient une erreur si les paramètres dépassent cette limite.|
|Extraction des entités connues|2 000 caractères|Méthode [item.getEntities](../../reference/outlook/Office.context.mailbox.item.md)<br/> <br/>Méthode [item.getEntitiesByType](../../reference/outlook/Office.context.mailbox.item.md)<br/> <br/>Méthode [item.getFilteredEntitiesByName](../../reference/outlook/Office.context.mailbox.item.md)|Limite utilisée par le serveur Exchange Server pour extraire les entités connues sur le corps d’élément. Le serveur Exchange Server ignore les entités au-delà de cette limite. Notez que cette limite est indépendante du fait que le complément utilise une règle  **ItemHasKnownEntity**.|
|Services web Exchange|1 Mo de caractères|Méthode [mailbox.makeEwsRequestAsync](../../reference/outlook/Office.context.mailbox.md)|Limite pour une demande ou une réponse à un appel  **Mailbox.makeEwsRequestAsync**.|
|Destinataires|100 destinataires|Propriété [item.requiredAttendees](../../reference/outlook/Office.context.mailbox.item.md)<br/> <br/>Propriété [item.optionalAttendees](../../reference/outlook/Office.context.mailbox.item.md)<br/> <br/>Propriété [item.resources](../../reference/outlook/Office.context.mailbox.item.md)<br/> <br/>Propriété [item.to](../../reference/outlook/Office.context.mailbox.item.md)<br/> <br/>Propriété [item.cc](../../reference/outlook/Office.context.mailbox.item.md)<br/> <br/>Méthode [Recipients.addAsync](../../reference/outlook/Recipients.md)<br/> <br/>Méthode [Recipient.getAsync](../../reference/outlook/Recipients.md)<br/> <br/>Méthode [Recipient.setAsync](../../reference/outlook/Recipients.md)|Limites pour les destinataires spécifiés dans chaque propriété.|
|Nom d’affichage|255 caractères|Propriété [EmailAddressDetails.displayName](../../reference/outlook/simple-types.md)<br/><br/> Objet [Recipients](../../reference/outlook/Recipients.md)<br/><br/> Propriété **item.requiredAttendees**<br/><br/> Propriété **item.optionalAttendees** <br/><br/>Propriété **item.resources** <br/><br/>Propriété **item.to** <br/><br/>Propriété **item.cc**|Limite pour la longueur du nom d’affichage d’un rendez-vous ou d’un message.|
|Définition de l’objet|255 caractères|Méthode [mailbox.displayNewAppointmentForm](../../reference/outlook/Office.context.mailbox.md)<br/><br/> Méthode [Subject.setAsync](../../reference/outlook/Subject.md)|Limite pour l’objet du nouveau formulaire de rendez-vous ou limite pour la définition de l’objet d’un rendez-vous ou d’un message.|
|Définition de l’emplacement|255 caractères|Méthode [Location.setAsync](../../reference/outlook/Location.md)|Limite pour la définition de l’emplacement d’un rendez-vous ou d’une demande de réunion.|
|Corps dans un nouveau formulaire de rendez-vous|32 Ko de caractères|Méthode **Mailbox.displayNewAppointmentForm**|Limite pour le corps dans un nouveau formulaire de rendez-vous.|
|Affichage du corps d’un élément existant|32 Ko de caractères|Méthode [mailbox.displayAppointmentForm](../../reference/outlook/Office.context.mailbox.md)<br/><br/> Méthode [mailbox.displayMessageForm](../../reference/outlook/Office.context.mailbox.md)|Pour Outlook Web App et OWA pour périphériques : limite pour le corps dans un formulaire de rendez-vous ou de message existant.|
|Définition du corps|1 Mo de caractères|Méthode [Body.prependAsync](../../reference/outlook/Body.md)<br/> <br/>[Body.setAsync](../../reference/outlook/Body.md)<br/><br/>Méthode [Body.setSelectedDataAsync](../../reference/outlook/Body.md)|Limite pour la définition du corps d’un élément de rendez-vous ou de message.|
|Nombre de pièces jointes|499 fichiers dans Outlook Web App et OWA pour périphériques|Méthode [item.addFileAttachmentAsync](../../reference/outlook/Office.context.mailbox.item.md)|Limite du nombre de fichiers pouvant être joints à un élément à envoyer. Outlook Web App et OWA pour périphériques limitent généralement ce nombre à 499 fichiers, par l’intermédiaire de l’interface utilisateur et de  **addFileAttachmentAsync**. Un client riche Outlook ne limite pas spécifiquement le nombre de pièces jointes. Cependant, tous les hôtes Outlook respectent la limite de taille des pièces jointes définie sur le serveur Exchange Server de l’utilisateur. Reportez-vous à la ligne suivante pour connaître la taille des pièces jointes.|
|Taille des pièces jointes|En fonction du serveur Exchange Server|Méthode **item.addFileAttachmentAsync**|La taille de toutes les pièces jointes d’un élément est limitée ; l’administrateur peut définir cette limite sur le serveur Exchange Server de la boîte aux lettres de l’utilisateur. Pour un client riche Outlook, limite le nombre de pièces jointes d’un élément. Pour Outlook Web App et OWA pour les appareils, la limite la plus restrictive (entre le nombre de pièces jointes et la taille des pièces jointes) restreint les pièces jointes réelles d’un élément.|
|Nom de fichier des pièces jointes|255 caractères|Méthode **item.addFileAttachmentAsync**|Limite pour la longueur du nom de fichier d’une pièce jointe à ajouter à un élément.|
|URI des pièces jointes|2 048 caractères|Méthode **item.addFileAttachmentAsync**|Limite pour l’URI du nom de fichier à ajouter en tant que pièce jointe à un élément.|
|ID des pièces jointes|100 caractères|Méthode [item.addItemAttachmentAsync](../../reference/outlook/Office.context.mailbox.item.md)<br/><br/> Méthode [item.removeAttachmentAsync](../../reference/outlook/Office.context.mailbox.item.md)|Limite pour la longueur de l’ID de la pièce jointe à ajouter à un élément ou à supprimer d’un élément.|
|Appels asynchrones|3 appels|Méthode **item.addFileAttachmentAsync**<br/><br/>Méthode **item.addItemAttachmentAsync**<br/><br/><br/>Méthode **item.removeAttachmentAsync**<br/><br/> Méthode [Body.getTypeAsync](../../reference/outlook/Body.md)<br/><br/>Méthode **Body.prependAsync**<br/><br/>Méthode **Body.setSelectedDataAsync**<br/><br/> Méthode [CustomProperties.saveAsync](../../reference/outlook/CustomProperties.md)<br/><br/><br/> Méthode [item.LoadCustomPropertiesAysnc](../../reference/outlook/Office.context.mailbox.item.md)<br/><br/><br/> Méthode [Location.getAsync](../../reference/outlook/Location.md)<br/><br/>Méthode **Location.setAsync**<br/><br/> Méthode [mailbox.getCallbackTokenAsync](../../reference/outlook/Office.context.mailbox.md)<br/><br/> Méthode [mailbox.getUserIdentityTokenAsync](../../reference/outlook/Office.context.mailbox.md)<br/><br/> Méthode [mailbox.makeEwsRequestAsync](../../reference/outlook/Office.context.mailbox.md)<br/><br/>Méthode **Recipients.addAsync**<br/><br/> Méthode [Recipients.getAsync](../../reference/outlook/Recipients.md)<br/><br/>Méthode **Recipients.setAsync**<br/><br/> Méthode [RoamingSettings.saveAsync](../../reference/outlook/RoamingSettings.md)<br/><br/> Méthode [Subject.getAsync](../../reference/outlook/Subject.md)<br/><br/>Méthode **Subject.setAsync**<br/><br/> Méthode [Time.getAsync](../../reference/outlook/Time.md)<br/><br/> Méthode [Time.setAsync](../../reference/outlook/Time.md)|Pour Outlook Web App ou OWA pour périphériques : limite du nombre d’appels asynchrones simultanés, car les navigateurs autorisent uniquement un nombre limité d’appels asynchrones aux serveurs. |

## Ressources supplémentaires



- [Déployer et installer des compléments Outlook à des fins de test](../outlook/testing-and-tips.md)
    
- [Confidentialité, autorisations et sécurité pour les compléments Outlook](../outlook/../../docs/develop/privacy-and-security.md)
    
