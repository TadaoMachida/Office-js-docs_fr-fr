
# Résoudre les problèmes d’activation des compléments Outlook


L’activation des compléments Outlook est contextuelle et basée sur les règles d’activation du manifeste du complément. Quand les conditions de l’élément actuellement sélectionné satisfont aux règles d’activation du complément, l’application hôte s’active et affiche le bouton du complément dans l’interface utilisateur d’Outlook (dans le volet de sélection de complément pour les compléments de composition et dans la barre de complément pour les compléments de lecture). Toutefois, si votre complément ne s’active pas comme prévu, essayez d’en déterminer les raisons à partir des points suivants.

<a name="troubleshootingmailapps"></a>
## Est-ce que la boîte aux lettres utilisateur se trouve sur une version d’Exchange Server correspondant au minimum à Exchange 2013 ?


En premier lieu, assurez-vous que le compte de messagerie utilisateur que vous employez pour le test se trouve sur une version d’Exchange Server correspondant au minimum à Exchange 2013. Si vous utilisez des fonctionnalités spécifiques ultérieures à Exchange 2013, assurez-vous que le compte utilisateur se trouve sur une version appropriée d’Exchange.

Vous pouvez vérifier la version d’Exchange 2013 en adoptant l’une des approches suivantes :


- Renseignez-vous auprès de votre administrateur Exchange Server.
    
- Si vous testez le complément sur Outlook Web App ou OWA pour les appareils, dans un débogueur de script (par exemple le débogueur JScript disponible avec Internet Explorer), recherchez l’attribut  **src** de la balise **script** qui spécifie l’emplacement à partir duquel les scripts sont chargés. Le chemin d’accès doit contenir une sous-chaîne **owa/15.0.516.x/owa2/...**, où  **15.0.516.x** représente la version du serveur Exchange Server (par exemple **15.0.516.2**).
    
- Vous pouvez également utiliser la propriété [Office.context.mailbox.diagnostics.hostVersion](../../reference/outlook/Office.context.mailbox.diagnostics.md) pour vérifier la version. Dans Outlook Web App et OWA pour périphériques, cette propriété renvoie la version du serveur Exchange Server.
    
- Si vous pouvez tester le complément sur Outlook, servez-vous de cette technique de débogage simple, qui fait appel au modèle objet Outlook et à Visual Basic Editor :
    
      1. Tout d’abord, assurez-vous que les macros sont activées pour Outlook. Choisissez **Fichier**, **Options**, **Centre de gestion de la confidentialité**, **Paramètres du Centre de gestion de la confidentialité**, **Paramètres des macros**. Assurez-vous que l’option **Notifications pour toutes les macros** est sélectionnée dans le Centre de gestion de la confidentialité. Vous devez également avoir sélectionné **Activer les macros** au cours du démarrage d’Outlook.
    
      2. Sous l’onglet **Développeur** du ruban, choisissez **Visual Basic**.
    
     >  **Note**  Not seeing the  **Developer** tab? See [How to: Show the Developer Tab on the Ribbon](http://msdn.microsoft.com/en-us/library/ce7cb641-44f2-4a40-867e-a7d88f8e98a9%28Office.15%29.aspx) to turn it on. 3. Dans Visual Basic Editor, choisissez  **Affichage**,  **Fenêtre exécution**.
    
      4. Tapez ce qui suit dans la fenêtre Exécution pour afficher la version du serveur Exchange Server. La version principale de la valeur retournée doit être égale ou supérieure à 15.
    
        - S’il n’y a qu’un seul compte Exchange dans le profil de l’utilisateur :
        
            
            ?Session.ExchangeMailboxServerVersion
            
        
        - S’il y a plusieurs comptes Exchange dans le même profil utilisateur :
        
            
            ?Session.Accounts.Item(emailAddress).ExchangeMailboxServerVersion
         
        
        - _emailAddress_ représente une chaîne qui contient l’adresse STMP principale de l’utilisateur. Par exemple, si l’adresse STMP principale de l’utilisateur est randy@contoso.com, tapez ce qui suit :
        
            
            ?Session.Accounts.Item("randy@contoso.com").ExchangeMailboxServerVersion
        


## Le complément est-il désactivé ?


N’importe lequel des clients riches Outlook peut désactiver un complément pour des raisons de performances, notamment en cas de dépassement des seuils suivants : utilisation de l’UC ou de la mémoire, tolérance des incidents et durée nécessaire au traitement de toutes les expressions régulières pour un complément. Quand cela se produit, le client riche Outlook affiche une notification pour indiquer qu’il désactive le complément. 


 >**Remarque**  Seuls les clients riches Outlook analysent l’utilisation des ressources. Toutefois, la désactivation d’un complément dans un client riche Outlook entraîne également la désactivation du complément dans Outlook Web App et OWA pour périphériques.

Utilisez l’une des approches suivantes pour vérifier si un complément est désactivé : 


- Dans Outlook Web App, connectez-vous directement au compte de messagerie, choisissez l’icône Paramètres, puis choisissez  **Gérer les compléments** afin d’accéder au Centre d’administration Exchange, où vous pouvez vérifier si le complément est activé.
    
- Dans Outlook, accédez au mode Backstage, puis choisissez  **Gérer les compléments**. Connectez-vous au Centre d’administration Exchange pour vérifier si le complément est activé.
    
- Dans Outlook pour Mac, choisissez  **Gérer les compléments** dans la barre du complément. Connectez-vous au Centre d’administration Exchange pour vérifier si le complément est activé.
    

## Les éléments testés prennent-ils en charge les compléments Outlook et sont-ils remis par une version d’Exchange Server correspondant au minimum à Exchange 2013 ?


Si votre complément Outlook est un complément de lecture et qu’il est censé être activé lorsque l’utilisateur affiche un message (messages électroniques, demandes de réunion, réponses et annulations) ou un rendez-vous, et même si ces éléments prennent généralement en charge les compléments, il existe certaines exceptions quand l’élément sélectionné est :


- protégé par la Gestion des droits relatifs à l’information (IRM) ;
    
- au format S/MIME ou chiffré par d’autres moyens de protection ;
    
- un brouillon (aucun expéditeur n’est affecté) ou est situé dans le dossier Brouillons d’Outlook ;
    
- situé dans le dossier Courrier indésirable ;
    
- un rapport ou une notification de remise qui a la classe de message IPM.Report.* (notamment les rapports de remise et les notifications d’échec de remise, ainsi que les notifications de lecture, de non-lecture et de retard) ;
    
- fichier .msg joint à un autre message ou ouvert à partir du système de fichiers.
    
En outre, les rendez-vous étant toujours enregistrés au format RTF, une règle [ItemHasRegularExpressionMatch](http://msdn.microsoft.com/en-us/library/bfb726cd-81b0-a8d5-644f-2ca90a5273fc%28Office.15%29.aspx) qui spécifie une valeur **PropertyName** de **BodyAsHTML** n’active pas de complément pour un rendez-vous ou un message enregistré au format texte brut ou RTF.

Même si un élément de messagerie ne correspond pas à l’un des types ci-dessus, si cet élément n’a pas été remis par une version d’Exchange Server correspondant au minimum à Exchange 2013, les entités et les propriétés connues telles que l’adresse SMTP de l’expéditeur ne sont pas identifiées pour l’élément. Les règles d’activation qui dépendent de ces entités ou propriétés ne sont pas satisfaites et le complément n’est pas activé.

Si votre complément est un complément de composition et qu’il est censé être activé lorsque l’utilisateur compose un message ou une demande de réunion, assurez-vous que l’élément n’est pas protégé par IRM.


## Est-ce que le manifeste du complément est correctement installé et est-ce qu’Outlook dispose d’une copie mise en cache ?


Ce scénario s’applique uniquement à Outlook pour Windows. Normalement, quand vous installez un complément Outlook pour une boîte aux lettres, le serveur Exchange copie le manifeste du complément de l’emplacement que vous indiquez vers la boîte aux lettres située sur ce serveur Exchange. Chaque fois qu’Outlook démarre, il lit l’ensemble des manifestes installés pour cette boîte aux lettres dans un cache temporaire situé à l’emplacement suivant : 

%LocalAppData%\Microsoft\Office\15.0\WEF 

Par exemple, pour l’utilisateur Jean, le cache peut se situer dans C:\Users\jean\AppData\Local\Microsoft\Office\15.0\WEF.

Si un complément ne s’active pour aucun élément, cela peut signifier que le manifeste n’a pas été correctement installé sur le serveur Exchange ou qu’Outlook n’a pas lu correctement le manifeste au démarrage. À l’aide du Centre d’administration Exchange, assurez-vous que le complément est installé et activé pour votre boîte aux lettres, puis redémarrez le serveur Exchange, si nécessaire.

La figure 1 montre un résumé des étapes à suivre pour vérifier si Outlook dispose d’une version valide du manifeste. 


**Figure 1. Organigramme des étapes à suivre pour vérifier si Outlook a correctement mis en cache le manifeste**

![Organigramme de vérification du manifeste](../../images/off15appsdk_TroubleshootManifest.png)La procédure suivante décrit les détails.



1. Si vous modifiez le manifeste quand Outlook est ouvert et si vous n’utilisez pas les Outils de développement Office 365 « Napa », Visual Studio 2012 ou une version ultérieure de Visual Studio pour développer le complément, désinstallez-le, puis réinstallez-le via le Centre d’administration Exchange. 
    
2. Redémarrez Outlook, puis vérifiez si Outlook active désormais le complément.
    
3. Si Outlook n’active pas le complément, vérifiez si Outlook dispose d’une copie correctement mise en cache du manifeste du complément. Regardez dans le chemin d’accès suivant :
    
    %LocalAppData%\Microsoft\Office\15.0\WEF
    
    Vous trouverez le manifeste dans le sous-dossier suivant :
```
    \<insert your guid>\<insert base 64 hash>\Manifests\<ManifestID>_<ManifestVersion>
```
    
     >**Note**  The following is an example of a path to a manifest installed for a mailbox for the user John:
    
    C:\Users\john\appdata\Local\Microsoft\Office\15.0\WEF\{8D8445A4-80E4-4D6B-B7AC-D4E6AF594E73}\GoRshCWa7vW8+jhKmyiDhA==\Manifests\b3d7d9d5-6f57-437d-9830-94e2aaccef16_1.2
    
    Verify whether the manifest of the add-in you're testing is among the cached manifests.
    
4. Si le manifeste est dans le cache, ignorez le reste de cette section, puis examinez les autres raisons possibles à la suite de cette section.
    
5. Si le manifeste n’est pas dans le cache, vérifiez si Outlook a réussi à lire le manifeste à partir du serveur Exchange Server. Pour ce faire, utilisez l’Observateur d’événements Windows :
    
      1. Sous  **Journaux Windows**, choisissez  **Application**.
    
      2. Recherchez un événement relativement récent pour lequel l’ID d’événement est égal à 63, ce qui correspond au téléchargement par Outlook d’un manifeste auprès d’un serveur Exchange Server.
    
      3. Si Outlook a réussi à lire un manifeste, l’événement journalisé doit présenter la description suivante :
    
         **La demande de service web Exchange GetAppManifests a réussi.**
    
        Ignorez ensuite le reste de cette section, puis examinez les autres raisons possibles à la suite de cette section.
    

    Pour plus d’informations sur l’ouverture de l’Observateur d’événements dans Windows 7, voir [Ouvrir l’Observateur d’événements](http://windows.microsoft.com/en-US/windows7/Open-Event-Viewer).
    
6. Si vous ne voyez pas d’événement réussi, fermez Outlook et supprimez tous les manifestes du chemin d’accès suivant :
```
    %LocalAppData%\Microsoft\Office\15.0\WEF\<insert your guid>\<insert base 64 hash>\Manifests\
```
    Start Outlook and test whether Outlook now activates the add-in.
    
7. Si Outlook n’active pas le complément, revenez à l’étape 3 pour revérifier si Outlook a correctement lu le manifeste.
    

## Utilisez-vous les règles d’activation appropriées ?


À partir de la version 1.1 du schéma des manifestes des Compléments Office, vous pouvez créer des compléments qui sont activés lorsque l’utilisateur se trouve dans un formulaire de composition (compléments de composition) ou de lecture (compléments de lecture). Assurez-vous que vous spécifiez les règles d’activation appropriées pour chaque type de formulaire dans lequel votre complément est censé être activé. Par exemple, vous ne pouvez activer des compléments de composition qu’à l’aide des règles [ItemIs](http://msdn.microsoft.com/en-us/library/f7dac4a3-1574-9671-1eda-47f092390669%28Office.15%29.aspx) avec l’attribut **FormType** défini sur **Edit** ou **ReadOrEdit** et vous ne pouvez utiliser aucun autre type de règle, comme les règles [ItemHasKnownEntity](http://msdn.microsoft.com/en-us/library/87e10fd2-eab4-c8aa-bec3-dff92d004d39%28Office.15%29.aspx) et [ItemHasRegularExpressionMatch](http://msdn.microsoft.com/en-us/library/bfb726cd-81b0-a8d5-644f-2ca90a5273fc%28Office.15%29.aspx) pour les compléments de composition. Pour plus d’informations, voir [Règles d’activation pour les compléments Outlook](../outlook/manifests/activation-rules.md).


## Si vous utilisez une expression régulière, est-elle correctement spécifiée ?


Les expressions régulières contenues dans les règles d’activation font partie du fichier manifeste XML d’un complément de lecture. Si une expression régulière utilise certains caractères, veillez à bien suivre la séquence d’échappement correspondante prise en charge par les processeurs XML. Le tableau 1 répertorie ces caractères spéciaux. 


**Tableau 1. Séquences d’échappement des expressions régulières**


|**Caractère**|**Description**|**Séquence d’échappement à utiliser**|
|:-----|:-----|:-----|
|"|Guillemets doubles|&amp;quot;|
|&amp;|Esperluette|&amp;amp;|
|'|Apostrophe|&amp;apos;|
|<|Signe inférieur à|&amp;lt;|
|>|Signe supérieur à|&amp;gt;|

## Si vous utilisez une expression régulière, est-ce que le complément de lecture s’active dans Outlook Web App ou OWA pour périphériques, mais pas dans l’un des clients riches Outlook ?


Les clients riches Outlook ont recours à un autre moteur d’expressions régulières que Outlook Web App et OWA pour périphériques. Ils utilisent le moteur d’expressions régulières C++ fourni dans le cadre de la bibliothèque de modèles standard Visual Studio. Ce moteur est conforme aux normes ECMAScript 5. Outlook Web App et OWA pour périphériques utilisent l’évaluation d’expression régulière incluse dans JavaScript. Celle-ci est fournie par le navigateur et prend en charge un sur-ensemble d’ECMAScript 5. 

Dans la plupart des cas, ces applications hôtes trouvent les mêmes correspondances pour une expression régulière identique dans une règle d’activation. Toutefois, il existe des exceptions : en effet, si l’expression régulière comprend une classe de caractères personnalisée basée sur des classes de caractères prédéfinies, un client riche Outlook peut renvoyer des résultats différents de Outlook Web App et OWA pour périphériques. Par exemple, les classes de caractères qui contiennent des classes de caractères abrégées  `[\d\w]` renvoient des résultats distincts. Dans ce cas, pour éviter d’obtenir des résultats différents en fonction de l’hôte, utilisez `(\d|\w)` à la place.

Testez minutieusement votre expression régulière. Si elle renvoie des résultats distincts, réécrivez-la pour assurer sa compatibilité avec les deux moteurs. Pour vérifier les résultats de l’évaluation sur un client riche Outlook, écrivez un petit programme en C++ qui applique l’expression régulière à un échantillon du texte que vous essayez de faire correspondre. S’exécutant dans Visual Studio, le programme de test en C++ utilise la bibliothèque de modèles standard, ce qui permet de simuler le comportement du client riche Outlook lors de l’exécution de la même expression régulière. Pour vérifier les résultats de l’évaluation dans Outlook Web App ou OWA pour périphériques, utilisez le programme de test d’expression régulière en JavaScript de votre choix.


## Si vous utilisez une règle ItemIs, ItemHasAttachment ou ItemHasRegularExpressionMatch, avez-vous vérifié la propriété de l’élément connexe ?


Si vous utilisez une règle d’activation  **ItemHasRegularExpressionMatch**, vérifiez si la valeur de l’attribut  **PropertyName** correspond à ce que vous attendez pour l’élément sélectionné. Voici quelques conseils pour déboguer les propriétés correspondantes :


- Si l’élément sélectionné est un message et que vous spécifiez  **BodyAsHTML** dans l’attribut **PropertyName**, ouvrez le message, puis choisissez  **Afficher la source** afin de vérifier le corps du message dans la représentation HTML de cet élément.
    
- Si l’élément sélectionné est un rendez-vous ou si la règle d’activation spécifie  **BodyAsPlaintext** dans **PropertyName**, vous pouvez utiliser le modèle objet Outlook et Visual Basic Editor dans Outlook pour Windows :
    
      1. Assurez-vous que les macros sont activées et que l’onglet **Développeur** est affiché dans le ruban d’Outlook. Si vous n’êtes pas sûr de savoir comment procéder, consultez les étapes 1 et 2 sous [Est-ce que la boîte aux lettres utilisateur se trouve sur une version d’Exchange Server correspondant au minimum à Exchange 2013 ?](#est-ce-que-la-boîte-aux-lettres-utilisateur-se-trouve-sur-une-version-dexchange-server-correspondant-au-minimum-à-exchange2013)
    
      2. Dans Visual Basic Editor, choisissez **Affichage**, **Fenêtre exécution**.
    
      3. Tapez ce qui suit pour afficher diverses propriétés en fonction du scénario.
    
      - Corps HTML de l’élément de message ou de rendez-vous sélectionné dans l’explorateur Outlook :
    
            
              ?ActiveExplorer.Selection.Item(1).HTMLBody
        


     - Corps en texte brut de l’élément de message ou de rendez-vous sélectionné dans l’explorateur Outlook :
    
            
              ?ActiveExplorer.Selection.Item(1).Body
            


      - Corps HTML de l’élément de message ou de rendez-vous ouvert dans l’inspecteur Outlook actif :
    
            
              ?ActiveInspector.CurrentItem.HTMLBody
        
      - Corps en texte brut de l’élément de message ou de rendez-vous ouvert dans l’inspecteur Outlook actif :
    
            
              ?ActiveInspector.CurrentItem.Body
            

Si la règle d’activation  **ItemHasRegularExpressionMatch** spécifie **Subject** ou **SenderSMTPAddress**, ou si vous utilisez une règle  **ItemIs** ou **ItemHasAttachment** et si vous êtes habitué à l’interface MAPI (ou si vous souhaitez l’utiliser), vous pouvez employer [MFCMAPI](http://mfcmapi.codeplex.com/) pour vérifier la valeur du tableau 2 dont dépend votre règle.


**Tableau 2. Règles d’activation et propriétés MAPI correspondantes**


|**Type de règle**|**Vérifier cette propriété MAPI**|
|:-----|:-----|
|Règle  **ItemHasRegularExpressionMatch** avec **Subject**|[PidTagSubject](http://msdn.microsoft.com/en-us/library/aa7ba4d9-c5e0-4ce7-a34e-65f675223bc9%28Office.15%29.aspx)|
|Règle  **ItemHasRegularExpressionMatch** avec **SenderSMTPAddress**|  [PidTagSenderSmtpAddress](http://msdn.microsoft.com/en-us/library/321cde5a-05db-498b-a9b8-cb54c8a14e34%28Office.15%29.aspx) et [PidTagSentRepresentingSmtpAddress](http://msdn.microsoft.com/en-us/library/5ed122a2-0967-4de3-a2ee-69f81ae77b16%28Office.15%29.aspx)|
|**ItemIs**|[PidTagMessageClass](http://msdn.microsoft.com/en-us/library/1e704023-1992-4b43-857e-0a7da7bc8e87%28Office.15%29.aspx)|
|**ItemHasAttachment**|[PidTagHasAttachments](http://msdn.microsoft.com/en-us/library/fd236d74-2868-46a8-bb3d-17f8365931b6%28Office.15%29.aspx)|
Après avoir vérifié la valeur de propriété, vous pouvez utiliser un outil d’évaluation d’expression régulière pour vérifier si l’expression régulière trouve une correspondance dans cette valeur.


## Est-ce que l’application hôte applique toutes les expressions régulières à la partie du corps de l’élément comme prévu ?


Cette section s’applique à toutes les règles d’activation qui utilisent des expressions régulières ; en particulier, celles appliquées au corps d’élément, qui peut être volumineux et demander plus de temps pour l’évaluation des correspondances. Notez que même si la propriété d’élément dont dépend une règle d’activation a la valeur attendue, l’application hôte ne parvient pas toujours à évaluer toutes les expressions régulières pour la valeur complète de la propriété d’élément. Pour offrir des performances raisonnables et contrôler l’utilisation excessive des ressources par un complément de lecture, Outlook, Outlook Web App et OWA pour les appareils respectent les limites suivantes en matière de traitement des expressions régulières des règles d’activation au moment de l’exécution :


- Taille du corps d’élément évalué -- il existe des limites à la partie d’un corps d’élément pour lequel l’application hôte évalue une expression régulière. Ces limites dépendent de l’application hôte, du facteur de forme et du format du corps d’élément. Consultez les détails du tableau 2 dans [Limites d’activation et d’API JavaScript des compléments Outlook](../outlook/limits-for-activation-and-javascript-api-for-outlook-add-ins.md).
    
- Nombre de correspondances des expressions régulières -- le client riche Outlook, Outlook Web App et OWA pour les appareils renvoient chacun 50 correspondances d’expressions régulières au maximum. Ces correspondances sont uniques. Par ailleurs, les correspondances en double ne sont pas prises en compte dans cette limite. Ne présumez pas de l’ordre des correspondances renvoyées. En outre, ne vous attendez pas à ce que l’ordre établi dans un client riche Outlook soit le même que dans Outlook Web App et OWA pour les appareils. Si vous prévoyez de nombreuses correspondances pour les expressions régulières de vos règles d’activation et qu’il vous en manque une, cela signifie peut-être que vous dépassez cette limite.
    
- Longueur d’une correspondance d’expression régulière -- il existe des limites à la longueur d’une correspondance d’expression régulière retournée par l’application hôte. L’application hôte n’inclut aucune correspondance au-delà de la limite et n’affiche aucun message d’avertissement. Vous pouvez exécuter votre expression régulière à l’aide d’autres outils d’évaluation d’expression régulière ou via un programme de test autonome en C++ afin de vérifier s’il existe une correspondance qui dépasse les limites définies. Le tableau 3 récapitule ces limites. Pour plus d’informations, voir le tableau 3 dans [Limites d’activation et d’API JavaScript des compléments Outlook](../outlook/limits-for-activation-and-javascript-api-for-outlook-add-ins.md).
    
    **Tableau 3. Limites de longueur pour une correspondance d’expression régulière**


|**Limite de longueur d’une correspondance d’expression régulière**|**Clients riches Outlook**|**Outlook Web App ou OWA pour périphériques**|
|:-----|:-----|:-----|
|Corps d’élément en texte brut|1,5 Ko|3 Ko|
|Corps d’élément en HTML|3 Ko|3 Ko|
- Temps consacré à l’évaluation de toutes les expressions régulières d’un complément de lecture (pour un client riche Outlook) : par défaut, pour chaque complément de lecture, Outlook doit terminer l’évaluation de toutes les expressions régulières contenues dans ses règles d’activation en moins d’une seconde. Sinon, Outlook effectue jusqu’à trois nouvelles tentatives avant de désactiver le complément si l’évaluation ne peut pas être achevée. Outlook affiche un message dans la barre de notification pour indiquer que le complément a été désactivé. Vous pouvez modifier le délai disponible pour votre expression régulière en définissant une stratégie de groupe ou une clé de Registre. 
    
     >**Remarque**  Si le client riche Outlook désactive un complément de lecture, le complément de lecture ne peut pas être utilisé pour la même boîte aux lettres sur le client riche Outlook, l’application web Outlook et OWA pour les appareils.

## Ressources supplémentaires



- [Déployer et installer des compléments Outlook à des fins de test](../outlook/testing-and-tips.md)
    
- [Règles d’activation pour les compléments Outlook](../outlook/manifests/activation-rules.md)
    
- [Utiliser des règles d’activation d’expression régulière pour afficher un complément Outlook](../outlook/use-regular-expressions-to-show-an-outlook-add-in.md)
    
- [Limites pour l’activation et l’API JavaScript pour les compléments Outlook](../outlook/limits-for-activation-and-javascript-api-for-outlook-add-ins.md)
    
- [Ouvrir l’Observateur d’événements](http://windows.microsoft.com/en-US/windows7/Open-Event-Viewer)
    
- [ItemHasAttachment complexType](http://msdn.microsoft.com/en-us/library/031db7be-8a25-5185-a9c3-93987e10c6c2%28Office.15%29.aspx)
    
- [ItemHasRegularExpressionMatch complexType](http://msdn.microsoft.com/en-us/library/bfb726cd-81b0-a8d5-644f-2ca90a5273fc%28Office.15%29.aspx)
    
- [ItemIs complexType](http://msdn.microsoft.com/en-us/library/926249ab-2d2f-39f5-1d73-fab1c989966f%28Office.15%29.aspx)
    
- [MailApp complexType](http://msdn.microsoft.com/en-us/library/696b9fcf-cd10-3f20-4d49-86d3690c887a%28Office.15%29.aspx)
    
