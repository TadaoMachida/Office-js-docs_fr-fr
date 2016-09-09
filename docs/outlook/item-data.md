
# Obtention et définition de données d’élément Outlook dans des formulaires de lecture ou de composition

À partir de la version 1.1 du schéma des manifestes des Compléments Office, Outlook peut activer des compléments lorsque l’utilisateur visualise ou compose un élément. Selon qu’un complément est activé dans un formulaire de lecture ou de composition, les propriétés disponibles pour le complément sur l’élément diffèrent également. Par exemple, les propriétés [dateTimeCreated](../../reference/outlook/Office.context.mailbox.item.md) et [dateTimeModified](../../reference/outlook/Office.context.mailbox.item.md) sont définies uniquement pour un élément qui a déjà été envoyé (l’élément est ensuite affiché dans un formulaire de lecture), mais pas pour un élément en cours de création (dans un formulaire de composition). Un autre exemple est la propriété [bcc](../../reference/outlook/Office.context.mailbox.item.md) qui n’est pertinente que si un message est en cours de création (dans un formulaire de composition) et qui n’est pas accessible à l’utilisateur dans un formulaire de lecture.

Le tableau 1 montre les propriétés de niveau élément de l’interface API JavaScript pour Office qui sont disponibles dans chacun des modes lecture et composition des compléments de messagerie. En règle générale, les propriétés disponibles dans des formulaires de lecture sont en lecture seule, et celles disponibles dans des formulaires de composition sont en lecture/écriture, à l’exception des propriétés [itemId](../../reference/outlook/Office.context.mailbox.item.md) et [conversationId](../../reference/outlook/Office.context.mailbox.item.md), qui sont toujours en lecture seule. Pour les propriétés de niveau élément restantes et disponibles dans des formulaires de composition, comme le complément et l’utilisateur peuvent être en train de lire ou d’écrire la même propriété en même temps, les méthodes pour les obtenir ou les définir en mode composition sont asynchrones. Par conséquent, les types d’objets renvoyés par ces propriétés sont également différents dans les formulaires de composition et dans les formulaires de lecture. Pour plus d’informations sur l’utilisation de méthodes asynchrones pour obtenir ou définir des propriétés de niveau élément en mode composition, voir [Obtenir et définir des données d’élément dans un formulaire de composition dans Outlook](../outlook/get-and-set-item-data-in-a-compose-form.md)


**Tableau 1. Propriétés d’éléments disponibles dans les formulaires de composition et de lecture**


|**Type d’élément**|**Propriété**|**Type de propriété dans les formulaires de lecture**|**Type de propriété dans les formulaires de composition**|
|:-----|:-----|:-----|:-----|
|Rendez-vous et messages|[dateTimeCreated](../../reference/outlook/Office.context.mailbox.item.md)|Objet  **Date** JavaScript|Propriété non disponible|
|Rendez-vous et messages|[dateTimeModified](../../reference/outlook/Office.context.mailbox.item.md)|Objet  **Date** JavaScript|Propriété non disponible|
|Rendez-vous et messages|[itemClass](../../reference/outlook/Office.context.mailbox.item.md)|Chaîne|Propriété non disponible|
|Rendez-vous et messages|[itemId](../../reference/outlook/Office.context.mailbox.item.md)|String|Propriété non disponible|
|Rendez-vous et messages|[itemType](../../reference/outlook/Office.context.mailbox.item.md)|String dans l’énumération [ItemType](../../reference/outlook/Office.MailboxEnums.md)|Propriété non disponible|
|Rendez-vous et messages|[pièces jointes](../../reference/outlook/Office.context.mailbox.item.md)|[AttachmentDetails](../../reference/outlook/simple-types.md)|Propriété non disponible|
|Rendez-vous et messages|[corps](../../reference/outlook/Office.context.mailbox.item.md)|[Body](../../reference/outlook/Body.md)|[Body](../../reference/outlook/Body.md)|
|Rendez-vous|[end](../../reference/outlook/Office.context.mailbox.item.md)|Objet  **Date** JavaScript|[Heure](../../reference/outlook/Time.md)|
|Rendez-vous|[location](../../reference/outlook/Office.context.mailbox.item.md)|String|[Emplacement](../../reference/outlook/Location.md)|
|Rendez-vous et messages|[normalizedSubject](../../reference/outlook/Office.context.mailbox.item.md)|String|Propriété non disponible|
|Rendez-vous|[optionalAttendees](../../reference/outlook/Office.context.mailbox.item.md)|[EmailAddressDetails](../../reference/outlook/simple-types.md)|[Destinataires](../../reference/outlook/Recipients.md)|
|Rendez-vous|[organizer](../../reference/outlook/Office.context.mailbox.item.md)|EmailAddressDetails|Propriété non disponible|
|Rendez-vous|[requiredAttendees](../../reference/outlook/Office.context.mailbox.item.md)|EmailAddressDetails|Destinataires|
|Rendez-vous|[ressources](../../reference/outlook/Office.context.mailbox.item.md)|Chaîne|Propriété non disponible|
|Rendez-vous|[démarrer](../../reference/outlook/Office.context.mailbox.item.md)|Objet  **Date** JavaScript|Heure|
|Rendez-vous et messages|[subject](../../reference/outlook/Office.context.mailbox.item.md)|String|[Objet](../../reference/outlook/Subject.md)|
|Messages|[bcc](../../reference/outlook/Office.context.mailbox.item.md)|Propriété non disponible|Destinataires|
|Messages|[cc](../../reference/outlook/Office.context.mailbox.item.md)|EmailAddressDetails|Destinataires|
|Messages|[conversationId](../../reference/outlook/Office.context.mailbox.item.md)|Chaîne|String (lecture seule)|
|Messages|[depuis](../../reference/outlook/Office.context.mailbox.item.md)|EmailAddressDetails|Propriété non disponible|
|Messages|[internetMessageId](../../reference/outlook/Office.context.mailbox.item.md)|Entier|Propriété non disponible|
|Messages|[sender](../../reference/outlook/Office.context.mailbox.item.md)|EmailAddressDetails|Propriété non disponible|
|Messages|[à](../../reference/outlook/Office.context.mailbox.item.md)|EmailAddressDetails|Destinataires|

## Utilisation de jetons de rappel Exchange Server à partir d’un complément de lecture


Si votre complément Outlook doit être activé dans des formulaires de lecture, vous pouvez obtenir un jeton de rappel Exchange. Celui-ci peut être utilisé dans le code côté serveur afin d’accéder à l’élément complément par le biais des services web Exchange (EWS). En spécifiant l’autorisation  **ReadItem** dans le manifeste du complément, vous pouvez utiliser la méthode [mailbox.getCallbackTokenAsync](../../reference/outlook/Office.context.mailbox.md) pour obtenir un jeton de rappel Exchange, la propriété [mailbox.ewsUrl](../../reference/outlook/Office.context.mailbox.md) pour obtenir l’URL du point de terminaison EWS pour la boîte aux lettres de l’utilisateur et [item.itemId](../../reference/outlook/Office.context.mailbox.item.md) pour obtenir l’ID EWS de l’élément sélectionné. Vous pouvez ensuite transmettre le jeton de rappel, l’URL de point de terminaison EWS et l’ID d’élément EWS au code côté serveur pour accéder à l’opération [GetItem](http://msdn.microsoft.com/en-us/library/e3590b8b-c2a7-4dad-a014-6360197b68e4%28Office.15%29.aspx) et obtenir plus de propriétés de l’élément.


## Accès à EWS à partir d’un complément de composition ou de lecture


Vous pouvez également utiliser la méthode [mailbox.makeEwsRequestAsync](../../reference/outlook/Office.context.mailbox.md) pour accéder aux opérations des services web Exchange (EWS)[GetItem](http://msdn.microsoft.com/en-us/library/e3590b8b-c2a7-4dad-a014-6360197b68e4%28Office.15%29.aspx) and [UpdateItem](http://msdn.microsoft.com/en-us/library/5d027523-e0bc-4da2-b60b-0cb9fc1fdfe4%28Office.15%29.aspx) directement à partir du complément. Vous pouvez utiliser ces opérations pour obtenir et définir de nombreuses propriétés d’un élément spécifié. Cette méthode est disponible pour les compléments Outlook indépendamment du fait que le complément ait été activé dans un formulaire de lecture ou de composition, tant que vous spécifiez l’autorisation **ReadWriteMailbox** dans le manifeste du complément. Pour plus d’informations sur l’utilisation de **makeEwsRequestAsync** pour accéder aux opérations EWS, voir [Appeler des services web à partir d’un complément Outlook](../outlook/web-services.md).


## Ressources supplémentaires



- [Compléments Outlook](../outlook/outlook-add-ins.md)
    
- [Obtenir et définir des données d’élément dans un formulaire de composition dans Outlook](../outlook/get-and-set-item-data-in-a-compose-form.md)
    
- [Appeler des services web à partir d’un complément Outlook](../outlook/web-services.md)
    


