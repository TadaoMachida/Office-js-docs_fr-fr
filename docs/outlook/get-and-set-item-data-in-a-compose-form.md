
# Obtenir et définir des données d’élément dans un formulaire de composition dans Outlook
Découvrez comment obtenir ou définir diverses propriétés d’un élément dans un complément Outlook d’un scénario de composition, y compris ses destinataires, son objet, son corps, et ses emplacement et heure de rendez-vous.




## Obtention et définition des propriétés d’un élément pour un complément de composition


Dans un formulaire de composition, vous pouvez obtenir la plupart des propriétés qui sont exposées sur le même genre d’élément que dans un formulaire de lecture (comme attendees, recipients, subject et body), et vous pouvez obtenir quelques propriétés supplémentaires qui sont pertinentes uniquement dans un formulaire de composition mais pas dans un formulaire de lecture (body, bcc). 

Pour la plupart de ces propriétés, comme il est possible qu’un complément Outlook et l’utilisateur modifient la même propriété dans l’interface utilisateur en même temps, les méthodes d’obtention et de définition de ces propriétés sont asynchrones. Le tableau 1 énumère les propriétés de niveau élément et les méthodes asynchrones correspondantes pour les obtenir et les définir dans un formulaire de composition. Les propriétés  [item.itemType](../../reference/outlook/Office.context.mailbox.item.md) et [item.conversationId](../../reference/outlook/Office.context.mailbox.item.md) constituent des exceptions, car les utilisateurs ne peuvent pas les modifier. Vous pouvez les obtenir par programmation de la même façon dans un formulaire de composition et dans un formulaire de lecture, directement à partir de l’objet parent.

En plus d’accéder aux propriétés de niveau élément dans l’interface API JavaScript pour Office, vous pouvez également y accéder à l’aide des services web Exchange (EWS). Avec l’autorisation  **ReadWriteMailbox**, vous pouvez utiliser la méthode [mailbox.makeEwsRequestAsync](../../reference/outlook/Office.context.mailbox.md) pour accéder aux opérations EWS, [GetItem](http://msdn.microsoft.com/en-us/library/e3590b8b-c2a7-4dad-a014-6360197b68e4%28Office.15%29.aspx) and [UpdateItem](http://msdn.microsoft.com/en-us/library/5d027523-e0bc-4da2-b60b-0cb9fc1fdfe4%28Office.15%29.aspx), pour obtenir et définir plus de propriétés d’au moins un élément dans la boîte aux lettres de l’utilisateur.  **makeEwsRequestAsync** est disponible à la fois dans les formulaires de lecture et de composition. Pour plus d’informations sur l’autorisation **ReadWriteMailbox** et l’accès à EWS par le biais de la plateforme des Compléments Office, voir [Spécifier les autorisations pour l’accès du complément Outlook à la boîte aux lettres de l’utilisateur](../outlook/understanding-outlook-add-in-permissions.md) et [Appeler des services web à partir d’un complément Outlook](../outlook/web-services.md).


**Tableau 1. Méthodes asynchrones pour obtenir ou définir des propriétés d’élément dans un formulaire de composition**


|**Propriété**|**Type de propriété**|**Méthode asynchrone d’obtention**|**Méthode(s) asynchrone(s) de définition**|
|:-----|:-----|:-----|:-----|
|[bcc](../../reference/outlook/Office.context.mailbox.item.md)|[Destinataires](../../reference/outlook/Recipients.md)|[Recipients.getAsync](../../reference/outlook/Recipients.md)|[Recipients.addAsync](../../reference/outlook/Recipients.md)[Recipients.setAsync](../../reference/outlook/Recipients.md)|
|[corps](../../reference/outlook/Office.context.mailbox.item.md)|[Body](../../reference/outlook/Body.md)|[Body.getAsync](../../reference/outlook/Body.md)|[Body.prependAsync](../../reference/outlook/Body.md)[Body.setAsync](../../reference/outlook/Body.md)[Body.setSelectedDataAsync](../../reference/outlook/Body.md)|
|[cc](../../reference/outlook/Office.context.mailbox.item.md)|Destinataires|Recipients.getAsync|Recipients.addAsync Recipients.setAsync|
|[end](../../reference/outlook/Office.context.mailbox.item.md)|[Heure](../../reference/outlook/Time.md)|[Time.getAsync](../../reference/outlook/Time.md)|[Time.setAsync](../../reference/outlook/Time.md)|
|[location](../../reference/outlook/Office.context.mailbox.item.md)|[Emplacement](../../reference/outlook/Location.md)|[Location.getAsync](../../reference/outlook/Location.md)|[Location.setAsync](../../reference/outlook/Location.md)|
|[optionalAttendees](../../reference/outlook/Office.context.mailbox.item.md)|Destinataires|Recipients.getAsync|Recipients.addAsync Recipients.setAsync|
|[requiredAttendees](../../reference/outlook/Office.context.mailbox.item.md)|Destinataires|Recipients.getAsync|Recipients.addAsync Recipients.setAsync|
|[démarrer](../../reference/outlook/Office.context.mailbox.item.md)|Heure|Time.getAsync|Time.setAsync|
|[subject](../../reference/outlook/Office.context.mailbox.item.md)|[Objet](../../reference/outlook/Subject.md)|[Subject.getAsync](../../reference/outlook/Subject.md)|[Subject.setAsync](../../reference/outlook/Subject.md)|
|[à](../../reference/outlook/Office.context.mailbox.item.md)|Destinataires|Recipients.getAsync|Recipients.addAsync Recipients.setAsync|



## Ressources supplémentaires



- [Créer des compléments Outlook pour les formulaires de composition](../outlook/compose-scenario.md)
    
- [Présentation des autorisations de complément Outlook](../outlook/understanding-outlook-add-in-permissions.md)
    
- [Appeler des services web à partir d’un complément Outlook](../outlook/web-services.md)
    
- [Obtention et définition de données d’élément Outlook dans des formulaires de lecture ou de composition](../outlook/item-data.md)
    


