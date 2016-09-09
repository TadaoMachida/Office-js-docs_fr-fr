
# Énumérations

Vous pouvez spécifier une valeur énumérée à l’aide de son nom d’énumération complet (`Office.CoercionType.Text`) ou de sa valeur de texte correspondante (`"text"`). Par exemple, l’appel de méthode suivant utilise des noms d’énumération :


```js
Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, {valueFormat:Office.ValueFormat.Unformatted, filterType:Office.FilterType.All},
   function (result) {
      if (result.status === Office.AsyncResultStatus.Success)
         var dataValue = result.value; // Get selected data.
         write('Selected data is ' + dataValue);
      else {
         var err = result.error;
         write(err.name + ": " + err.message);
      }
   });

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message;
}
```


Voici le même appel qui utilise les valeurs texte d’énumération :




```js
Office.context.document.getSelectedDataAsync("text", {valueFormat:"unformatted", filterType:"all"},
   function (result) {
      if (result.status === "success")
         var dataValue = result.value; // Get selected data.
         write('Selected data is ' + dataValue);
      else {
         var err = result.error;
         write(err.name + ": " + err.message);
      }
   });
```


## Référence



|**Nom**|**Définition**|
|:-----|:-----|
|[ActiveView](activeview-enumeration.md)|Spécifie l’état de l’affichage dynamique du document, par exemple, si l’utilisateur peut modifier le document.|
|[AsyncResultStatus](asyncresultstatus-enumeration.md)|Spécifie le résultat d’un appel asynchrone.|
|[AttachmentType](http://msdn.microsoft.com/library/83883a47-a937-4afb-a55e-e789057335c4%28Office.15%29.aspx)|Spécifie le type d’une pièce jointe d’un message électronique ou d’une demande de rendez-vous. Outlook 2013 ne prend pas en charge cette énumération.|
|[BindingType](bindingtype-enumeration.md)|Spécifie le type de l’objet de liaison qui doit être retourné.|
|[BodyType](http://msdn.microsoft.com/library/31350fe6-4c42-4cbb-a5b2-4fb2d360fa11%28Office.15%29.aspx)|Spécifie le type de texte pour le corps d’un message ou un rendez-vous.|
|[CoercionType](coerciontype-enumeration.md)|Indique comment forcer le type des données retournées ou définies par la méthode appelée.|
|[CustomXMLNodeType](customxmlnodetype-enumeration.md)|Spécifie le type de nœud.|
|[DocumentMode](documentmode-enumeration.md)|Spécifie si le document de l’application associée est en lecture seule ou en lecture/écriture. |
|[EntityType](http://msdn.microsoft.com/library/0035be38-8a65-4693-bcc4-0a8dd7b1495b%28Office.15%29.aspx)|Spécifie le type d’une entité.|
|[EventType](eventtype-enumeration.md)|Spécifie le genre de l’événement qui a été déclenché.|
|[FileType](filetype-enumeration.md)|Spécifie le format de retour du document.|
|[GoToType](gototype-enumeration.md)|Spécifie le type d’emplacement ou d’objet auquel accéder|
|[FilterType](filtertype-enumeration.md)|Spécifie si le filtrage à partir de l’application hôte est appliqué quand les données sont récupérées.|
|[InitializationReason](initializationreason-enumeration.md)|Indique si le complément vient d’être inséré ou s’il était déjà contenu dans le document.|
|[ItemType](http://msdn.microsoft.com/library/e0bb23fd-f360-4b0f-b72c-1cf08d4cab3f%28Office.15%29.aspx)|Spécifie le type d’un élément.|
|[notificationMessageType](http://msdn.microsoft.com/library/ff00c89d-0019-4545-a95b-7ed0db712ce9%28Office.15%29.aspx)|Spécifie le message de notification pour un rendez-vous ou un message.|
|[ProjectProjectFields](projectprojectfields-enumeration.md)|Spécifie les champs de projet disponibles en tant que paramètres pour la méthode [getProjectFieldAsync](projectdocument.getprojectfieldasync.md).|
|[ProjectResourceFields](projectresourcefields-enumeration.md)|Spécifie les champs de ressource disponibles en tant que paramètres pour la méthode [getResourceFieldAsync](projectdocument.gettaskfieldasync.md).|
|[ProjectTaskFields](projecttaskfields-enumeration.md)|Spécifie les champs de tâche disponibles en tant que paramètres pour la méthode [getTaskFieldAsync](projectdocument.gettaskfieldasync.md).|
|[ProjectViewTypes](projectviewtypes-enumeration.md)|Spécifie les types d’affichage que la méthode [getSelectedViewAsync](projectdocument.getselectedviewasync.md) peut reconnaître.|
|[RecipientType](http://msdn.microsoft.com/library/6e7c4029-6e52-47f6-98d2-4cd3ce7bd8b4%28Office.15%29.aspx)|Spécifie le type de destinataire d’un rendez-vous.|
|[ResponseType](http://msdn.microsoft.com/library/b3e723ca-4be0-4846-ad97-0eecab4355eb%28Office.15%29.aspx)|Spécifie la réponse à une invitation à une réunion.|
|[SelectionMode](selectionmode-enumeration.md)|Spécifie s’il faut sélectionner (mettre en surbrillance) l’emplacement à atteindre (lorsque la méthode [Document.goToByIdAsync](document.gotobyidasync.md) est utilisée).|
|[SourceProperty](http://msdn.microsoft.com/library/6a209a7f-57cd-4dc3-869e-07b0f5928b28%28Office.15%29.aspx)|Spécifie la source des données renvoyées par la méthode appelée.|
|[Table](table-enumeration.md)|Spécifie les valeurs énumérées de la propriété `cells:` dans le paramètre _cellFormat_ des [méthodes de mise en forme de tableau](../../docs/excel/format-tables-in-add-ins-for-excel.md).|
|[ValueFormat](valueformat-enumeration.md)|Spécifie si les valeurs (telles que les nombres et les dates) retournées par la méthode appelée sont retournées avec leur mise en forme appliquée.|

## Informations de prise en charge


La prise en charge de chaque énumération diffère dans les applications hôtes Office. Voir la section « Informations de prise en charge » de la rubrique de chaque énumération pour découvrir les informations de prise en charge d’hôte.

Pour plus d’informations sur les exigences de l’application et du serveur hôtes Office, voir [Configuration requise pour exécuter des compléments pour Office](../../docs/overview/requirements-for-running-office-add-ins.md).


|||
|:-----|:-----|
|**Types de complément**|De contenu, de volet de tâche, Outlook|
|**Bibliothèque**|Office.js|
|**Espace de noms**|Bureau|
