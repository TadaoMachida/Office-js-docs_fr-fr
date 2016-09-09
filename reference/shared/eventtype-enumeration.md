
# EventType, énumération
Spécifie le genre de l’événement qui a été déclenché. Renvoyé par la propriété **type** d’un objet _EventArgs_ **EventName**.

|||
|:-----|:-----|
|**Hôtes :**|Access, Excel, PowerPoint, Project, Word|
|**Dernière modification dans la sélection**|1.1|

```js
Office.EventType
```


## Membres


**Valeurs**


|Énumération|Valeur|Description|
|:-----|:-----|:-----|
|Office.EventType.ActiveViewChanged|"documentActiveViewChanged"|Un événement [Document.ActiveViewChanged](../../reference/shared/document.activeviewchanged.md) a été déclenché.|
|Office.EventType.DocumentSelectionChanged|"documentSelectionChanged"|Un événement [Document.SelectionChanged](../../reference/shared/document.selectionchanged.event.md) a été déclenché.|
|Office.EventType.BindingSelectionChanged|"bindingSelectionChanged"|Un événement [Binding.BindingSelectionChanged](../../reference/shared/binding.bindingselectionchangedevent.md) a été déclenché.|
|Office.EventType.BindingDataChanged|"bindingDataChanged"|Un événement [Binding.BindingDataChanged](../../reference/shared/binding.bindingdatachangedevent.md) a été déclenché.|
|Office.EventType.DataNodeDeleted|"nodeDeleted"|Un événement [CustomXmlPart.dataNodeDeleted](../../reference/shared/customxmlpart.datanodedeleted.event.md) a été déclenché.|
|Office.EventType.DataNodeInserted|"nodeInserted"|Un événement [CustomXmlPart.dataNodeInserted](../../reference/shared/customxmlpart.datanodeinserted.event.md) a été déclenché.|
|Office.EventType.DataNodeReplaced|"nodeReplaced"|Un événement [CustomXmlPart.dataNodeReplaced](../../reference/shared/customxmlpart.datanodereplaced.event.md) a été déclenché.|
|Office.EventType.SettingsChanged|"settingsChanged"|Un événement[Settings.settingsChanged](../../reference/shared/settings.settingschangedevent.md) a été déclenché.|

## Remarques


 >**Remarque** :  Les compléments pour Project prennent en charge les types d’événements **Office.EventType.ResourceSelectionChanged**, **Office.EventType.TaskSelectionChanged** et **Office.EventType.ViewSelectionChanged**.


## Informations de prise en charge


Un Y majuscule dans la matrice suivante indique que cette énumération est prise en charge dans l'application hôte Office correspondante. Une cellule vide indique que l'application hôte Office ne prend pas en charge cette énumération.

Pour plus d’informations sur les exigences de l’application et du serveur hôtes Office, voir [Configuration requise pour exécuter des compléments pour Office](../../docs/overview/requirements-for-running-office-add-ins.md).


**Hôtes pris en charge par la plateforme**


||**Office pour Bureau Windows**|**Office Online (dans un navigateur)**|**Office pour iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|v|v|v|
|**PowerPoint**|v|v||
|**Projet**|v|||
|**Word**|v||v|

|||
|:-----|:-----|
|**Types de complément**|Application de contenu et de volet de tâches|
|**Bibliothèque**|Office.js|
|**Espace de noms**|Bureau|

## Historique de prise en charge



|**Version**|**Modifications**|
|:-----|:-----|
|1.1| Ajout de l’énumération Office.EventType.ActiveViewChanged pour le nouvel événement **Document.ActiveViewChanged**.|
|1.0|Introduit|
