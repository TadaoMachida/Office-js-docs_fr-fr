
# Propriété BindingSelectionChangedEventArgs.type
Obtient une valeur d’énumération **EventType** qui identifie le type d’événement déclenché.

|||
|:-----|:-----|
|**Hôtes :**|Access, Excel, Word|
|**Dernière modification dans **|1.1|

```js
var myEvent = eventArgsObj.type;
```


## Valeur renvoyée

[EventType](../../reference/shared/eventtype-enumeration.md) de l’événement qui a été déclenché.


## Informations de prise en charge


Un Y majuscule dans la matrice suivante indique que cette propriété est prise en charge dans l'application hôte Office correspondante. Une cellule vide indique que l'application hôte Office ne prend pas en charge cette propriété.

Pour plus d’informations sur les exigences de l’application et du serveur hôtes Office, voir [Configuration requise pour exécuter des compléments pour Office](../../docs/overview/requirements-for-running-office-add-ins.md).


**Hôtes pris en charge par la plateforme**


||**Office pour Bureau Windows**|**Office Online (dans un navigateur)**|**Office pour iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||v||
|**Excel**|v|v|v|
|**Word**|v||v|

|||
|:-----|:-----|
|**Disponible dans les ensembles de ressources requis**|Selection|
|**Niveau d’autorisation minimal**|[WriteDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Types de complément**|Application de contenu et de volet de tâches|
|**Bibliothèque**|Office.js|
|**Espace de noms**|Bureau|

## Historique de prise en charge


|**Version**|**Modifications**|
|:-----|:-----|
|1.1|Prise en charge supplémentaire d’Excel et de Word dans Office pour iPad.|
|1.1|Prise en charge supplémentaire des compléments pour Access.|
|1.0|Introduit|
