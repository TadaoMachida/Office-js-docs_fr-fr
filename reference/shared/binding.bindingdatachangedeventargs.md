
# BindingDataChangedEventArgs, objet
Fournit des informations sur la liaison qui a déclenché l’événement [DataChanged](../../reference/shared/binding.bindingdatachangedevent.md).

|||
|:-----|:-----|
|**Hôtes :**|Access, Excel, Word|
|**Dernière modification dans BindingEvents**|1.1|

```js
Office.EventType.BindingDataChanged
```


## Membres


**Propriétés**


|**Nom**|**Description**|
|:-----|:-----|
|[liaison](../../reference/shared/binding.bindingdatachangedeventargs.binding.md)|Obtient un objet [Binding](../../reference/shared/binding.md) qui représente la liaison ayant déclenché l’événement **DataChanged**.|
|[type](../../reference/shared/binding.bindingdatachangedeventargs.type.md)|Obtient une valeur d’énumération [EventType](../../reference/shared/eventtype-enumeration.md) qui identifie le genre d’événement déclenché.|

## Informations de prise en charge


Un Y majuscule dans la matrice suivante indique que cet objet est pris en charge dans l’application hôte Office correspondante. Une cellule vide indique que l’application hôte Office ne prend pas en charge cet objet.

Pour plus d’informations sur les exigences de l’application et du serveur hôtes Office, voir [Configuration requise pour exécuter des compléments pour Office](../../docs/overview/requirements-for-running-office-add-ins.md).


**Hôtes pris en charge par la plateforme**


||**Office pour Bureau Windows**|**Office Online (dans un navigateur)**|**Office pour iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||v||
|**Excel**|v|v|v|
|**Word**|v||v|

|||
|:-----|:-----|
|**Types de complément**|Application de contenu et de volet de tâches|
|**Bibliothèque**|Office.js|
|**Espace de noms**|Bureau|

## Historique de prise en charge




|**Version**|**Modifications**|
|:-----|:-----|
|1.1|Prise en charge supplémentaire d’Excel et de Word dans Office pour iPad.|
|1.1|Prise en charge supplémentaire de cet événement dans les compléments pour Access.|
|1.0|Introduit|
