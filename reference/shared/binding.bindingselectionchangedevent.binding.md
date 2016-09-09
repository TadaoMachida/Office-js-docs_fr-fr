
# Propriété BindingSelectionChangedEventArgs.binding
Obtient un objet **Binding** qui représente la liaison ayant déclenché l’événement **SelectionChanged**.

|||
|:-----|:-----|
|**Hôtes :**|Access, Excel, Word|
|**Dernière modification dans **|1.1|

```
var myBinding = eventArgsObj.binding;
```


## Valeur renvoyée

Objet [Binding](../../reference/shared/binding.md) qui représente la liaison ayant déclenché l’événement [SelectionChanged](../../reference/shared/binding.bindingselectionchangedevent.md).


## Informations de prise en charge


Un Y majuscule dans la matrice suivante indique que cette propriété est prise en charge dans l'application hôte Office correspondante. Une cellule vide indique que l'application hôte Office ne prend pas en charge cette propriété.

Pour plus d’informations sur les exigences de l’application et du serveur hôtes Office, voir [Configuration requise pour exécuter des compléments Office](../../docs/overview/requirements-for-running-office-add-ins.md).


**Hôtes pris en charge par la plateforme**


||**Office pour Bureau Windows**|**Office Online (dans un navigateur)**|**Office pour iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||v||
|**Excel**|v|v|v|
|**Word**|v||v|

|||
|:-----|:-----|
|**Niveau d’autorisation minimal**|[Restricted](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Types de complément**|Application de contenu et de volet de tâches|
|**Bibliothèque**|Office.js|
|**Espace de noms**|Bureau|

## Historique de prise en charge





****


|**Version**|**Modifications**|
|:-----|:-----|
|1.1|Prise en charge supplémentaire d’Excel et de Word dans Office pour iPad.|
|1.1|Vous pouvez désormais ajouter et supprimer des gestionnaires d’événements pour l’événement **SelectionChanged** dans les compléments de contenu pour Access.|
|1.0|Introduit|
