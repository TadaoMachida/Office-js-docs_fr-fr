
# Événement Binding.bindingSelectionChanged
Se produit quand la sélection est modifiée dans la liaison.

|||
|:-----|:-----|
|**Hôtes :**|Access, Excel, Word|
|**Disponible dans l’[ensemble de ressources requis](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|BindingEvents|
|**Dernière modification dans la sélection**|1.1|

```
Office.EventType.BindingSelectionChanged
```

## Remarques

Pour ajouter un gestionnaire d’événements à l’événement **BindingSelectionChanged** d’une liaison, utilisez la méthode [addHandlerAsync](../../reference/shared/binding.addhandlerasync.md) de l’objet **Binding**. Le gestionnaire d’événements reçoit un argument de type [BindingSelectionChangedEventArgs](../../reference/shared/binding.bindingselectionchangedeventargs.md).


## Exemple




```
function addEventHandlerToBinding() {
 Office.select("bindings#MyBinding").addHandlerAsync(Office.EventType.BindingSelectionChanged, onBindingSelectionChanged);
}

function onBindingSelectionChanged(eventArgs) {
    write(eventArgs.binding.id + " has been selected.");
}
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


## Informations de prise en charge


Un Y majuscule dans la matrice suivante indique que cet événement est pris en charge dans l'application hôte Office correspondante. Une cellule vide indique que l'application hôte Office ne prend pas en charge cet événement.

Pour plus d’informations sur les exigences de l’application et du serveur hôtes Office, voir [Configuration requise pour exécuter des compléments pour Office](../../docs/overview/requirements-for-running-office-add-ins.md).


**Hôtes pris en charge par la plateforme**


||**Office pour Bureau Windows**|**Office Online (dans un navigateur)**|**Office pour iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||v||
|**Excel**|v|v|v|
|**Word**|v||v|

|||
|:-----|:-----|
|**Disponible dans les ensembles de ressources requis**|BindingEvents|
|**Types de complément**|Application de contenu et de volet de tâches|
|**Bibliothèque**|Office.js|
|**Espace de noms**|Bureau|

## Historique de prise en charge





****


|**Version**|**Modifications**|
|:-----|:-----|
|1.1|Prise en charge supplémentaire d’Excel et de Word dans Office pour iPad.|
|1.1|Prise en charge supplémentaire de cet événement dans les compléments pour Access.|
|1.0|Introduit|
