
# Événement Binding.bindingDataChanged
Se produit quand des données sont modifiées dans la liaison.

|||
|:-----|:-----|
|**Hôtes :**|Access, Excel, Word|
|**Dernière modification dans BindingEvents**|1.1|

```js
Office.EventType.BindingDataChanged
```


## Remarques

Pour ajouter un gestionnaire d’événements à l’événement **BindingDataChanged** d’une liaison, utilisez la méthode [addHandlerAsync](../../reference/shared/binding.addhandlerasync.md) de l’objet **Binding**. Le gestionnaire d’événements reçoit un argument de type [BindingDataChangedEventArgs](../../reference/shared/binding.bindingdatachangedeventargs.md).


## Exemple




```js
function addEventHandlerToBinding() {
    Office.select("bindings#MyBinding").addHandlerAsync(Office.EventType.BindingDataChanged, onBindingDataChanged);
}

function onBindingDataChanged(eventArgs) {
    write("Data has changed in binding: " + eventArgs.binding.id);
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


## Informations de prise en charge


Un Y majuscule dans la matrice suivante indique que cette méthode est prise en charge dans l'application hôte Office correspondante. Une cellule vide indique que l'application hôte Office ne prend pas en charge cette méthode.

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

|**Version**|**Modifications**|
|:-----|:-----|
|1.1|Prise en charge supplémentaire d’Excel et de Word dans Office pour iPad.|
|1.1|Prise en charge supplémentaire de cet événement dans les compléments pour Access.|
|1.0|Introduit|
