
# Événement Document.SelectionChanged
Se produit quand la sélection change dans le document.

|||
|:-----|:-----|
|**Hôtes :**|Excel, PowerPoint, Word|
|**Nouveauté de**|1.1|

```
Office.EventType.DocumentSelectionChanged
```

## Remarques

Pour ajouter un gestionnaire d’événements pour l’événement **SelectionChanged** d’un document, utilisez la méthode [addHandlerAsync](../../reference/shared/document.addhandlerasync.md) de l’objet **Document**.


## Exemple




```
function addEventHandlerToDocument() {
    Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, MyHandler);
}

function MyHandler(eventArgs) {
    doSomethingWithDocument(eventArgs.document);
}

```




## Informations de prise en charge


Un Y majuscule dans la matrice suivante indique que cette méthode est prise en charge dans l'application hôte Office correspondante. Une cellule vide indique que l'application hôte Office ne prend pas en charge cette méthode.

Pour plus d’informations sur les exigences de l’application et du serveur hôtes Office, voir [Configuration requise pour exécuter des compléments pour Office](../../docs/overview/requirements-for-running-office-add-ins.md).


**Hôtes pris en charge par la plateforme**


||**Office pour Bureau Windows**|**Office Online (dans un navigateur)**|**Office pour iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|v|v|v|
|**PowerPoint**|v|v|v|
|**Word**|v|v|v|

|||
|:-----|:-----|
|**Types de complément**|Application de contenu et de volet de tâches|
|**Bibliothèque**|Office.js|
|**Espace de noms**|Bureau|

## Historique de prise en charge



****


|**Version**|**Modifications**|
|:-----|:-----|
|1.1|Prise en charge supplémentaire d’Excel, de PowerPoint et de Word dans Office pour iPad.|
|1.0|Introduit|
