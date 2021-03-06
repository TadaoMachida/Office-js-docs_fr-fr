
# Propriété BindingSelectionChangedEventArgs.columnCount
Obtient le nombre de colonnes sélectionnées.

|||
|:-----|:-----|
|**Hôtes :**|Access, Excel, Word|
|**Dernière modification dans **|1.1|

```
var colCount = eventArgsObj.columnCount;
```


## Valeur renvoyée

Nombre de colonnes sélectionnées. Si une seule cellule est sélectionnée, retourne 1.


## Remarques

Si l’utilisateur effectue une sélection non contiguë, le décompte de la dernière sélection contiguë au sein de la liaison est retourné. 

En ce qui concerne Word, cette propriété ne fonctionne que pour les liaisons dont le [BindingType](../../reference/shared/bindingtype-enumeration.md) est « table ». Si la liaison est de type « matrix », une valeur **null** est retournée. En outre, l’appel échoue si le tableau contient des cellules fusionnées, car la structure du tableau doit être uniforme pour que cette propriété fonctionne correctement.


## Exemple

L’exemple suivant ajoute un gestionnaire d’événements pour l’événement [SelectionChanged](../../reference/shared/binding.bindingselectionchangedevent.md) à la liaison dont l’[id](../../reference/shared/binding.id.md) est `myTable`. Quand l’utilisateur modifie la sélection, le gestionnaire affiche les coordonnées de la première cellule de la sélection, ainsi que le nombre de lignes et de colonnes sélectionnées.


```js
function addSelectionHandler() {
    Office.context.document.bindings.getByIdAsync("myTable", function (result) {
        result.value.addHandlerAsync("bindingSelectionChanged", myHandler);
    });
}

// Display selection start coordinates and row/column count.
function myHandler(bArgs) {
    write("Selection start row/col: " + bArgs.startRow + "," + bArgs.startColumn);
    write("Selection row count: " + bArgs.rowCount);
    write("Selection col count: " + bArgs.columnCount);
}
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


## Informations de prise en charge


Un Y majuscule dans la matrice suivante indique que cette propriété est prise en charge dans l'application hôte Office correspondante. Une cellule vide indique que l'application hôte Office ne prend pas en charge cette propriété.

Pour plus d’informations sur les exigences de l’application et du serveur hôtes Office, voir [Configuration requise pour exécuter des compléments pour Office](../../docs/overview/requirements-for-running-office-add-ins.md).


**Hôtes pris en charge par la plateforme**


| |**Office pour Bureau Windows**|**Office Online (dans un navigateur)**|**Office pour iPad**|
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
