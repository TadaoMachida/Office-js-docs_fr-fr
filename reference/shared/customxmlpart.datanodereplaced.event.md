
# Événement CustomXmlPart.dataNodeReplaced
Se produit quand un nœud est remplacé.

|||
|:-----|:-----|
|**Hôtes :**|Word|
|**Disponible dans l’[ensemble de ressources requis](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|CustomXmlParts|
|**Dernière modification dans **|1.1|

```
Office.EventType.DataNodeReplaced
```


## Remarques

Pour ajouter un gestionnaire d’événements pour l’événement **dataNodeInserted**, utilisez la méthode [addHandlerAsync](../../reference/shared/customxmlpart.addhandlerasync.md) de l’objet **CustomXmlPart**.


## Exemple




```js
function addNodeReplacedEvent() {
    Office.context.document.customXmlParts.getByIdAsync("{3BC85265-09D6-4205-B665-8EB239A8B9A1}", function (result) {
        var xmlPart = result.value;
        xmlPart.addHandlerAsync(Office.EventType.DataNodeReplaced, function (eventArgs) {
            write("A node has been replaced.");
        });
    });
}
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message;
}
```




## Informations de prise en charge


Un Y majuscule dans la matrice suivante indique que cette méthode est prise en charge dans l'application hôte Office correspondante. Une cellule vide indique que l'application hôte Office ne prend pas en charge cette méthode.

Pour plus d’informations sur les exigences de l’application et du serveur hôtes Office, voir [Configuration requise pour exécuter des compléments pour Office](../../docs/overview/requirements-for-running-office-add-ins.md).

||**Office pour Bureau Windows**|**Office Online (dans un navigateur)**|**Office pour iPad**|
|:-----|:-----|:-----|:-----|
|**Word**|v||v|

|||
|:-----|:-----|
|**Disponible dans les ensembles de ressources requis**|CustomXmlParts|
|**Niveau d’autorisation minimal**|[ReadWriteDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Types de complément**|Volet de tâches|
|**Bibliothèque**|Office.js|
|**Espace de noms**|Bureau|

## Historique de prise en charge



****


|**Version**|**Modifications**|
|:-----|:-----|
|1.1|Prise en charge supplémentaire de Word dans Office pour iPad.|
|1.0|Introduit|
