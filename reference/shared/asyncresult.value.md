
# Propriété AsyncResult.value
Obtient la charge utile ou le contenu de l’opération asynchrone, le cas échéant.

|||
|:-----|:-----|
|**Hôtes :**|Access, Excel, Outlook, PowerPoint, Project, Word|
|**Dernière modification dans **|1.1|

```js
var dataValue = asyncResult.value;
```


## Valeur renvoyée

Retourne la valeur de la demande au moment où l’appel asynchrone a été effectué. 


 >**Remarque** : le contenu renvoyé par la propriété **value** pour une méthode « Async » particulière varie selon la finalité et le contexte de cette méthode. Pour déterminer le contenu renvoyé par la propriété **value** pour une méthode « Async », voir la section « Valeur de rappel » de la rubrique relative à la méthode. Pour obtenir la liste complète des méthodes « Async », voir la section « Remarques » de la rubrique relative à l’objet [AsyncResult](../../reference/shared/asyncresult.md).


## Remarques

Vous accédez à l’objet **AsyncResult** dans la fonction transmise en tant qu’argument au paramètre _callback_ d’une méthode « Async ». C’est le cas par exemple pour les méthodes [getSelectedDataAsync](../../reference/shared/document.getselecteddataasync.md) et [setSelectedDataAsync](../../reference/shared/document.setselecteddataasync.md) de l’objet **Document**.


## Exemple




```js
function getData() {
    Office.context.document.getSelectedDataAsync(Office.CoercionType.Table, function(asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Failed) {
            write(asyncResult.error.message);
        }
        else {
            write(asyncResult.value);
        }
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

||**Office pour Bureau Windows**|**Office Online (dans un navigateur)**|**Office pour iPad**|**OWA pour périphériques**|**Office pour Mac**|
|:-----|:-----|:-----|:-----|:-----|:-----|
|**Access**||v||||
|**Excel**|v|v|v|||
|**Outlook**|v|v||v|v|
|**PowerPoint**|v|v|v|||
|**Projet**|v|||||
|**Word**|v|v|v|||

|||
|:-----|:-----|
|**Niveau d’autorisation minimal**|[Restricted](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Types de complément**|De contenu, de volet de tâche, Outlook|
|**Bibliothèque**|Office.js|
|**Espace de noms**|Bureau|

## Historique de prise en charge



|**Version**|**Modifications**|
|:-----|:-----|
|1.1|Prise en charge supplémentaire de PowerPoint Online.|
|1.1|Prise en charge supplémentaire d’Excel, de PowerPoint et de Word dans Office pour iPad.|
|1.1|Prise en charge supplémentaire des compléments pour Access.|
|1.0|Introduit|
