
# Error, objet
Fournit des informations spécifiques sur une erreur qui s’est produite pendant une opération de données asynchrone.

|||
|:-----|:-----|
|**Hôtes :**|Access, Excel, Outlook, PowerPoint, Project, Word|
|**Dernière modification dans **|1.1|

```
asyncResult.error
```


## Membres


**Propriétés**


|**Nom**|**Description**|
|:-----|:-----|
|[code](../../reference/shared/error.code.md)|Obtient le code numérique de l’erreur.|
|[name](../../reference/shared/error.name.md)|Obtient le nom de l’erreur.|
|[message](../../reference/shared/error.message.md)|Obtient une description détaillée de l’erreur.|

## Remarques

L’objet **Error** est accessible à partir de l’objet [AsyncResult](../../reference/shared/asyncresult.md) renvoyé dans la fonction transmise en tant qu’argument _callback_ d’une opération de données asynchrone, par exemple la méthode [setSelectedDataAsync](../../reference/shared/document.setselecteddataasync.md) de l’objet **Document**.


## Exemple

L’exemple suivant utilise la méthode **setSelectedDataAsync** pour affecter au texte sélectionné la valeur « Hello World! ». En cas d’échec, les valeurs des propriétés **name** et **message** de l’objet **Error** sont affichées.


```js
function setText() {

    Office.context.document.setSelectedDataAsync("Hello World!", {},
        function (asyncResult) {
            if (asyncResult.status === "failed")
            var err = asyncResult.error; 
                write(err.name + ": " + err.message);
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

||**Office pour Bureau Windows**|**Office Online (dans un navigateur)**|**Office pour iPad**|**OWA pour périphériques**|**Outlook pour Mac**|
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



****


|**Version**|**Modifications**|
|:-----|:-----|
|1.1|Prise en charge supplémentaire d’Excel, de PowerPoint et de Word dans Office pour iPad.|
|1.1|Prise en charge supplémentaire des compléments de contenu pour Access.|
|1.0|Introduit|
