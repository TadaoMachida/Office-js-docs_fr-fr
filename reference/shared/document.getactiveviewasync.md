
# Méthode Document.getActiveViewAsync
 Renvoie l’état de l’affichage actuel de la présentation (modification ou lecture).

|||
|:-----|:-----|
|**Hôtes :** Excel, PowerPoint, Word|**Types de compléments :** Application de contenu et de volet Office|
|**Disponible dans l’[ensemble de ressources requis](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|ActiveView|
|**Ajouté dans ActiveView**|1.1|

```
Office.context.document.getActiveViewAsync([,options], callback);
```


## Paramètres



|**Nom**|**Type**|**Description**|**Notes de prise en charge**|
|:-----|:-----|:-----|:-----|
| _options_|**object**|Spécifie l’un des [paramètres facultatifs](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods) suivants||
| _asyncContext_|**tableau**, **booléen**, **null**, **numérique**, **objet**, **chaîne** ou **non défini**|Élément défini par l’utilisateur de n’importe quel type qui est renvoyé dans l’objet **AsyncResult** sans être modifié.||
| _callback_|**object**|Fonction appelée quand le rappel est renvoyé, dont le seul paramètre est de type **AsyncResult**.||

## Valeur de rappel

Quand la fonction que vous avez transmise au paramètre _callback_ s’exécute, elle reçoit un objet [AsyncResult](../../reference/shared/asyncresult.md) accessible à partir de l’unique paramètre de la fonction de rappel.

Dans la fonction de rappel transmise à la méthode **getActiveViewAsync**, la propriété [AsyncResult.value](../../reference/shared/asyncresult.value.md) renvoie l’état de la vue actuelle de la présentation. La valeur renvoyée peut être `edit` ou `read`. `edit` correspond à l’une des vues dans lesquelles vous pouvez modifier les diapositives, comme **Normal** ou **Mode Plan**. `read` correspond à **Diaporama** ou **Mode Lecture**.


## Remarques

Peut déclencher un événement lorsque l’affichage change.


## Exemple

Pour obtenir l’affichage de la présentation en cours, vous devez écrire une fonction de rappel qui renvoie cette valeur, comme dans l’exemple suivant.


-  **Transmettre une fonction de rappel anonyme** qui renvoie le type d’affichage au paramètre _callback_ de la méthode **getActiveViewAsync**.
    
-  **Afficher la valeur** sur la page du complément.
    

```js
function getFileView() {
    // Get whether the current view is edit or read.
    Office.context.document.getActiveViewAsync(function (asyncResult) {
        if (asyncResult.status == "failed") {
            showMessage("Action failed with error: " + asyncResult.error.message);
        }
        else {
            showMessage(asyncResult.value);
        }
    });
}
```




## Informations de prise en charge


Un Y majuscule dans la matrice suivante indique que cette méthode est prise en charge dans l'application hôte Office correspondante. Une cellule vide indique que l'application hôte Office ne prend pas en charge cette méthode.

Pour plus d’informations sur les exigences de l’application et du serveur hôtes Office, voir [Configuration requise pour exécuter des compléments pour Office](../../docs/overview/requirements-for-running-office-add-ins.md).


**Hôtes pris en charge par la plateforme**


||**Office pour Bureau Windows**|**Office Online (dans un navigateur)**|**Office pour iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|||v|
|**PowerPoint**|v|v|v|
|**Word**|||v|

|||
|:-----|:-----|
|**Disponible dans les ensembles de ressources requis**|ActiveView|
|**Ajouté dans ActiveView**|1.1|
|**Niveau d’autorisation minimal**|[Restricted](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Types de complément**|Application de contenu et de volet de tâches|
|**Bibliothèque**|Office.js|
|**Espace de noms**|Bureau|

## Historique de prise en charge





****


|**Version**|**Modifications**|
|:-----|:-----|
|1.1|Prise en charge supplémentaire d’Excel, de PowerPoint et de Word dans Office pour iPad.|
|1.1|Introduites|
