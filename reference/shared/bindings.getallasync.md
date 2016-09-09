
# Méthode Bindings.getAllAsync
Obtient toutes les liaisons qui ont été créées précédemment.

|||
|:-----|:-----|
|**Hôtes :**|Access, Excel, Word|
|**Disponible dans l’[ensemble de ressources requis](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|MatrixBindings, TableBindings, TextBindings|
|**Dernière modification**|1.1|

```
bindingsObj.getAllAsync([options], callback);
```


## Paramètres



|**Nom**|**Type**|**Description**|**Notes de prise en charge**|
|:-----|:-----|:-----|:-----|
| _options_|**object**|Spécifie l’un des [paramètres facultatifs](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods) suivants||
| _asyncContext_|**tableau**, **booléen**, **null**, **numérique**, **objet**, **chaîne** ou **non défini**|Élément défini par l’utilisateur de n’importe quel type qui est renvoyé dans l’objet **AsyncResult** sans être modifié.||
| _callback_|**object**|Fonction appelée quand le rappel est renvoyé, dont le seul paramètre est de type **AsyncResult**.||

## Valeur de rappel

Quand la fonction que vous avez transmise au paramètre _callback_ s’exécute, elle reçoit un objet [AsyncResult](../../reference/shared/asyncresult.md) accessible à partir de l’unique paramètre de la fonction de rappel.

Dans la fonction de rappel transmise à la méthode **getAllAsync**, vous pouvez utiliser les propriétés de l’objet **AsyncResult** pour renvoyer les informations suivantes.



|**Propriété**|**Utiliser pour...**|
|:-----|:-----|
|[AsyncResult.value](../../reference/shared/asyncresult.value.md)|Accéder à un tableau contenant toutes les liaisons créées pour l’objet [Bindings](../../reference/shared/bindings.bindings.md) référencé.|
|[AsyncResult.status](../../reference/shared/asyncresult.status.md)|Déterminer si l’opération a réussi ou échoué.|
|[AsyncResult.error](../../reference/shared/asyncresult.error.md)|Accéder à un objet [Error](../../reference/shared/error.md) fournissant des informations sur l’erreur en cas d’échec de l’opération.|
|[AsyncResult.asyncContext](../../reference/shared/asyncresult.asynccontext.md)|Accéder à votre valeur ou **objet** défini par l’utilisateur, si vous en avez transmis un en tant que paramètre _asyncContext_.|

## Exemple




```js
function displayAllBindingNames() {
    Office.context.document.bindings.getAllAsync(function (asyncResult) {
        var bindingString = '';
        for (var i in asyncResult.value) {
            bindingString += asyncResult.value[i].id + '\n';
        }
        write('Existing bindings: ' + bindingString);
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
|**Access**||v||
|**Excel**|v|v|v|
|**Word**|v||v|

|||
|:-----|:-----|
|**Disponible dans les ensembles de ressources requis**|MatrixBindings, TableBindings, TextBindings|
|**Niveau d’autorisation minimal**|[Restricted](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Types de complément**|Application de contenu et de volet de tâches|
|**Bibliothèque**|Office.js|
|**Espace de noms**|Bureau|

## Historique de prise en charge




|**Version**|**Modifications**|
|:-----|:-----|
|1.1|Prise en charge supplémentaire d’Excel et de Word dans Office pour iPad.|
|1.1|Prise en charge supplémentaire de la liaison de tableau dans les compléments de contenu pour Access. |
|1.0|Introduit|
