
# Méthode ProjectDocument.getMaxResourceIndexAsync (interface API JavaScript pour Office version 1.1)
Obtient de façon asynchrone l’index maximal de la collection de ressources dans le projet en cours.
 **Important :** cette API fonctionne uniquement dans Project 2016 sur le bureau Windows.

|||
|:-----|:-----|
|**Hôtes :**|Projet|
|**Disponible dans l’[ensemble de ressources requis](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Selection|
|**Ajouté dans**|1.1|

```js
Office.context.document.getMaxResourceIndexAsync([options][, callback]);
```


## Paramètres


_options_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;Le **[paramètre facultatif suivant](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods) :**<br/><br/>&nbsp;&nbsp;&nbsp;&nbsp;_asyncContext_<br/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Type : **array**, **boolean**, **null**, **number**, **object**, **string** ou **non défini**<br/><br/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Élément défini par l’utilisateur de n’importe quel type qui est renvoyé dans l’objet [AsyncResult](../../reference/shared/asyncresult.md) sans être modifié. Facultatif.<br/><br/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Par exemple, vous pouvez transmettre l’argument _asyncContext_ en utilisant le format `{asyncContext: 'Some text'}` ou `{asyncContext: <object>}`.

_callback_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;Type : **function**<br/>
&nbsp;&nbsp;&nbsp;&nbsp;Fonction appelée quand l’appel de méthode est renvoyé, dont le seul paramètre est de type [AsyncResult](../../reference/shared/asyncresult.md). Facultatif.
    

## Valeur de rappel

Lorsque la fonction _callback_ s’exécute, elle reçoit un objet [AsyncResult](../../reference/shared/asyncresult.md) accessible à partir du paramètre de la fonction de rappel.

Pour la méthode **getMaxResourceIndexAsync**, l’objet [AsyncResult](../../reference/shared/asyncresult.md) renvoyé contient les propriétés suivantes.



|**Nom**|**Description**|
|:-----|:-----|
|[asyncContext](../../reference/shared/asyncresult.asynccontext.md)|Données transmises dans le paramètre _asyncContext_ facultatif si le paramètre a été utilisé.|
|[erreur](../../reference/shared/asyncresult.error.md)|Informations sur l’erreur, si la propriété **status** est **failed**.|
|[statut](../../reference/shared/asyncresult.status.md)|Statut **succeeded** ou **failed** de l’appel asynchrone.|
|[value](../../reference/shared/asyncresult.value.md)|Nombre d’index maximal dans la collection de ressources du projet en cours.|

## Remarques

Vous pouvez utiliser la valeur renvoyée avec la méthode [getResourceByIndexAsync](../../reference/shared/projectdocument.getresourcebyindexasync.md) pour obtenir les GUID de ressource. Une collection de ressources ne contient pas de ressource au niveau de l’index 0.


## Exemple

L’exemple de code suivant appelle **getResourceTaskIndexAsync** pour obtenir l’index maximal de la collection de ressources dans le projet en cours. Ensuite, il utilise la valeur renvoyée et la méthode [getResourceByIndexAsync](../../reference/shared/projectdocument.getresourcebyindexasync.md) pour obtenir chaque GUID de ressource.

L’exemple suppose que votre complément comporte une référence à la bibliothèque jQuery et que les contrôles de page suivants sont définis dans la balise div de contenu du corps de la page.




```HTML
<input id="get-info" type="button" value="Get info" /><br />
<span id="message"></span>
```




```js
(function () {
    "use strict";
    var resourceGuids = [];

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {

            // After the DOM is loaded, add-in-specific code can run.
            app.initialize();
            $('#get-info').click(getResourceInfo);
        });
    };

    // Get the maximum resource index, and then get the resource GUIDs.
    function getResourceInfo() {
        getMaxResourceIndex().then(
            function (data) {
                getResourceGuids(data);
            }
        );
    }

    // Get the maximum index of the resources for the current project.
    function getMaxResourceIndex() {
        var defer = $.Deferred();
        Office.context.document.getMaxResourceIndexAsync(
            function (result) {
                if (result.status === Office.AsyncResultStatus.Failed) {
                    onError(result.error);
                }
                else {
                    defer.resolve(result.value);
                }
            }
        );
        return defer.promise();
    }

    // Get each resource GUID, and then display the GUIDs in the add-in.
    // There is no 0 index for resources, so start with index 1.
    function getResourceGuids(maxResourceIndex) {
        var defer = $.Deferred();
        for (var i = 1; i <= maxResourceIndex; i++) {
            getResourceGuid(i);
        }
        return defer.promise();
        function getResourceGuid(index) {
            Office.context.document.getResourceByIndexAsync(index,
                function (result) {
                    if (result.status === Office.AsyncResultStatus.Succeeded) {
                        resourceGuids.push(result.value);
                        if (index == maxResourceIndex) {
                            defer.resolve();
                            $('#message').html(resourceGuids.toString());
                        }
                    }
                    else {
                        onError(result.error);
                    }
                }
            );
        }
    }
    function onError(error) {
        app.showNotification(error.name + ' ' + error.code + ': ' + error.message);
    }
})();
```


## Informations de prise en charge


Un Y majuscule dans la matrice suivante indique que cette méthode est prise en charge dans l'application hôte Office correspondante. Une cellule vide indique que l'application hôte Office ne prend pas en charge cette méthode.

Pour plus d’informations sur les exigences de l’application et du serveur hôtes Office, voir [Configuration requise pour exécuter des compléments pour Office](../../docs/overview/requirements-for-running-office-add-ins.md).


||**Office pour Bureau Windows**|**Office Online (dans un navigateur)**|
|:-----|:-----|:-----|
|**Projet**|v||

|||
|:-----|:-----|
|**Disponible dans les ensembles de ressources requis**||
|**Niveau d’autorisation minimal**|[ReadDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Types de complément**|Volet de tâches|
|**Bibliothèque**|Office.js|
|**Espace de noms**|Bureau|

## Historique de prise en charge



****


|**Version**|**Modifications**|
|:-----|:-----|
|1.1|Introduit|

## Voir aussi



#### Autres ressources


[getResourceByIndexAsync](../../reference/shared/projectdocument.getresourcebyindexasync.md)

[Objet AsyncResult](../../reference/shared/asyncresult.md)

[ProjectDocument, objet](../../reference/shared/projectdocument.projectdocument.md)
