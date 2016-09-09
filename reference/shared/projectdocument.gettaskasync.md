

# Méthode ProjectDocument.getTaskAsync
Obtient de manière asynchrone le nom de la tâche spécifiée, les ressources affectées et l’ID de la tâche dans la liste de tâches SharePoint synchronisée.

|||
|:-----|:-----|
|**Hôtes :**|Projet|
|**Disponible dans l’[ensemble de ressources requis](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Selection|
|**Ajouté dans**|1,0|

```js
Office.context.document.getTaskAsync(taskId [,options][, callback]);
```


## Paramètres



|**Nom**|**Type**|**Description**|**Notes de prise en charge**|
|:-----|:-----|:-----|:-----|
| _taskId_|**string**|GUID de la tâche. Obligatoire.||
| _options_|**object**|Spécifie l’un des [paramètres facultatifs](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods) suivants.||
| _asyncContext_|**tableau**, **booléen**, **null**, **numérique**, **objet**, **chaîne** ou **non défini**|Élément défini par l’utilisateur de n’importe quel type qui est renvoyé dans l’objet **AsyncResult** sans être modifié.||
| _callback_|**object**|Fonction appelée quand le rappel est renvoyé, dont le seul paramètre est de type **AsyncResult**.||

## Valeur de rappel

Lorsque la fonction _callback_ s’exécute, elle reçoit un objet [AsyncResult](../../reference/shared/asyncresult.md) accessible à partir du paramètre de la fonction de rappel.

Pour la méthode **getTaskAsync**, l’objet [AsyncResult](../../reference/shared/asyncresult.md) renvoyé contient les propriétés suivantes.


****


|**Nom**|**Description**|
|:-----|:-----|
|[asyncContext](../../reference/shared/asyncresult.asynccontext.md)|Données transmises dans le paramètre _asyncContext_ facultatif si le paramètre a été utilisé.|
|[erreur](../../reference/shared/asyncresult.error.md)|Informations sur l’erreur, si la propriété **status** est **failed**.|
|[statut](../../reference/shared/asyncresult.status.md)|Statut **succeeded** ou **failed** de l’appel asynchrone.|
|[value](../../reference/shared/asyncresult.value.md)|Contient les propriétés suivantes :<br/><br/><ul><li><b>taskName</b> : nom de la tâche.</li><li><b>wssTaskId</b> : ID de la tâche dans la liste de tâches SharePoint synchronisée. Si le projet n’est pas synchronisé avec une liste de tâches SharePoint, la valeur est <b>0</b>.</li><li><b>resourceNames</b> : liste délimitée par des virgules des noms des ressources affectées à la tâche.</li></ul>|

## Remarques

Avant d’appeler la méthode **getTaskAsync**, appelez la méthode [getSelectedTaskAsync](../../reference/shared/projectdocument.getselectedtaskasync.md) pour obtenir le GUID de la tâche.


## Exemple

L’exemple de code suivant appelle [getSelectedTaskAsync](../../reference/shared/projectdocument.getselectedtaskasync.md) pour obtenir le GUID de la tâche sélectionnée. Ensuite, il appelle **getTaskAsync** pour obtenir les propriétés de la tâche qui sont disponibles à partir de l’API JavaScript pour Office.

L’exemple suppose que votre complément comporte une référence à la bibliothèque jQuery et que les contrôles de page suivants sont définis dans la balise div de contenu du corps de la page.




```HTML
<input id="get-info" type="button" value="Get info" /><br />
<span id="message"></span>
```




```js
(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {

            // After the DOM is loaded, add-in-specific code can run.
            $('#get-info').click(getTaskInfo);
        });
    };

    // Get the GUID of the task, and then get local task properties.
    function getTaskInfo() {
        getTaskGuid().then(
            function (data) {
                getTaskProperties(data);
            }
        );
    }

    // Get the GUID of the selected task.
    function getTaskGuid() {
        var defer = $.Deferred();
        Office.context.document.getSelectedTaskAsync(
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

    // Get local properties for the selected task, and then display it in the add-in.
    function getTaskProperties(taskGuid) {
        Office.context.document.getTaskAsync(
            taskGuid,
            function (result) {
                if (result.status === Office.AsyncResultStatus.Failed) {
                    onError(result.error);
                }
                else {
                    var taskInfo = result.value;
                    var output = String.format(
                        'Name: {0}<br/>GUID: {1}<br/>SharePoint task ID: {2}<br/>Resource names: {3}',
                        taskInfo.taskName, taskGuid, taskInfo.wssTaskId, taskInfo.resourceNames);
                    $('#message').html(output);
                }
            }
        );
    }

    function onError(error) {
        $('#message').html(error.name + ' ' + error.code + ': ' + error.message);
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
|**Disponible dans les ensembles de ressources requis**|Selection|
|**Niveau d’autorisation minimal**|[ReadDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Types de complément**|Volet de tâches|
|**Bibliothèque**|Office.js|
|**Espace de noms**|Bureau|

## Historique de prise en charge



|**Version**|**Modifications**|
|:-----|:-----|
|1,0|Introduit|

## Voir aussi



#### Autres ressources


[Méthode getSelectedTaskAsync](../../reference/shared/projectdocument.getselectedtaskasync.md)
[Objet AsyncResult](../../reference/shared/asyncresult.md)
[Objet ProjectDocument](../../reference/shared/projectdocument.projectdocument.md)
