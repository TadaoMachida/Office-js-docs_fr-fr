
# Méthode ProjectDocument.getSelectedTaskAsync
Obtient de manière asynchrone le GUID de la tâche sélectionnée dans un affichage des tâches.

|||
|:-----|:-----|
|**Hôtes :**|Projet|
|**Disponible dans l’[ensemble de ressources requis](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Selection|
|**Ajouté dans**|1,0|

```
Office.context.document.getSelectedTaskAsync([options,] [callback]);
```


## Paramètres



|**Nom**|**Type**|**Description**|**Notes de prise en charge**|
|:-----|:-----|:-----|:-----|
| _options_|**object**|Spécifie l’un des [paramètres facultatifs](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods) suivants.||
| _asyncContext_|**tableau**, **booléen**, **null**, **numérique**, **objet**, **chaîne** ou **non défini**|Élément défini par l’utilisateur de n’importe quel type qui est renvoyé dans l’objet **AsyncResult** sans être modifié.||
| _callback_|**object**|Fonction appelée quand le rappel est renvoyé, dont le seul paramètre est de type **AsyncResult**.||

## Valeur de rappel

Lorsque la fonction _callback_ s’exécute, elle reçoit un objet [AsyncResult](../../reference/shared/asyncresult.md) accessible à partir du paramètre de la fonction de rappel.

Pour la méthode **getSelectedTaskAsync**, l’objet [AsyncResult](../../reference/shared/asyncresult.md) renvoyé contient les propriétés suivantes :


****


|**Name**|**Description**|
|:-----|:-----|
|[asyncContext](../../reference/shared/asyncresult.asynccontext.md)|Données transmises dans le paramètre _asyncContext_ facultatif si le paramètre a été utilisé.|
|[erreur](../../reference/shared/asyncresult.error.md)|Informations sur l’erreur, si la propriété **status** est **failed**.|
|[statut](../../reference/shared/asyncresult.status.md)|Statut **succeeded** ou **failed** de l’appel asynchrone.|
|[value](../../reference/shared/asyncresult.value.md)|GUID de la tâche sélectionnée au format **chaîne**.|

## Remarques

Le GUID de tâche est plus utile dans les compléments Project que le numéro d’ID de la tâche (par exemple, l’ID de la première tâche dans le diagramme de Gantt est **1**). Le GUID de tâche peut être utilisé pour accéder aux informations sur la tâche Project, par exemple pour les tâches d’un projet SharePoint qui sont synchronisées avec Project Server en mode Visibilité. Vous pouvez également enregistrer le GUID de tâche dans une variable locale et l’utiliser pour les méthodes [getTaskAsync](../../reference/shared/projectdocument.gettaskasync.md) et [getTaskFieldAsync](../../reference/shared/projectdocument.gettaskfieldasync.md)

Si l’affichage actif n’est pas un affichage des tâches (par exemple, un affichage Diagramme de Gantt ou Utilisation des tâches) ou si aucune tâche n’est sélectionnée dans un affichage des tâches, **getSelectedTaskAsync** renvoie une erreur 5001 (erreur interne). Voir [Méthode addHandlerAsync](../../reference/shared/projectdocument.addhandlerasync.md) pour obtenir un exemple qui utilise l’événement [ViewSelectionChanged](../../reference/shared/projectdocument.viewselectionchanged.event.md) et la méthode [getSelectedViewAsync](../../reference/shared/projectdocument.getselectedviewasync.md)pour activer un bouton en fonction du type d’affichage actif.


## Exemple

L’exemple de code suivant appelle **getSelectedTaskAsync** pour obtenir le GUID de la tâche qui est actuellement sélectionnée dans un affichage des tâches. Il obtient ensuite les propriétés de la tâche en appelant [getTaskAsync](../../reference/shared/projectdocument.gettaskasync.md).

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

    // // Get the GUID of the task, and then get local task properties.
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



****


|**Version**|**Modifications**|
|:-----|:-----|
|1,0|Introduit|

## Voir aussi



#### Autres ressources


[Méthode getTaskAsync](../../reference/shared/projectdocument.gettaskasync.md)

[AsyncResult, objet](../../reference/shared/asyncresult.md)

[ProjectDocument, objet](../../reference/shared/projectdocument.projectdocument.md)
