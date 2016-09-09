

# Méthode ProjectDocument.getTaskFieldAsync
Obtient de manière asynchrone la valeur du champ spécifié pour la tâche spécifiée.

|||
|:-----|:-----|
|**Hôtes :**|Projet|
|**Disponible dans l’[ensemble de ressources requis](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Selection|
|**Ajouté dans**|1,0|

```js
Office.context.document.getTaskFieldAsync(taskId, fieldId[, options][, callback]);
```


## Paramètres



|**Nom**|**Type**|**Description**|**Notes de prise en charge**|
|:-----|:-----|:-----|:-----|
| _taskId_|**string**|GUID de la tâche. Obligatoire.||
| _fieldId_|[ProjectTaskFields](../../reference/shared/projecttaskfields-enumeration.md)|ID du champ cible. Obligatoire.||
| _options_|**object**|Spécifie l’un des [paramètres facultatifs](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods) suivants.||
| _asyncContext_|**tableau**, **booléen**, **null**, **numérique**, **objet**, **chaîne** ou **non défini**|Élément défini par l’utilisateur de n’importe quel type qui est renvoyé dans l’objet **AsyncResult** sans être modifié.||
| _callback_|**object**|Fonction appelée quand le rappel est renvoyé, dont le seul paramètre est de type **AsyncResult**.||

## Valeur de rappel

Lorsque la fonction _callback_ s’exécute, elle reçoit un objet [AsyncResult](../../reference/shared/asyncresult.md) accessible à partir du paramètre de la fonction de rappel.

Pour la méthode **getTaskFieldAsync**, l’objet [AsyncResult](../../reference/shared/asyncresult.md) renvoyé contient les propriétés suivantes.



|**Nom**|**Description**|
|:-----|:-----|
|[asyncContext](../../reference/shared/asyncresult.asynccontext.md)|Données transmises dans le paramètre _asyncContext_ facultatif si le paramètre a été utilisé.|
|[erreur](../../reference/shared/asyncresult.error.md)|Informations sur l’erreur, si la propriété **status** est **failed**.|
|[statut](../../reference/shared/asyncresult.status.md)|Statut **succeeded** ou **failed** de l’appel asynchrone.|
|[value](../../reference/shared/asyncresult.value.md)|Contient la propriété **fieldValue**, qui représente la valeur du champ spécifié.|

## Remarques

Appelez d’abord la méthode [getSelectedTaskAsync](../../reference/shared/projectdocument.getselectedtaskasync.md) pour obtenir le GUID de tâche, puis transmettez-le comme argument _taskId_ à **getTaskFieldAsync**. Si l’affichage actif n’est pas un affichage des tâches (par exemple, un diagramme de Gantt ou un affichage Utilisation des tâches) ou si aucune tâche n’est sélectionnée dans un affichage des tâches, [getSelectedTaskAsync](../../reference/shared/projectdocument.getselectedtaskasync.md) renvoie une erreur 5001 (erreur interne). Voir [Méthode addHandlerAsync](../../reference/shared/projectdocument.addhandlerasync.md) pour consulter un exemple qui utilise l’événement [ViewSelectionChanged](../../reference/shared/projectdocument.viewselectionchanged.event.md) et la méthode [getSelectedViewAsync](../../reference/shared/projectdocument.getselectedviewasync.md) pour activer un bouton en fonction du type d’affichage actif.


## Exemple

L’exemple de code suivant appelle [getSelectedTaskAsync](../../reference/shared/projectdocument.getselectedtaskasync.md) pour obtenir le GUID de la tâche actuellement sélectionnée dans un affichage des tâches. Ensuite, il obtient trois valeurs de champ de tâche en appelant **getTaskFieldAsync** de manière récursive.

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

    // Get the GUID of the task, and then get the task fields.
    function getTaskInfo() {
        getTaskGuid().then(
            function (data) {
                getTaskFields(data);
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

    // Get the specified fields for the selected task.
    function getTaskFields(taskGuid) {
        var output = '';
        var targetFields = [Office.ProjectTaskFields.Priority, Office.ProjectTaskFields.PercentComplete];
        var fieldValues = ['Priority: ', '% Complete: '];
        var index = 0;
        getField();

        // Get each field, and then display the field values in the add-in.
        function getField() {
            if (index == targetFields.length) {
                for (var i = 0; i < fieldValues.length; i++) {
                    output += fieldValues[i] + '<br />';
                }
                $('#message').html(output);
            }

            // Get the field value. If the call is successful, then get the next field.
            else {
                Office.context.document.getTaskFieldAsync(
                    taskGuid,
                    targetFields[index],
                    function (result) {
                        if (result.status === Office.AsyncResultStatus.Succeeded) {
                            fieldValues[index] += result.value.fieldValue;
                            getField(index++);
                        }
                        else {
                            onError(result.error);
                        }
                    }
                );
            }
        }
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
|**Disponible dans les ensembles de ressources requis**||
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


[Méthode getSelectedTaskAsync](../../reference/shared/projectdocument.getselectedresourceasync.md)
[Objet AsyncResult](../../reference/shared/asyncresult.md)
[Énumération ProjectTaskFields](../../reference/shared/projecttaskfields-enumeration.md)
[Objet ProjectDocument](../../reference/shared/projectdocument.projectdocument.md)
