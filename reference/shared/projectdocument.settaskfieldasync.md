
# Méthode ProjectDocument.setTaskFieldAsync (interface API JavaScript pour Office version 1.1)
Définit de manière asynchrone la valeur du champ spécifié pour la tâche spécifiée.
 **Important :** cette API fonctionne uniquement dans Project 2016 sur le bureau Windows.

|||
|:-----|:-----|
|**Hôtes :**|Projet|
|**Disponible dans l’[ensemble de ressources requis](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Selection|
|**Ajouté dans**|1.1|

```js
Office.context.document.setTaskFieldAsync(taskId, fieldId, fieldValue[, options][, callback]);
```


## Paramètres


_taskId_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;GUID de la tâche. Obligatoire.<br/><br/>
_fieldId_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;ID du champ cible, en tant que constante [ProjectTaskFields](../../reference/shared/projecttaskfields-enumeration.md) ou sa valeur entière correspondante. Obligatoire.<br/><br/>
_fieldValue_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;Valeur du champ cible, au format **string**, **number**, **boolean** ou **object**. Obligatoire.<br/><br/>
_options_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;Le [paramètre facultatif suivant](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods) :<br/><br/>

&nbsp;&nbsp;&nbsp;&nbsp;_asyncContext_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Type : **array, boolean, null, number, object, string** ou **non défini**<br/></br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Élément défini par l’utilisateur de n’importe quel type qui est renvoyé dans l’objet [AsyncResult](../../reference/shared/asyncresult.md) sans être modifié. Facultatif.</br></br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Par exemple, vous pouvez transmettre l’argument _asyncContext_ en utilisant le format `{asyncContext: 'Some text'}` ou `{asyncContext: <object>}`.<br/><br/>
_callback_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;Type : **function**<br/><br/>
&nbsp;&nbsp;&nbsp;&nbsp;Fonction appelée quand l’appel de méthode est renvoyé, dont le seul paramètre est de type [AsyncResult](../../reference/shared/asyncresult.md). Facultatif.
    

## Valeur de rappel

Lorsque la fonction _callback_ s’exécute, elle reçoit un objet [AsyncResult](../../reference/shared/asyncresult.md) accessible à partir du paramètre de la fonction de rappel.

Pour la méthode **setTaskFieldAsync**, l’objet [AsyncResult](../../reference/shared/asyncresult.md) renvoyé contient les propriétés suivantes.



|**Nom**|**Description**|
|:-----|:-----|
|[asyncContext](../../reference/shared/asyncresult.asynccontext.md)|Données transmises dans le paramètre _asyncContext_ facultatif si le paramètre a été utilisé.|
|[erreur](../../reference/shared/asyncresult.error.md)|Informations sur l’erreur, si la propriété **status** est **failed**.|
|[statut](../../reference/shared/asyncresult.status.md)|Statut **succeeded** ou **failed** de l’appel asynchrone.|
|[value](../../reference/shared/asyncresult.value.md)|Cette méthode ne renvoie pas de valeur.|

## Remarques

Appelez d’abord la méthode [getSelectedTaskAsync](../../reference/shared/projectdocument.getselectedtaskasync.md) ou [getTaskByIndexAsync](../../reference/shared/projectdocument.settaskfieldasync.md) pour obtenir le GUID de tâche, puis transmettez le GUID en tant qu’argument _taskId_ à **setTaskFieldAsync**. Vous ne pouvez mettre à jour qu’un seul champ pour une seule tâche dans chaque appel asynchrone.


## Exemple

L’exemple de code suivant appelle [getSelectedTaskAsync](../../reference/shared/projectdocument.getselectedtaskasync.md) pour obtenir le GUID de la tâche actuellement sélectionnée dans un affichage des tâches. Ensuite, il définit deux valeurs de champ de tâche en appelant **setTaskFieldAsync** de façon récursive.

la méthode [getSelectedTaskAsync](../../reference/shared/projectdocument.getselectedtaskasync.md) utilisée dans l’exemple nécessite qu’un affichage des tâches (par exemple, Utilisation des tâches) soit la vue active et qu’une tâche soit sélectionnée. Voir la méthode [addHandlerAsync](../../reference/shared/projectdocument.addhandlerasync.md) pour obtenir un exemple qui permet d’activer un bouton en fonction du type de vue active.

L’exemple suppose que votre complément comporte une référence à la bibliothèque jQuery et que les contrôles de page suivants sont définis dans la balise div de contenu du corps de la page.




```HTML
<input id="set-info" type="button" value="Set info" /><br />
<span id="message"></span>
```




```js
(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            
            // After the DOM is loaded, add-in-specific code can run.
            app.initialize();
            $('#set-info').click(setTaskInfo);
        });
    };

    // Get the GUID of the task, and then get the task fields.
    function setTaskInfo() {
        getTaskGuid().then(
            function (data) {
                setTaskFields(data);
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

    // Set the specified fields for the selected task.
    function setTaskFields(taskGuid) {
        var targetFields = [Office.ProjectTaskFields.Active, Office.ProjectTaskFields.Notes];
        var fieldValues = [true, 'Notes for the task.'];

        // Set the field value. If the call is successful, set the next field.
        for (var i = 0; i < targetFields.length; i++) {
            Office.context.document.setTaskFieldAsync(
                taskGuid,
                targetFields[i],
                fieldValues[i],
                function (result) {
                    if (result.status === Office.AsyncResultStatus.Succeeded) {
                        i++;
                    }
                    else {
                        onError(result.error);
                    }
                }
            );
        }
        $('#message').html('Field values set');
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
|**Niveau d’autorisation minimal**|[WriteDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Types de complément**|Volet de tâches|
|**Bibliothèque**|Office.js|
|**Espace de noms**|Bureau|

## Historique de prise en charge



|**Version**|**Modifications**|
|:-----|:-----|
|1.1|Introduit|

## Voir aussi



#### Autres ressources


[Méthode getSelectedTaskAsync](../../reference/shared/projectdocument.getselectedresourceasync.md)
[getTaskByIndexAsync](../../reference/shared/projectdocument.settaskfieldasync.md)
[Objet AsyncResult ](../../reference/shared/asyncresult.md)
[Énumération ProjectTaskFields](../../reference/shared/projecttaskfields-enumeration.md)
[Objet ProjectDocument](../../reference/shared/projectdocument.projectdocument.md)
