

# Méthode ProjectDocument.setResourceFieldAsync
Définit de manière asynchrone la valeur du champ spécifié pour la ressource spécifiée.
 **Important :** cette API fonctionne uniquement dans Project 2016 sur le bureau Windows.

|||
|:-----|:-----|
|**Hôtes :**|Projet|
|**Disponible dans l’[ensemble de ressources requis](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Selection|
|**Ajouté dans**|1.1|

```js
Office.context.document.setResourceFieldAsync(resourceId, fieldId, fieldValue[, options][, callback]);
```


## Paramètres

_resourceId_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;GUID de la ressource. Obligatoire.
    
_fieldId_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;ID du champ cible, en tant que constante [ProjectResourceFields](../../reference/shared/projectresourcefields-enumeration.md) ou sa valeur entière correspondante. Obligatoire.
    
_fieldValue_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;Valeur du champ cible, au format **string**, **number**, **boolean** ou **object**. Obligatoire.
    
_options_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;Le [paramètre facultatif suivant](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods) :

&nbsp;&nbsp;&nbsp;&nbsp;_asyncContext_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Type : **array, boolean, null, number, object, string** ou **non défini**<br/></br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Élément défini par l’utilisateur de n’importe quel type qui est renvoyé dans l’objet [AsyncResult](../../reference/shared/asyncresult.md) sans être modifié. Facultatif.</br></br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Par exemple, vous pouvez transmettre l’argument _asyncContext_ en utilisant le format `{asyncContext: 'Some text'}` ou `{asyncContext: <object>}`.


_callback_<br/>
&nbsp;&nbsp;&nbsp;&nbsp;Type : **function**

&nbsp;&nbsp;&nbsp;&nbsp;Fonction appelée quand l’appel de méthode est renvoyé, dont le seul paramètre est de type [AsyncResult](../../reference/shared/asyncresult.md). Facultatif.

    

## Valeur de rappel

Lorsque la fonction _callback_ s’exécute, elle reçoit un objet [AsyncResult](../../reference/shared/asyncresult.md) accessible à partir du paramètre de la fonction de rappel.

Pour la méthode **setResourceFieldAsync**, l’objet [AsyncResult](../../reference/shared/asyncresult.md) renvoyé contient les propriétés suivantes.


|**Nom**|**Description**|
|:-----|:-----|
|[asyncContext](../../reference/shared/asyncresult.asynccontext.md)|Données transmises dans le paramètre _asyncContext_ facultatif si le paramètre a été utilisé.|
|[erreur](../../reference/shared/asyncresult.error.md)|Informations sur l’erreur, si la propriété **status** est **failed**.|
|[statut](../../reference/shared/asyncresult.status.md)|Statut **succeeded** ou **failed** de l’appel asynchrone.|
|[value](../../reference/shared/asyncresult.value.md)|Cette méthode ne renvoie pas de valeur.|

## Remarques

Appelez d’abord la méthode [getSelectedResourceAsync](../../reference/shared/projectdocument.getselectedtaskasync.md) ou [getResourceByIndexAsync](../../reference/shared/projectdocument.getresourcebyindexasync.md) pour obtenir le GUID de ressource, puis transmettez le GUID en tant qu’argument _resourceId_ à **setResourceFieldAsync**. Vous ne pouvez mettre à jour qu’un seul champ pour une seule ressource dans chaque appel asynchrone.


## Exemple

L’exemple de code suivant appelle [getSelectedResourceAsync](../../reference/shared/projectdocument.getselectedtaskasync.md) pour obtenir le GUID de la ressource actuellement sélectionnée dans un affichage des ressources. Ensuite, il définit deux valeurs de champ de ressource en appelant **setResourceFieldAsync** de manière récursive.

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
            $('#set-info').click(setResourceInfo);
        });
    };

    // Get the GUID of the resource, and then get the resource fields.
    function setResourceInfo() {
        getResourceGuid().then(
            function (data) {
                setResourceFields(data);
            }
        );
    }

    // Get the GUID of the selected resource.
    function getResourceGuid() {
        var defer = $.Deferred();
        Office.context.document.getSelectedResourceAsync(
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

    // Set the specified fields for the selected resource.
    function setResourceFields(resourceGuid) {
        var targetFields = [Office.ProjectResourceFields.StandardRate, Office.ProjectResourceFields.Notes];
        var fieldValues = [.28, 'Notes for the resource.'];

        // Set the field value. If the call is successful, set the next field.
        for (var i = 0; i < targetFields.length; i++) {
            Office.context.document.setResourceFieldAsync(
                resourceGuid,
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


[getSelectedResourceAsync](../../reference/shared/projectdocument.getselectedtaskasync.md)
[getResourceByIndexAsync](../../reference/shared/projectdocument.getresourcebyindexasync.md)
[Objet AsyncResult](../../reference/shared/asyncresult.md)
[Énumération ProjectResourceFields](../../reference/shared/projectresourcefields-enumeration.md)
[Objet ProjectDocument](../../reference/shared/projectdocument.projectdocument.md)

