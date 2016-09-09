

# Méthode ProjectDocument.removeHandlerAsync
Supprime de manière asynchrone un gestionnaire d’événements pour un événement de changement de sélection de tâche dans un objet [ProjectDocument](../../reference/shared/projectdocument.projectdocument.md).

|||
|:-----|:-----|
|**Hôtes :**|Projet|
|**Disponible dans l’[ensemble de ressources requis](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Selection|
|**Ajouté dans**|1,0|

```js
Office.context.document.removeHandlerAsync(eventType[, options][, callback]);
```


## Paramètres
|**Nom**|**Type**|**Description**|**Notes de prise en charge**|
|:-----|:-----|:-----|:-----|
|_eventType_|[EventType](../../reference/shared/eventtype-enumeration.md)|Type d’événement à supprimer, comme une constante [EventType](../../reference/shared/eventtype-enumeration.md) ou sa valeur de texte correspondante. Obligatoire.<br/><br/>Le tableau suivant présente les arguments eventType valides pour un objet [ProjectDocument](../../reference/shared/projectdocument.projectdocument.md).<br/><br/><table><tr><th>Énumération</th><th>Valeur texte</th></tr><tr><td>
  <a href="https://msdn.microsoft.com/en-us/library/office/fp179836.aspx">Office.EventType.ResourceSelectionChanged</a></td><td>resourceSelectionChanged</td></tr><tr><td>
  <a href="https://msdn.microsoft.com/en-us/library/office/fp179816.aspx">Office.EventType.TaskSelectionChanged</a></td><td>taskSelectionChanged</td></tr><tr><td>
  <a href="https://msdn.microsoft.com/en-us/library/office/fp179839.aspx">Office.EventType.ViewSelectionChanged</a></td><td>viewSelectionChanged</td></tr></table>||
|_options_|**object**|Spécifie l’un des [paramètres facultatifs](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods) suivants.||
|_asyncContext_|**tableau**, **booléen**, **null**, **numérique**, **objet**, **chaîne** ou **non défini**|Élément défini par l’utilisateur de n’importe quel type qui est renvoyé dans l’objet **AsyncResult** sans être modifié.||
|_callback_|**object**|Fonction appelée quand le rappel est renvoyé, dont le seul paramètre est de type **AsyncResult**.||


## Valeur de rappel

Lorsque la fonction _callback_ s’exécute, elle reçoit un objet [AsyncResult](../../reference/shared/asyncresult.md) accessible à partir du paramètre de la fonction de rappel.

Pour la méthode **removeHandlerAsync**, l’objet [AsyncResult](../../reference/shared/asyncresult.md) renvoyé contient les propriétés suivantes.


|**Nom**|**Description**|
|:-----|:-----|
|[asyncContext](../../reference/shared/asyncresult.asynccontext.md)|Données transmises dans le paramètre _asyncContext_ facultatif si le paramètre a été utilisé.|
|[erreur](../../reference/shared/asyncresult.error.md)|Informations sur l’erreur, si la propriété **status** est **failed**.|
|[statut](../../reference/shared/asyncresult.status.md)|Statut **succeeded** ou **failed** de l’appel asynchrone.|
|[value](../../reference/shared/asyncresult.value.md)|**removeHandlerAsync** renvoie toujours **undefined**.|

## Exemple

L’exemple de code suivant utilise [addHandlerAsync](../../reference/shared/projectdocument.addhandlerasync.md) pour ajouter un gestionnaire d’événements pour l’événement [ResourceSelectionChanged](../../reference/shared/projectdocument.resourceselectionchanged.event.md) et **removeHandlerAsync** pour supprimer le gestionnaire.

Lorsqu’une ressource est sélectionnée dans un affichage de ressource, le gestionnaire affiche le GUID de ressource. Lorsque le gestionnaire est supprimé, le GUID n’est pas affiché.

L’exemple suppose que votre complément comporte une référence à la bibliothèque jQuery et que le contrôle de page suivant est défini dans la balise div de contenu du corps de la page.




```HTML
<input id="remove-handler" type="button" value="Remove handler" /><br />
<span id="message"></span>
```




```js
(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {

            // After the DOM is loaded, add-in-specific code can run.
            Office.context.document.addHandlerAsync(
                Office.EventType.ResourceSelectionChanged,
                getResourceGuid);
            $('#remove-handler').click(removeEventHandler);
        });
    };

    // Remove the event handler.
    function removeEventHandler() {
        Office.context.document.removeHandlerAsync(
            Office.EventType.ResourceSelectionChanged,
            {handler:getResourceGuid,
            asyncContext:'The handler is removed.'},
            function (result) {
                if (result.status === Office.AsyncResultStatus.Failed) {
                    onError(result.error);
                }
                else {
                    $('#remove-handler').attr('disabled', 'disabled');
                    $('#message').html(result.asyncContext);
                }
            }
        );
    }

    // Get the GUID of the currently selected resource and display it in the add-in.
    function getResourceGuid() {
        Office.context.document.getSelectedResourceAsync(
            function (result) {
                if (result.status === Office.AsyncResultStatus.Failed) {
                    onError(result.error);
                }
                else {
                    $('#message').html('Resource GUID: ' + result.value);
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
|**Niveau d’autorisation minimal**|[ReadWriteDocument](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Types de complément**|Volet de tâches|
|**Bibliothèque**|Office.js|
|**Espace de noms**|Bureau|

## Historique de prise en charge

|**Version**|**Modifications**|
|:-----|:-----|
|1,0|Introduit|

## Voir aussi



#### Autres ressources


[Méthode addHandlerAsync](../../reference/shared/projectdocument.addhandlerasync.md)
[Énumération EventType](../../reference/shared/eventtype-enumeration.md)
[Objet ProjectDocument](../../reference/shared/projectdocument.projectdocument.md)

