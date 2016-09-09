

# Méthode ProjectDocument.getSelectedViewAsync
Obtient de manière asynchrone le type et le nom de l’affichage actif dans le document.

|||
|:-----|:-----|
|**Hôtes :**|Projet|
|**Disponible dans l’[ensemble de ressources requis](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Selection|
|**Ajouté dans**|1,0|

```js
Office.context.document.getSelectedViewAsync([options,] [callback]);
```


## Paramètres



|**Nom**|**Type**|**Description**|**Notes de prise en charge**|
|:-----|:-----|:-----|:-----|
| _options_|**object**|Spécifie l’un des [paramètres facultatifs](../../docs/develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods) suivants.||
| _asyncContext_|**tableau**, **booléen**, **null**, **numérique**, **objet**, **chaîne** ou **non défini**|Élément défini par l’utilisateur de n’importe quel type qui est renvoyé dans l’objet **AsyncResult** sans être modifié.||
| _callback_|**object**|Fonction appelée quand le rappel est renvoyé, dont le seul paramètre est de type **AsyncResult**.||

## Valeur de rappel

Lorsque la fonction _callback_ s’exécute, elle reçoit un objet [AsyncResult](../../reference/shared/asyncresult.md) accessible à partir du paramètre de la fonction de rappel.

Pour la méthode **getSelectedViewAsync**, l’objet [AsyncResult](../../reference/shared/asyncresult.md) renvoyé contient les propriétés suivantes.


****


|**Nom**|**Description**|
|:-----|:-----|
|[asyncContext](../../reference/shared/asyncresult.asynccontext.md)|Données transmises dans le paramètre _asyncContext_ facultatif si le paramètre a été utilisé.|
|[erreur](../../reference/shared/asyncresult.error.md)|Informations sur l’erreur, si la propriété **status** est **failed**.|
|[statut](../../reference/shared/asyncresult.status.md)|Statut **succeeded** ou **failed** de l’appel asynchrone.|
|[value](../../reference/shared/asyncresult.value.md)|Contient les propriétés suivantes :<br/><br/><div>* **viewName** : nom de l’affichage, sous la forme d’une constante [ProjectViewTypes](../../reference/shared/projectviewtypes-enumeration.md).<br/>* **viewType** : type d’affichage, sous la forme d’une valeur entière d’une constante [ProjectViewTypes](../../reference/shared/projectviewtypes-enumeration.md).</div>|

## Exemple

L’exemple de code suivant ajoute un gestionnaire d’événements [ViewSelectionChanged](../../reference/shared/projectdocument.viewselectionchanged.event.md) qui appelle **getSelectedViewAsync** pour obtenir le nom et le type de l’affichage actif dans le document.

L’exemple suppose que votre complément comporte une référence à la bibliothèque jQuery et que le contrôle de page suivant est défini dans la balise div de contenu du corps de la page.




```HTML
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
                Office.EventType.ViewSelectionChanged,
                getActiveView);
            getActiveView();
        });
    };

    // Get the active view's name and type.
    function getActiveView() {
        Office.context.document.getSelectedViewAsync(
            function (result) {
                if (result.status === Office.AsyncResultStatus.Failed) {
                    onError(result.error);
                }
                else {
                    var output = String.format(
                        'View name: {0}<br/>View type: {1}',
                        result.value.viewName, viewType);
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


[Énumération ProjectViewTypes](../../reference/shared/projectviewtypes-enumeration.md)
[AsyncResult, objet](../../reference/shared/asyncresult.md)
[Événement ViewSelectionChanged](../../reference/shared/projectdocument.viewselectionchanged.event.md)
[Objet ProjectDocument](../../reference/shared/projectdocument.projectdocument.md)
