

# Événement ProjectDocument.ViewSelectionChanged
Se produit quand l’affichage actif change dans le projet actif.

|||
|:-----|:-----|
|**Hôtes :**|Projet|
|**Disponible dans l’[ensemble de ressources requis](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Selection|
|**Ajouté dans**|1,0|

```js
Office.EventType.ViewSelectionChanged
```


## Remarques

 **ViewSelectionChanged** est une constante d’énumération [EventType](../../reference/shared/eventtype-enumeration.md) pouvant être utilisée dans les méthodes [ProjectDocument.addHandlerAsync](../../reference/shared/projectdocument.addhandlerasync.md) et [ProjectDocument.removeHandlerAsync](../../reference/shared/projectdocument.removehandlerasync.md) pour ajouter ou supprimer un gestionnaire pour l’événement.


## Exemple

L’exemple de code suivant ajoute un gestionnaire pour l’événement **ViewSelectionChanged**. Lorsque l’affichage actif change, il obtient le nom et le type de l’affichage actif.

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

    // Get the name and type of the active view and display it in the add-in.
    function getActiveView() {
        Office.context.document.getSelectedViewAsync(
            function (result) {
                if (result.status === Office.AsyncResultStatus.Failed) {
                    onError(result.error);
                }
                else {
                    var output = String.format(
                        'View name: {0}<br/>View type: {1}',
                        result.value.viewName, result.value.viewType);
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

Pour obtenir un exemple qui montre comment utiliser un gestionnaire d’événements **ViewSelectionChanged** dans un complément Projet, voir l’article expliquant comment [créer votre premier complément du volet Office pour Project 2013 à l’aide d’un éditeur de texte](../../docs/project/create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md).


## Informations de prise en charge


Un Y majuscule dans la matrice suivante indique que cet événement est pris en charge dans l'application hôte Office correspondante. Une cellule vide indique que l'application hôte Office ne prend pas en charge cet événement.

Pour plus d’informations sur les exigences de l’application et du serveur hôtes Office, voir [Configuration requise pour exécuter des compléments pour Office](../../docs/overview/requirements-for-running-office-add-ins.md).



||**Office pour Bureau Windows**|**Office Online (dans un navigateur)**|
|:-----|:-----|:-----|
|**Projet**|v||

|||
|:-----|:-----|
|**Disponible dans les ensembles de ressources requis**||
|**Niveau d’autorisation minimal**|[Restricted](../../docs/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)|
|**Types de complément**|Volet de tâches|
|**Bibliothèque**|Office.js|
|**Espace de noms**|Bureau|

## Historique de prise en charge



|**Version**|**Modifications**|
|:-----|:-----|
|1,0|Introduit|

## Voir aussi



#### Autres ressources


[Création de votre premier complément du volet Office pour Project 2013 à l’aide d’un éditeur de texte](../../docs/project/create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md)
[Énumération EventType](../../reference/shared/eventtype-enumeration.md)
[Méthode ProjectDocument.addHandlerAsync](../../reference/shared/projectdocument.addhandlerasync.md)
[Méthode ProjectDocument.removeHandlerAsync](../../reference/shared/projectdocument.removehandlerasync.md)
[Objet ProjectDocument](../../reference/shared/projectdocument.projectdocument.md)

