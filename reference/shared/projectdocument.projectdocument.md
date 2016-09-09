

# ProjectDocument, objet
Classe abstraite qui représente le document du projet (projet actif) avec lequel le complément Office interagit.

|||
|:-----|:-----|
|**Hôtes :**|Projet|
|**Ajouté dans**|1,0|

```js
Office.context.document
```


## Membres


**Méthodes**


|**Nom**|**Description**|
|:-----|:-----|
|[addHandlerAsync, méthode](../../reference/shared/projectdocument.addhandlerasync.md)|Ajoute de manière asynchrone un gestionnaire d’événements pour un événement dans un objet **ProjectDocument**.|
|[Méthode getMaxResourceIndexAsync](../../reference/shared/projectdocument.getmaxresourceindexasync.md)|Obtient de façon asynchrone l’index maximal de la collection de ressources dans le projet en cours.|
|[Méthode getMaxTaskIndexAsync](../../reference/shared/projectdocument.getmaxtaskindexasync.md)|Obtenez de façon asynchrone l’index maximal de la collection de tâches dans le projet en cours.|
|[Méthode getProjectFieldAsync](../../reference/shared/projectdocument.getprojectfieldasync.md)|Obtient de manière asynchrone la valeur du champ spécifié dans le projet actif.|
|[Méthode getResourceByIndexAsync](../../reference/shared/projectdocument.getresourcebyindexasync.md)|Obtient de façon asynchrone le GUID de la ressource contenant l’index indiqué dans la collection de ressources.|
|[Méthode getResourceFieldAsync](../../reference/shared/projectdocument.getresourcefieldasync.md)|Obtient de manière asynchrone la valeur du champ spécifié pour la ressource spécifiée.|
|[Méthode getSelectedDataAsync](../../reference/shared/projectdocument.getselecteddataasync.md)|Obtient de manière asynchrone les données contenues dans la sélection actuelle d’une ou de plusieurs cellules du diagramme de Gantt.|
|[Méthode getSelectedResourceAsync](../../reference/shared/projectdocument.getselectedresourceasync.md)|Obtient de manière asynchrone le GUID de la ressource sélectionnée.|
|[Méthode getSelectedTaskAsync](../../reference/shared/projectdocument.getselectedtaskasync.md)|Obtient de manière asynchrone le GUID de la tâche sélectionnée.|
|[Méthode getSelectedViewAsync](../../reference/shared/projectdocument.getselectedviewasync.md)|Obtient de manière asynchrone le type d’affichage et le nom de l’affichage actif.|
|[Méthode getTaskAsync](../../reference/shared/projectdocument.gettaskasync.md)|Obtient de manière asynchrone le nom de la tâche, les ressources affectées à la tâche et l’ID de la tâche dans la liste de tâches SharePoint synchronisées.|
|[Méthode getTaskByIndexAsync](../../reference/shared/projectdocument.gettaskbyindexasync.md)|Obtient de manière asynchrone le GUID de la tâche comportant l’index spécifié dans la collection de tâches.|
|[Méthode getTaskFieldAsync](../../reference/shared/projectdocument.gettaskfieldasync.md)|Obtient de manière asynchrone la valeur du champ spécifié pour la tâche spécifiée.|
|[Méthode getWSSUrlAsync](../../reference/shared/projectdocument.getwssurlasync.md)|Obtient de manière asynchrone l’URL de la liste de tâches SharePoint synchronisées.|
|[Méthode removeHandlerAsync](../../reference/shared/projectdocument.removehandlerasync.md)|Supprime de manière asynchrone un gestionnaire d’événements pour un événement dans un objet **ProjectDocument**.|
|[Méthode setResourceFieldAsync](../../reference/shared/projectdocument.setresourcefieldasync.md)|Définit de manière asynchrone la valeur du champ spécifié pour la ressource spécifiée.|
|[Méthode setTaskFieldAsync](../../reference/shared/projectdocument.settaskfieldasync.md)|Définit de manière asynchrone la valeur du champ spécifié pour la tâche spécifiée.|

**Événements**


|**Nom**|**Description**|
|:-----|:-----|
|[Événement ResourceSelectionChanged](../../reference/shared/projectdocument.resourceselectionchanged.event.md)|Se produit quand la sélection de ressource change dans le projet actif.|
|[Événement TaskSelectionChanged](../../reference/shared/projectdocument.taskselectionchanged.event.md)|Se produit quand la sélection de tâche change dans le projet actif.|
|[Événement ViewSelectionChanged](../../reference/shared/projectdocument.viewselectionchanged.event.md)|Se produit quand l’affichage actif change dans le projet actif.|

## Remarques

N’appelez pas ou n’instanciez pas directement l’objet **ProjectDocument** dans votre script.


## Exemple

L’exemple suivant initialise le complément et obtient les propriétés de l’objet [Document](../../reference/shared/document.md) qui sont disponibles dans le contexte d’un document Project. Un document Project est le projet ouvert et actif. Pour accéder aux membres de l’objet **ProjectDocument**, utilisez l’objet **Office.context.document** comme le montrent les exemples de code pour les méthodes et les événements **ProjectDocument**.

L’exemple suppose que votre complément comporte une référence à la bibliothèque jQuery et que le contrôle de page suivant est défini dans la balise div de contenu du corps de la page :




```HTML
<span id="message"></span>
```




```js
(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {

            // Get information about the document.
            showDocumentProperties();
        });
    };

    // Get the document mode and the URL of the active project.
    function showDocumentProperties() {
        var output = String.format(
            'The document mode is {0}.<br/>The URL of the active project is {1}.',
            Office.context.document.mode,
            Office.context.document.url);
        $('#message').html(output);
    }
})();
```


## Informations de prise en charge


Un Y majuscule dans la matrice suivante indique que cet objet est pris en charge dans l’application hôte Office correspondante. Une cellule vide indique que l’application hôte Office ne prend pas en charge cet objet.

Pour plus d’informations sur les exigences de l’application et du serveur hôtes Office, voir [Configuration requise pour exécuter des compléments pour Office](../../docs/overview/requirements-for-running-office-add-ins.md).


||**Office pour Bureau Windows**|**Office Online (dans un navigateur)**|
|:-----|:-----|:-----|
|**Projet**|v||

|||
|:-----|:-----|
|**Types de complément**|Volet de tâches|
|**Bibliothèque**|Office.js|
|**Espace de noms**|Bureau|

## Historique de prise en charge


|**Version**|**Modifications**|
|:-----|:-----|
|1,0|Introduit|

## Voir aussi



#### Autres ressources


[Compléments du volet Office pour Project](../../docs/project/project-add-ins.md)
[Objet Document](../../reference/shared/document.md)

