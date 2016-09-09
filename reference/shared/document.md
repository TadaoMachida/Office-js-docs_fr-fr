
# Objet Document
Une classe abstraite qui représente le document avec lequel interagit le complément.

|||
|:-----|:-----|
|**Hôtes :**|Access, Excel, PowerPoint, Project, Word|
|**Ajouté dans**|1,0|
|**Dernière modification dans **|1.1|

```
Office.context.document
```


## Membres


**Propriétés**


|**Nom**|**Description**|**Notes de prise en charge**|
|:-----|:-----|:-----|
|[bindings](../../reference/shared/document.bindings.md)|Obtient un objet qui fournit l’accès aux liaisons définies dans le document.|Dans la version 1.1, prise en charge supplémentaire des compléments de contenu pour Access.|
|[customXmlParts](../../reference/shared/document.customxmlparts.md)|Obtient un objet qui représente les parties XML personnalisées contenues dans le document.||
|[mode](../../reference/shared/document.mode.md)|Obtient le mode dans lequel se trouve le document.|Dans la version 1.1, prise en charge supplémentaire des compléments de contenu pour Access.|
|[paramètres](../../reference/shared/document.settings.md)|Obtient un objet qui représente les paramètres personnalisés enregistrés du complément de contenu ou de volet des tâches pour le document actif.|Dans la version 1.1, prise en charge supplémentaire des compléments de contenu pour Access.|
|[url](../../reference/shared/document.url.md)|Obtient l’URL du document actuellement ouvert dans l’application hôte.|Dans la version 1.1, prise en charge supplémentaire des compléments de contenu pour Access.|

**Méthodes**


|**Nom**|**Description**|**Notes de prise en charge**|
|:-----|:-----|:-----|
|[addHandlerAsync](../../reference/shared/document.addhandlerasync.md)|Ajoute un gestionnaire d’événements pour un événement d’objet **Document**.||
|[getActiveViewAsync](../../reference/shared/document.getactiveviewasync.md)|Retourne l’affichage actuel de la présentation.|Dans la version 1.1, prise en charge supplémentaire de [compléments pour PowerPoint](../../docs/powerpoint/powerpoint-add-ins.md).|
|[getFileAsync](../../reference/shared/document.getfileasync.md)|Retourne l’intégralité du fichier de document sous forme de sections pouvant aller jusqu’à 4 194 304 octets (4 Mo).|Dans la version 1.1, prise en charge supplémentaire de l’obtention d’un fichier PDF dans les compléments pour PowerPoint et Word.|
|[getFilePropertiesAsync](../../reference/shared/document.getfilepropertiesasync.md)|Obtient les propriétés de fichier du document actif. Dans cette version, seule l’URL du document peut être obtenue.|Dans la version 1.1, ajout de l’obtention de l’URL du document dans les compléments pour Excel, Word et PowerPoint.|
|[getSelectedDataAsync](../../reference/shared/document.getselecteddataasync.md)|Lit les données contenues dans la sélection actuelle du document.|Dans la version 1.1, prise en charge supplémentaire de l’obtention de l’identifiant, du titre et de l’index pour la plage sélectionnée de diapositives dans les compléments pour PowerPoint.|
|[goToByIdAsync](../../reference/shared/document.gotobyidasync.md)|Accède à l’emplacement ou l’objet spécifié dans le document.|Dans la version 1.1, prise en charge supplémentaire de la navigation dans le document dans les compléments pour Excel et PowerPoint.|
|[removeHandlerAsync](../../reference/shared/document.removehandlerasync.md)|Supprime un gestionnaire d’événements pour un événement d’objet **Document**.||
|[setSelectedDataAsync](../../reference/shared/document.setselecteddataasync.md)|Écrit des données dans la sélection actuelle au sein du document.|Dans la version 1.1, prise en charge supplémentaire de la [définition de la mise en forme du tableau sélectionné lors de l’écriture de données dans les compléments pour Excel](../../docs/excel/format-tables-in-add-ins-for-excel.md).|

**Événements**


|**Nom**|**Description**|**Notes de prise en charge**||
|:-----|:-----|:-----|:-----|
|[ActiveViewChanged](../../reference/shared/document.activeviewchanged.md)|Survient lorsque l’utilisateur modifie l’affichage actuel du document.|Dans la version 1.1, prise en charge supplémentaire de compléments pour PowerPoint.||
|[SelectionChanged](../../reference/shared/document.selectionchanged.event.md)|Se produit quand la sélection change dans le document.|||

## Remarques

N’instanciez pas l’objet **Document** directement dans votre script. Pour appeler des membres de l’objet **Document** afin d’interagir avec le document actif ou la feuille de calcul active, utilisez `Office.context.document` dans votre script.


## Exemple

L’exemple suivant utilise la méthode **getSelectedDataAsync** de l’objet **Document** pour récupérer la sélection actuelle de l’utilisateur sous forme de texte et l’afficher ensuite dans la page du complément.


```js

// Display the user's current selection.
function showSelection() {
    Office.context.document.getSelectedDataAsync(
        "text",                        // coercionType
        {valueFormat: "unformatted",   // valueFormat
        filterType: "all"},            // filterType
        function (result) {            // callback
            var dataValue; 
            dataValue = result.value;
            write('Selected data is: ' + dataValue);
        });
}
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```




## Informations de prise en charge


La prise en charge de chaque membre d’API de l’objet **Document** diffère dans les applications hôtes Office. Voir la section « Informations de prise en charge » de la rubrique de chaque membre pour découvrir les informations de prise en charge d’hôte.

Pour plus d’informations sur les exigences de l’application et du serveur hôtes Office, voir [Configuration requise pour exécuter des compléments pour Office](../../docs/overview/requirements-for-running-office-add-ins.md).


|||
|:-----|:-----|
|**Ajouté dans**|1,0|
|**Dernière modification dans **|1.1|
|**Types de complément**|Application de contenu et de volet de tâches|
|**Bibliothèque**|Office.js|
|**Espace de noms**|Bureau|
