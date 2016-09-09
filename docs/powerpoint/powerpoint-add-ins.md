
# Création de compléments de contenu et du volet Office pour PowerPoint

Les exemples de code dans l’article illustrent quelques tâches élémentaires de développement de compléments de contenu PowerPoint. Les informations affichées dans ces exemples dépendent de la fonction  `app.showNotification`, qui figure dans les modèles de projet Visual Studio d’Compléments Office. Si vous n’utilisez pas Visual Studio pour développer votre complément, vous devez remplacer la fonction  `showNotification` par votre propre code. Un certain nombre de ces exemples dépendent également de cet objet `globals` qui se trouve en dehors de la portée des fonctions suivantes : `var globals = {activeViewHandler:0, firstSlideId:0};`

Pour obtenir ces exemples de code, votre projet doit faire référence à la [bibliothèque Office.js v1.1 ou version ultérieure](../../docs/develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md).


## Détecter l’affichage actif de la présentation et gérer l’événement ActiveViewChanged

La fonction  `getFileView` appelle la méthode [Document.getActiveViewAsync](../../reference/shared/document.getactiveviewasync.md) afin de renvoyer si la vue actuelle de la présentation est une vue de « modification » (toutes les vues dans lesquelles vous modifiez des diapositives, telles que les vues **Normal** ou **Mode Plan**) ou « lecture » ( **Diaporama** ou **Mode Lecture**).


```js
function getFileView() {
    //Gets whether the current view is edit or read.
    Office.context.document.getActiveViewAsync(function (asyncResult) {
        if (asyncResult.status == "failed") {
            app.showNotification("Action failed with error: " + asyncResult.error.message);
        }
        else {
            app.showNotification(asyncResult.value);
        }
    });
}
```

La fonction  `registerActiveViewChanged` appelle la méthode [addHandlerAsync](../../reference/shared/document.addhandlerasync.md) pour inscrire un gestionnaire pour l’événement [Document.ActiveViewChanged](../../reference/shared/document.activeviewchanged.md). Après l’exécution de la fonction, lorsque vous modifiez la vue de la présentation, la notification  `app.showNotification` affiche le mode d’affichage actif (« lecture » ou « modification »).




```js
function registerActiveViewChanged() {
    Globals.activeViewHandler = function (args) {
        app.showNotification(JSON.stringify(args));
    }

    Office.context.document.addHandlerAsync(Office.EventType.ActiveViewChanged, Globals.activeViewHandler, 
        function (asyncResult) {
            if (asyncResult.status == "failed") {
            app.showNotification("Action failed with error: " + asyncResult.error.message);
        }
            else {
            app.showNotification(asyncResult.status);
            }
        });
}
```


## Obtenir l’URL de la présentation

La fonction `getFileUrl` appelle la méthode [Document.getFileProperties](../../reference/shared/document.getfilepropertiesasync.md) pour obtenir l’URL du fichier de présentation.


```js
function getFileUrl() {
    //Get the URL of the current file.
    Office.context.document.getFilePropertiesAsync(function (asyncResult) {
        var fileUrl = asyncResult.value.url;
        if (fileUrl == "") {
            app.showNotification("The file hasn't been saved yet. Save the file and try again");
        }
        else {
            app.showNotification(fileUrl);
        }
    });
}
```


## Accéder à une diapositive spécifique dans la présentation

La fonction  `getSelectedRange` appelle la méthode [Document.getSelectedDataAsync](../../reference/shared/document.getselecteddataasync.md) pour obtenir un objet JSON renvoyé par `asyncResult.value` et qui contient un tableau intitulé « diapositives » répertoriant les ID, les titres et les index de la série de diapositives sélectionnée (ou uniquement de la diapositive en cours). Elle enregistre également l’ID de la première diapositive de la série sélectionnée dans une variable globale.


```js
function getSelectedRange() {
    // Get the id, title, and index of the current slide (or selected slides) and store the first slide id */
    Globals.firstSlideId = 0;

    Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange, function (asyncResult) {
        if (asyncResult.status == "failed") {
            app.showNotification("Action failed with error: " + asyncResult.error.message);
        }
        else {
            Globals.firstSlideId = asyncResult.value.slides[0].id;
            app.showNotification(JSON.stringify(asyncResult.value));
        }
    });
}
```

La fonction  `goToFirstSlide` appelle la méthode [Document.goToByIdAsync](../../reference/shared/document.gotobyidasync.md) pour accéder à l’ID de la première diapositive stockée par la fonction `getSelectedRange` ci-dessus.




```js
function goToFirstSlide() {
    Office.context.document.goToByIdAsync(Globals.firstSlideId, Office.GoToType.Slide, function (asyncResult) {
        if (asyncResult.status == "failed") {
            app.showNotification("Action failed with error: " + asyncResult.error.message);
        }
        else {
            app.showNotification("Navigation successful");
        }
    });
}
```


## Naviguer entre les diapositives de la présentation

La fonction  `goToSlideByIndex` appelle la méthode **Document.goToByIdAsync** pour passer à la diapositive suivante dans la présentation.


```js
function goToSlideByIndex() {
    var goToFirst = Office.Index.First;
    var goToLast = Office.Index.Last;
    var goToPrevious = Office.Index.Previous;
    var goToNext = Office.Index.Next;

    Office.context.document.goToByIdAsync(goToNext, Office.GoToType.Index, function (asyncResult) {
        if (asyncResult.status == "failed") {
            app.showNotification("Action failed with error: " + asyncResult.error.message);
        }
        else {
            app.showNotification("Navigation successful");
        }
    });
}
```




## Ressources supplémentaires

- [Enregistrement de l’état et des paramètres d’un complément par document pour les compléments de contenu et du volet Office](../../docs/develop/persisting-add-in-state-and-settings.md#how-to-save-add-in-state-and-settings-per-document-for-content-and-task-pane-add-ins)

- [Lire et écrire des données dans la sélection active dans un document ou une feuille de calcul](../../docs/develop/read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md)
    
- [Obtention de l’intégralité d’un document pour un complément pour PowerPoint ou Word](../../docs/develop/get-the-whole-document-from-an-add-in-for-powerpoint-or-word.md)
    
- [Utiliser des thèmes de document dans vos compléments PowerPoint](../powerpoint/use-document-themes-in-your-powerpoint-add-ins.md)
    
