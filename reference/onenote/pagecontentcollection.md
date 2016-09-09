# Objet PageContentCollection (interface API JavaScript pour OneNote)

_S’applique à : OneNote Online_  


Représente le contenu d’une page, sous la forme d’une collection d’objets PageContent.

## Propriétés

| Propriété     | Type   |Description|Commentaires|
|:---------------|:--------|:----------|:-------|
|count|int|Renvoie le nombre de contenus de page de la collection. En lecture seule.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContentCollection-count)|
|items|[PageContent[]](pagecontent.md)|Collection d’objets PageContent. En lecture seule.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContentCollection-items)|

_Voir des [exemples](#exemples) d’accès aux propriétés._

## Relations
Aucun


## Méthodes

| Méthode           | Type renvoyé    |Description| Commentaires|
|:---------------|:--------|:----------|:-------|
|[getItem(index: number or string)](#getitemindex-number-or-string)|[PageContent](pagecontent.md)|Obtient un objet PageContent en fonction de son ID ou de son index dans la collection. En lecture seule.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContentCollection-getItem)|
|[getItemAt(index: number)](#getitematindex-number)|[PageContent](pagecontent.md)|Obtient un contenu de page en fonction de sa position dans la collection.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContentCollection-getItemAt)|
|[load(param: object)](#loadparam-object)|void|Remplit l’objet proxy créé dans le calque JavaScript avec des valeurs de propriété et d’objet spécifiées dans le paramètre.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-pageContentCollection-load)|

## Détails des méthodes


### getItem(index: number or string)
Obtient un objet PageContent en fonction de son ID ou de son index dans la collection. En lecture seule.

#### Syntaxe
```js
pageContentCollectionObject.getItem(index);
```

#### Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|index|number or string|ID ou emplacement d’index de l’objet PageContent dans la collection.|

#### Retourne
[PageContent](pagecontent.md)

### getItemAt(index: number)
Obtient un contenu de page en fonction de sa position dans la collection.

#### Syntaxe
```js
pageContentCollectionObject.getItemAt(index);
```

#### Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|index|number|Valeur d’indice de l’objet à récupérer. Avec indice zéro.|

#### Retourne
[PageContent](pagecontent.md)

#### Exemples
```js
OneNote.run(function (context) {

    var page = context.application.getActivePage();
    var pageContents = page.contents;
    var firstPageContent = pageContents.getItemAt(0);
    firstPageContent.load('type');

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {
            console.log("The first page content item is of type: " + firstPageContent.type);
            return context.sync();
        });
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

### load(param: object)
Remplit l’objet proxy créé dans le calque JavaScript avec des valeurs de propriété et d’objet spécifiées dans le paramètre.

#### Syntaxe
```js
object.load(param);
```

#### Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|param|object|Facultatif. Accepte les noms de paramètre et de relation sous forme de chaîne délimitée ou de tableau. Sinon, indiquez l’objet [loadOption](loadoption.md).|

#### Retourne
void
### Exemples d’accès aux propriétés

**items**
```js
OneNote.run(function (context) {

    // Get the collection of pageContent items from the page.
    var pageContents = context.application.getActivePage().contents;

    // Queue a command to load the type of each pageContent.
    pageContents.load("type");

    // Run the queued commands, and return a promise to indicate task completion.
    return context.sync()
        .then(function () {

            $.each(pageContents.items, function(index, pageContent) {
                console.log("PageContent type: " + pageContent.type);
            });
        });
})                
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

**traverse for outlines**
```js
OneNote.run(function (context) {
   var page = context.application.getActivePage();
   var pageContents = page.contents;
   pageContents.load('type');
   var outlines = [];
   return context.sync()
       .then(function () {    
              $.each(pageContents.items, function (index, pageContent) {
                     console.log(pageContent.type);
                     if (pageContent.type === 'Outline') {
                           outlines.push(pageContent);
                     }
              });
              $.each(outlines, function (index, outline) {
                     outline.load("id,paragraphs,paragraphs/type");
              });
              return context.sync();
       })
       .then(function () {
              $.each(outlines, function (index, outline) {
                     console.log("An outline was found with id : " + outline.id);
              });
              return Promise.resolve(outlines);
       });
});
```

