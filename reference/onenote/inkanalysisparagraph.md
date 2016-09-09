# Objet InkAnalysisParagraph (interface API JavaScript pour OneNote)

_S’applique à : OneNote Online_  


Représente les données d’analyse des entrées manuscrites pour un paragraphe identifié formé de traits d’encre.

## Propriétés

| Propriété     | Type   |Description|Commentaires|
|:---------------|:--------|:----------|:-------|
|id|chaîne|Obtient l’ID de l’objet InkAnalysisParagraph. En lecture seule.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisParagraph-id)|

_Voir des [exemples](#exemples) d’accès aux propriétés._

## Relations
| Relation | Type   |Description| Commentaires|
|:---------------|:--------|:----------|:-------|
|inkAnalysis|[InkAnalysis](inkanalysis.md)|Référence à l’objet InkAnalysisPage parent. En lecture seule.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisParagraph-inkAnalysis)|
|lines|[InkAnalysisLineCollection](inkanalysislinecollection.md)|Obtient les lignes d’analyse des entrées manuscrites dans ce paragraphe d’analyse des entrées manuscrites. En lecture seule.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisParagraph-lines)|

## Méthodes

| Méthode           | Type renvoyé    |Description| Commentaires|
|:---------------|:--------|:----------|:-------|
|[load(param: object)](#loadparam-object)|void|Remplit l’objet proxy créé dans le calque JavaScript avec des valeurs de propriété et d’objet spécifiées dans le paramètre.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkAnalysisParagraph-load)|

## Détails des méthodes


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

**lines**
```js
OneNote.run(function (ctx) {        
    var app = ctx.application;
    
    // Gets the active page.
    var page = app.getActivePage();
    
    // Load a line of ink words.
    page.load('inkAnalysisOrNull/paragraphs/lines');
    
    return ctx.sync()
        .then(function() {
            var inkParagraphs = page.inkAnalysisOrNull.paragraphs;
            
            // Log id of each line in ink paragraphs.
            $.each(inkParagraphs.items, function(i, inkParagraph){
                var inkLines = inkParagraph.lines;
                $.each(inkLines.items, function (j, inkLine) {
                    console.log(inkLine.id);
                })
            })
        })
})
.catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
}); 
```