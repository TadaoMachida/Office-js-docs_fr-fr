# Objet Workbook (interface API JavaScript pour Excel)

Le classeur est l’objet de niveau supérieur qui contient des objets connexes tels que des feuilles de calcul, des tableaux, des plages, etc.

## Propriétés

Aucun

## Relations
| Relation | Type   |Description|
|:---------------|:--------|:----------|
|application|[Application](application.md)|Représente une instance de l’application Excel contenant ce classeur. En lecture seule.|
|bindings|[BindingCollection](bindingcollection.md)|Représente une collection de liaisons appartenant au classeur. En lecture seule.|
|fonctions|[Fonctions](functions.md)|Représente l’instance de l’application Excel contenant ce classeur. En lecture seule.|
|names|[NamedItemCollection](nameditemcollection.md)|Représente une collection d’éléments nommés portant sur le classeur (appelés plages et constantes). En lecture seule.|
|tables|[TableCollection](tablecollection.md)|Représente une collection de tableaux associés au classeur. En lecture seule.|
|Worksheets|[WorksheetCollection](worksheetcollection.md)|Représente une collection de feuilles de calcul associées au classeur. En lecture seule.|

## Méthodes

| Méthode           | Type renvoyé    |Description|
|:---------------|:--------|:----------|
|[getSelectedRange()](#getselectedrange)|[Range](range.md)|Obtient la plage sélectionnée dans le classeur.|
|[load(param: object)](#loadparam-object)|void|Remplit l’objet proxy créé dans le calque JavaScript avec des valeurs de propriété et d’objet spécifiées dans le paramètre.|

## Détails des méthodes


### getSelectedRange()
Obtient la plage sélectionnée dans le classeur.

#### Syntaxe
```js
workbookObject.getSelectedRange();
```

#### Paramètres
Aucun

#### Retourne
[Range](range.md)

#### Exemples

```js
Excel.run(function (ctx) { 
    var selectedRange = ctx.workbook.getSelectedRange();
    selectedRange.load('address');
    return ctx.sync().then(function() {
            console.log(selectedRange.address);
    });
}).catch(function(error) {
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

#### Renvoie
void
