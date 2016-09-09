# Objet ChartLineFormat (interface API JavaScript pour Excel)

Regroupe les options de mise en forme pour les éléments de ligne.

## Propriétés

| Propriété     | Type   |Description
|:---------------|:--------|:----------|
|color|string|Code couleur HTML qui représente la couleur des lignes dans le graphique.|

_Voir des [exemples](#exemples) d’accès aux propriétés._

## Relations
Aucun


## Méthodes

| Méthode           | Type renvoyé    |Description|
|:---------------|:--------|:----------|
|[clear()](#clear)|void|Désactiver le format de ligne d’un élément de graphique.|
|[load(param: object)](#loadparam-object)|void|Remplit l’objet proxy créé dans le calque JavaScript avec des valeurs de propriété et d’objet spécifiées dans le paramètre.|

## Détails des méthodes


### Effacer
Désactiver le format de ligne d’un élément de graphique.

#### Syntaxe
```js
chartLineFormatObject.clear();
```

#### Paramètres
Aucun

#### Retourne
void

#### Exemples

Désactiver le format des lignes de quadrillage principal pour l’axe des ordonnées du graphique « Chart1 »

```js
Excel.run(function (ctx) { 
    var gridlines = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").axes.valueaxis.majorGridlines;   
    gridlines.format.line.clear();
    return ctx.sync().then(function() {
            console.log("Chart Major Gridlines Format Cleared");
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
### Exemples d’accès aux propriétés

Définir le rouge comme couleur des lignes de quadrillage principal pour l’axe des ordonnées

```js
Excel.run(function (ctx) { 
    var gridlines = ctx.workbook.worksheets.getItem("Sheet1").charts.axes.valueaxis.majorGridlines;
    gridlines.format.line.color = "#FF0000";
    return ctx.sync().then(function() {
            console.log("Chart Gridlines Color Updated");
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
