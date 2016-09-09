# Objet ChartAxis (interface API JavaScript pour Excel)

Représente un axe unique dans un graphique.

## Propriétés

| Propriété     | Type   |Description
|:---------------|:--------|:----------|
|majorUnit|object|Représente l’intervalle entre deux graduations principales. Peut être défini sur une valeur numérique ou une chaîne vide. La valeur renvoyée est toujours un nombre.|
|maximum|object|Représente la valeur maximale pour l’axe des ordonnées. Peut être défini sur une valeur numérique ou une chaîne vide (pour les valeurs d’axe automatique). La valeur renvoyée est toujours un nombre.|
|minimum|object|Représente la valeur minimale pour l’axe des ordonnées. Peut être défini sur une valeur numérique ou une chaîne vide (pour les valeurs d’axe automatique). La valeur renvoyée est toujours un nombre.|
|minorUnit|object|Représente l’intervalle entre deux graduations secondaires. Peut être défini sur une valeur numérique ou une chaîne vide (pour les valeurs d’axe automatique). La valeur renvoyée est toujours un nombre.|

_Voir des [exemples](#exemples) d’accès aux propriétés._

## Relations
| Relation | Type   |Description|
|:---------------|:--------|:----------|
|format|[ChartAxisFormat](chartaxisformat.md)|Représente la mise en forme d’un objet de graphique, à savoir le format des lignes et de la police. En lecture seule.|
|majorGridlines|[ChartGridlines](chartgridlines.md)|Renvoie un objet de quadrillage qui représente les lignes de quadrillage principal de l’axe spécifié. En lecture seule.|
|minorGridlines|[ChartGridlines](chartgridlines.md)|Renvoie un objet de quadrillage qui représente les lignes de quadrillage secondaire de l’axe spécifié. En lecture seule.|
|Fonction|[ChartAxisTitle](chartaxistitle.md)|Représente le titre de l’axe. En lecture seule.|

## Méthodes

| Méthode           | Type renvoyé    |Description|
|:---------------|:--------|:----------|
|[load(param: object)](#loadparam-object)|void|Remplit l’objet proxy créé dans le calque JavaScript avec des valeurs de propriété et d’objet spécifiées dans le paramètre.|

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

#### Renvoie
void
### Exemples d’accès aux propriétés
Obtenir la valeur `maximum` de l’axe du graphique Chart1

```js
Excel.run(function (ctx) { 
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1"); 
    var axis = chart.axes.valueaxis;
    axis.load('maximum');
    return ctx.sync().then(function() {
            console.log(axis.maximum);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

Définir la valeur `maximum`, `minimum`, `majorunit` ou `minorunit` de l’axe des ordonnées. 

```js
Excel.run(function (ctx) { 
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1"); 
    chart.axes.valueaxis.maximum = 5;
    chart.axes.valueaxis.minimum = 0;
    chart.axes.valueaxis.majorunit = 1;
    chart.axes.valueaxis.minorunit = 0.2;
    return ctx.sync().then(function() {
            console.log("Axis Settings Changed");
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
