# Objet ChartDataLabels (interface API JavaScript pour Excel)

Représente une collection de toutes les étiquettes de données associées à un point de graphique.

## Propriétés

| Propriété     | Type   |Description
|:---------------|:--------|:----------|
|Position|string|Valeur DataLabelPosition qui représente la position de l’étiquette de données. Les valeurs possibles sont les suivantes : None, Center, InsideEnd, InsideBase, OutsideEnd, Left, Right, Top, Bottom, BestFit, Callout. En écriture seule.|
|Separator|string|Chaîne représentant le séparateur utilisé pour les étiquettes de données d’un graphique. En écriture seule.|
|showBubbleSize|bool|Valeur booléenne indiquant si la taille des bulles des étiquettes de données est visible ou non. En écriture seule.|
|showCategoryName|bool|Valeur booléenne indiquant si le nom de catégorie des étiquettes de données est visible ou non. En écriture seule.|
|showLegendKey|bool|Valeur booléenne indiquant si le symbole de légende des étiquettes de données est visible ou non. En écriture seule.|
|showPercentage|bool|Valeur booléenne indiquant si le pourcentage des étiquettes de données est visible ou non. En écriture seule.|
|showSeriesName|bool|Valeur booléenne indiquant si le nom de série des étiquettes de données est visible ou non. En écriture seule.|
|showValue|bool|Valeur booléenne indiquant si la valeur des étiquettes de données est visible ou non. En écriture seule.|

_Voir des [exemples](#exemples) d’accès aux propriétés._

## Relations
| Relation | Type   |Description|
|:---------------|:--------|:----------|
|format|[ChartDataLabelFormat](chartdatalabelformat.md)|Représente le format des étiquettes de données du graphique, à savoir le format de remplissage et de la police. En lecture seule.|

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

Faire apparaître le nom de série dans les étiquettes de données et définir la valeur `position` sur « top »

```js
Excel.run(function (ctx) { 
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1"); 
    chart.datalabels.visible = true;
    chart.datalabels.position = "top";
    chart.datalabels.ShowSeriesName = true;
    return ctx.sync().then(function() {
            console.log("Datalabels Shown");
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
