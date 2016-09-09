# Objet Chart (interface API JavaScript pour Excel)

Représente un objet de graphique dans un classeur.

## Propriétés

| Propriété     | Type   |Description
|:---------------|:--------|:----------|
|height|Double|Représente la hauteur, exprimée en points, de l’objet de graphique.|
|id|string|Extrait un graphique en fonction de sa position dans la collection. En lecture seule.|
|left|Double|Distance, en points, entre le côté gauche du graphique et l’origine de la feuille de calcul.|
|name|string|Représente le nom d’un objet de graphique.|
|top|Double|Représente la distance, en points, entre le bord supérieur de l’objet et la partie supérieure de la ligne 1 (sur une feuille de calcul) ou le haut de la zone de graphique (sur un graphique).|
|width|Double|Représente la largeur, en points, de l’objet de graphique.|

_Voir des [exemples](#exemples) d’accès aux propriétés._

## Relations
| Relation | Type   |Description|
|:---------------|:--------|:----------|
|axes|[ChartAxes](chartaxes.md)|Représente les axes du graphique. En lecture seule.|
|dataLabels|[ChartDataLabels](chartdatalabels.md)|Représente les étiquettes des données sur le graphique. En lecture seule.|
|format|[ChartAreaFormat](chartareaformat.md)|Regroupe les propriétés de format de la zone de graphique. En lecture seule.|
|legend|[ChartLegend](chartlegend.md)|Représente la légende du graphique. En lecture seule.|
|Série|[ChartSeriesCollection](chartseriescollection.md)|Représente une série ou une collection de séries dans le graphique. En lecture seule.|
|Fonction|[ChartTitle](charttitle.md)|Représente le titre du graphique indiqué et comprend le texte, la visibilité, la position et la mise en forme du titre. En lecture seule.|

## Méthodes

| Méthode           | Type renvoyé    |Description|
|:---------------|:--------|:----------|
|[delete()](#delete)|void|Supprime l’objet de graphique.|
|[getImage(height: number, width: number, fittingMode: string)](#getimageheight-number-width-number-fittingmode-string)|System.IO.Stream|Affiche le graphique sous forme d’image codée en Base64 ajusté aux dimensions spécifiées.|
|[load(param: object)](#loadparam-object)|void|Remplit l’objet proxy créé dans le calque JavaScript avec des valeurs de propriété et d’objet spécifiées dans le paramètre.|
|[setData(sourceData: Range, seriesBy: string)](#setdatasourcedata-range-seriesby-string)|void|Redéfinit les données sources du graphique.|
|[setPosition(startCell: Range or string, endCell: Range or string)](#setpositionstartcell-range-or-string-endcell-range-or-string)|void|Positionne le graphique par rapport aux cellules dans la feuille de calcul.|

## Détails des méthodes


### delete()
Supprime l’objet de graphique.

#### Syntaxe
```js
chartObject.delete();
```

#### Paramètres
Aucun

#### Retourne
void

#### Exemples
```js
Excel.run(function (ctx) { 
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1"); 
    chart.delete();
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### getImage(height: number, width: number, fittingMode: string)
Affiche le graphique sous forme d’image codée en Base64 ajusté aux dimensions spécifiées.

#### Syntaxe
```js
chartObject.getImage(height, width, fittingMode);
```

#### Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|height|number|Facultatif. (Facultatif) Hauteur souhaitée de l’image produite.|
|width|number|Facultatif. (Facultatif) Largeur souhaitée de l’image produite.|
|fittingMode|string|Facultatif. (Facultatif) Méthode utilisée pour mettre à l’échelle le graphique aux dimensions spécifiées (si la hauteur et la largeur sont définies).  Les valeurs possibles sont les suivantes : Fit (ajuster), FitAndCenter (ajuster et centrer), Fill (remplir)|

#### Retourne
System.IO.Stream

#### Exemples
```js
Excel.run(function (ctx) { 
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1"); 
    var image = chart.getImage();
    return ctx.sync(); 
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

#### Retourne
void

### setData(sourceData: Range, seriesBy: string)
Redéfinit les données sources du graphique.

#### Syntaxe
```js
chartObject.setData(sourceData, seriesBy);
```

#### Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|sourceData|Range|Objet Range correspondant aux données source.|
|seriesBy|string|Facultatif. Spécifie la façon dont les colonnes ou les lignes sont utilisées comme séries de données sur le graphique. Les valeurs possibles sont les suivantes : Auto (automatique), Columns (colonnes), Rows (lignes). Dans la version bureau, l’option « auto » inspecte la forme des données source pour déterminer automatiquement si les données sont présentées en lignes ou en colonnes. Dans Excel Online, « auto » est défini par défaut sur « columns » (colonnes).|

#### Retourne
void

#### Exemples

Définissez `sourceData` sur « A1: B4 » et `seriesBy` sur « Columns »

```js
Excel.run(function (ctx) { 
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1"); 
    var sourceData = "A1:B4";
    chart.setData(sourceData, "Columns");
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### setPosition(startCell: Range or string, endCell: Range or string)
Positionne le graphique par rapport aux cellules dans la feuille de calcul.

#### Syntaxe
```js
chartObject.setPosition(startCell, endCell);
```

#### Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|startCell|Range or string|Cellule de début. Il s’agit de l’emplacement où le graphique sera déplacé. La cellule de début est la cellule supérieure gauche ou supérieure droite, selon les paramètres d’affichage gauche-droite définis par l’utilisateur.|
|endCell|Range or string|Facultatif. Cellule de fin. Si une valeur est indiquée, la largeur et la hauteur du graphique sont définies de manière à couvrir entièrement cette cellule/plage.|

#### Renvoie
void

#### Exemples


```js
Excel.run(function (ctx) { 
    var sheetName = "Charts";
    var sourceData = sheetName + "!" + "A1:B4";
    var chart = ctx.workbook.worksheets.getItem(sheetName).charts.add("pie", sourceData, "auto");
    chart.width = 500;
    chart.height = 300;
    chart.setPosition("C2", null);
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### Exemples d’accès aux propriétés

Obtenir un graphique nommé « Chart1 »

```js
Excel.run(function (ctx) { 
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1"); 
    chart.load('name');
    return ctx.sync().then(function() {
            console.log(chart.name);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

Mettre à jour un graphique, y compris son nom, sa position et ses dimensions

```js
Excel.run(function (ctx) { 
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1"); 
    chart.name="New Name";
    chart.top = 100;
    chart.left = 100;
    chart.height = 200;
    chart.weight = 200;
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

Donner un nouveau nom au graphique ; définir les dimensions du graphique sur 200 points en hauteur et en largeur. Déplacer Chart1 de 100 points vers le haut et vers la gauche. 

```js
Excel.run(function (ctx) { 
    var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");
    chart.name="New Name";  
    chart.top = 100;
    chart.left = 100;
    chart.height =200;
    chart.width =200;
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

