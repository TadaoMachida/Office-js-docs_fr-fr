# Objet ChartFill (interface API JavaScript pour Excel)

_S’applique à : Excel 2016, Excel Online, Office 2016_

Représente le format de remplissage d’un élément de graphique.

## Propriétés

Aucun

## Relations
Aucun


## Méthodes

| Méthode   | Type renvoyé|Description|
|:---------------|:--------|:----------|
|[clear()](#clear)|void|Supprime la couleur de remplissage d’un élément de graphique.|
|[setSolidColor(color: string)](#setsolidcolorcolor-string)|void|Définit le format de remplissage d’un élément de graphique sur une couleur unie.|

## Détails des méthodes

### Effacer
Supprime la couleur de remplissage d’un élément de graphique.

#### Syntaxe
```js
chartFillObject.clear();
```

#### Paramètres
Aucun

#### Retourne
void

#### Exemples

Désactiver le format des lignes de quadrillage principal sur l’axe des ordonnées du graphique « Chart1 »

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
### setSolidColor(color: string)
Définit le format de remplissage d’un élément de graphique sur une couleur unie.

#### Syntaxe
```js
chartFillObject.setSolidColor(color);
```

#### Paramètres
| Paramètre   | Type|Description|
|:---------------|:--------|:----------|
|color|string|Code couleur HTML qui représente la couleur de la ligne de bordure, au format #RRGGBB (par exemple : « FFA500 ») ou sous forme de couleur HTML nommée (par exemple, « orange »).|

#### Retourne
void

#### Exemples

Définir le rouge comme couleur d’arrière-plan de Chart1

```js
Excel.run(function (ctx) { 
	var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	

	chart.format.fill.setSolidColor("#FF0000");

	return ctx.sync().then(function() {
			console.log("Chart1 Background Color Changed.");
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

