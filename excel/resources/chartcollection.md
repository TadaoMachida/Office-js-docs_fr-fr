# Objet ChartCollection (interface API JavaScript pour Excel)

_S’applique à : Excel 2016, Excel Online, Office 2016_

Collection de tous les objets de graphique d’une feuille de calcul.

## Propriétés

| Propriété   | Type|Description
|:---------------|:--------|:----------|
|count|int|Renvoie le nombre de graphiques dans la feuille de calcul. En lecture seule.|
|Items|[Chart[]](chart.md)|Collection d’objets de graphique. En lecture seule.|

_Voir des [exemples](#property-access-examples) d’accès aux propriétés._

## Relations
Aucun


## Méthodes

| Méthode   | Type renvoyé|Description|
|:---------------|:--------|:----------|
|[add(type: string, sourceData: Range, seriesBy: string)](#addtype-string-sourcedata-range-seriesby-string)|[Chart](chart.md)|Crée un graphique.|
|[getItem(name: string)](#getitemname-string)|[Chart](chart.md)|Extrait un graphique à l’aide de son nom. Si plusieurs graphiques portent le même nom, c’est le premier d’entre eux qui est renvoyé.|
|[getItemAt(index: number)](#getitematindex-number)|[Chart](chart.md)|Extrait un graphique en fonction de sa position dans la collection.|
|[load(param: object)](#loadparam-object)|void|Remplit l’objet proxy créé dans le calque JavaScript avec des valeurs de propriété et d’objet spécifiées dans le paramètre.|

## Détails des méthodes

### add(type: string, sourceData: Range, seriesBy: string)
Crée un graphique.

#### Syntaxe
```js
chartCollectionObject.add(type, sourceData, seriesBy);
```

#### Paramètres
| Paramètre   | Type|Description|
|:---------------|:--------|:----------|
|type|string|Représente le type d’un graphique. Les valeurs possibles sont les suivantes : ColumnClustered, ColumnStacked, ColumnStacked100, BarClustered, BarStacked, BarStacked100, LineStacked, LineStacked100, LineMarkers, LineMarkersStacked, LineMarkersStacked100, PieOfPie, etc.|
|sourceData|Range|Plage qui contient les données sources.|
|seriesBy|string|Facultatif. Spécifie la façon dont les colonnes ou les lignes sont utilisées comme séries de données sur le graphique.  Les valeurs possibles sont les suivantes : Auto (automatique), Columns (colonnes), Rows (lignes)|

#### Retourne
[Chart](chart.md)

#### Exemples

Ajouter un graphique dont la valeur `chartType` est « ColumnClustered » dans la feuille de calcul « Charts », avec la propriété `sourceData` définie sur la plage « A1:B4 » et la propriété `seriesBy` définie sur « auto »

```js
Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var sourceData = sheetName + "!" + "A1:B4";
	var chart = ctx.workbook.worksheets.getItem(sheetName).charts.add("ColumnClustered", sourceData, "auto");
	return ctx.sync().then(function() {
			console.log("New Chart Added");
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

### getItem(name: string)
Extrait un graphique à l’aide de son nom. Si plusieurs graphiques portent le même nom, c’est le premier d’entre eux qui est renvoyé.

#### Syntaxe
```js
chartCollectionObject.getItem(name);
```

#### Paramètres
| Paramètre   | Type|Description|
|:---------------|:--------|:----------|
|name|string|Nom du graphique à extraire.|

#### Retourne
[Chart](chart.md)

#### Exemples

```js
Excel.run(function (ctx) { 
	var chartname = 'Chart1';
	var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem(chartname);
	return ctx.sync().then(function() {
			console.log(chart.height);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```


#### Exemples

```js
Excel.run(function (ctx) { 
	var chartId = 'SamplChartId';
	var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem(chartId);
	return ctx.sync().then(function() {
			console.log(chart.height);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```



#### Exemples

```js
Excel.run(function (ctx) { 
	var lastPosition = ctx.workbook.worksheets.getItem("Sheet1").charts.count - 1;
	var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItemAt(lastPosition);
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

### getItemAt(index: number)
Extrait un graphique en fonction de sa position dans la collection.

#### Syntaxe
```js
chartCollectionObject.getItemAt(index);
```

#### Paramètres
| Paramètre   | Type|Description|
|:---------------|:--------|:----------|
|index|number|Valeur d’indice de l’objet à récupérer. Avec indice zéro.|

#### Retourne
[Chart](chart.md)

#### Exemples

```js
Excel.run(function (ctx) { 
	var lastPosition = ctx.workbook.worksheets.getItem("Sheet1").charts.count - 1;
	var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItemAt(lastPosition);
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

### load(param: object)
Remplit l’objet proxy créé dans le calque JavaScript avec des valeurs de propriété et d’objet spécifiées dans le paramètre.

#### Syntaxe
```js
object.load(param);
```

#### Paramètres
| Paramètre   | Type|Description|
|:---------------|:--------|:----------|
|param|object|Facultatif. Accepte les noms de paramètre et de relation sous forme de chaîne délimitée ou de tableau. Sinon, indiquez l’objet [loadOption](loadoption.md).|

#### Renvoie
void
### Exemples d’accès aux propriétés

```js
Excel.run(function (ctx) { 
	var charts = ctx.workbook.worksheets.getItem("Sheet1").charts;
	charts.load('items');
	return ctx.sync().then(function() {
		for (var i = 0; i < charts.items.length; i++)
		{
			console.log(charts.items[i].name);
			console.log(charts.items[i].index);
		}
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

Obtenir le nombre de graphiques

```js
Excel.run(function (ctx) { 
	var charts = ctx.workbook.worksheets.getItem("Sheet1").charts;
	charts.load('count');
	return ctx.sync().then(function() {
		console.log("charts: Count= " + charts.count);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```


