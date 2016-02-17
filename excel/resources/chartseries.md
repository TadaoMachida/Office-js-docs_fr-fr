# Objet ChartSeries (interface API JavaScript pour Excel)

_S’applique à : Excel 2016, Excel Online, Office 2016_

Cet objet représente une série dans un graphique.

## Propriétés

| Propriété   | Type|Description
|:---------------|:--------|:----------|
|name|string|Représente le nom d’une série dans un graphique.|

_Voir des [exemples](#property-access-examples) d’accès aux propriétés._

## Relations
| Relation | Type|Description|
|:---------------|:--------|:----------|
|format|[ChartSeriesFormat](chartseriesformat.md)|Représente le format d’une série de graphique, à savoir le format de remplissage et des lignes. En lecture seule.|
|Points|[ChartPointsCollection](chartpointscollection.md)|Représente la collection de tous les points de la série. En lecture seule.|

## Méthodes

| Méthode   | Type renvoyé|Description|
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
| Paramètre   | Type|Description|
|:---------------|:--------|:----------|
|param|object|Facultatif. Accepte les noms de paramètre et de relation sous forme de chaîne délimitée ou de tableau. Sinon, indiquez l’objet [loadOption](loadoption.md).|

#### Renvoie
void
### Exemples d’accès aux propriétés

Renommer la première série de Chart1 sur « New Series Name »

```js
Excel.run(function (ctx) { 
	var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	
	chart.series.getItemAt(0).name = "New Series Name";
	return ctx.sync().then(function() {
			console.log("Series1 Renamed");
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

