# Objet RangeFormat (interface API JavaScript pour Excel)

_S’applique à : Excel 2016, Excel Online, Office 2016_

Objet de format qui comprend les propriétés de police, de remplissage, de bordures, d’alignement, etc. de la plage.

## Propriétés

| Propriété   | Type|Description
|:---------------|:--------|:----------|
|horizontalAlignment|string|Représente l’alignement horizontal de l’objet spécifié. Les valeurs possibles sont les suivantes : General (général), Left (gauche), Center (centré), Right (droit), Fill (remplir), Justify (justifié), CenterAcrossSelection (centré pour toute la sélection), Distributed (distribué).|
|verticalAlignment|string|Représente l’alignement vertical de l’objet spécifié. Les valeurs possibles sont les suivantes : Top (haut), Center (centré), Bottom (bas), Justify (justifié), Distributed (distribué).|
|wrapText|bool|Indique que le contrôle de texte Excel est défini pour renvoyer à la ligne automatiquement le texte dans l’objet. Une valeur null indique que les paramètres de retour à la ligne ne sont pas les mêmes sur l’ensemble de la plage.|

_Voir des [exemples](#property-access-examples) d’accès aux propriétés._

## Relations
| Relation | Type|Description|
|:---------------|:--------|:----------|
|Borders|[RangeBorderCollection](rangebordercollection.md)|Collection d’objets de bordure qui s’appliquent à l’ensemble de la plage sélectionnée. En lecture seule.|
|fill|[RangeFill](rangefill.md)|Renvoie l’objet de remplissage défini sur la plage globale. En lecture seule.|
|font|[RangeFont](rangefont.md)|Renvoie l’objet de police défini sur la plage globale sélectionnée. En lecture seule.|

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

Cet exemple affiche toutes les propriétés de format d’une plage. 

```js
Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "F:G";
	var worksheet = ctx.workbook.worksheets.getItem(sheetName);
	var range = worksheet.getRange(rangeAddress);
	range.load(["format/*", "format/fill", "format/borders", "format/font"]);
	return ctx.sync().then(function() {
		console.log(range.format.wrapText);
		console.log(range.format.fill.color);
		console.log(range.format.font.name);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

L’exemple ci-dessous définit le nom de police et la couleur de remplissage d’une plage et renvoie automatiquement le texte à la ligne. 

```js
Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "F:G";
	var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
	range.format.wrapText = true;
	range.format.font.name = 'Times New Roman';
	range.format.fill.color = '0000FF';
	return ctx.sync(); 
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

L’exemple suivant ajoute une bordure de grille autour de la plage.

```js
Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "F:G";
	var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
	range.format.borders('InsideHorizontal').lineStyle = 'Continuous';
	range.format.borders('InsideVertical').lineStyle = 'Continuous';
	range.format.borders('EdgeBottom').lineStyle = 'Continuous';
	range.format.borders('EdgeLeft').lineStyle = 'Continuous';
	range.format.borders('EdgeRight').lineStyle = 'Continuous';
	range.format.borders('EdgeTop').lineStyle = 'Continuous';
	return ctx.sync(); 
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```
