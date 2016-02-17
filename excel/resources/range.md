# Objet Range (interface API JavaScript pour Excel)

_S’applique à : Excel 2016, Excel Online, Office 2016_

Une plage représente un ensemble constitué d’une ou de plusieurs cellules contiguës comme une cellule, une ligne, une colonne, un bloc de cellules, etc.

## Propriétés

| Propriété   | Type|Description
|:---------------|:--------|:----------|
|address|string|Représente la référence de plage dans le style A1. La valeur d’adresse contient la référence de feuille (par exemple, Feuille1! A1:B4). En lecture seule.|
|addressLocal|string|Représente la référence de la plage spécifiée dans le langage de l’utilisateur. En lecture seule.|
|cellCount|int|Nombre de cellules dans la plage. En lecture seule.|
|columnCount|int|Représente le nombre total de colonnes dans la plage. En lecture seule.|
|columnIndex|int|Représente le numéro de colonne de la première cellule de la plage. Avec indice zéro. En lecture seule.|
|formulas|object[][]|Représente la formule dans le style de notation A1.|
|formulasLocal|object[][]|Représente la formule en notation A1, en utilisant le langage et les paramètres de format de nombre régionaux de l’utilisateur. Par exemple, la formule « =SUM(A1, 1.5) » en anglais deviendrait « =SUMME(A1; 1,5) » en allemand.|
|numberFormat|object[][]|Représente le code de format de nombre pour une cellule donnée.|
|rowCount|int|Renvoie le nombre total de lignes de la plage. En lecture seule.|
|rowIndex|int|Renvoie le numéro de ligne de la première cellule de la plage. Avec indice zéro. En lecture seule.|
|text|object[][]|Valeurs de texte de la plage spécifiée. La valeur de texte ne dépend pas de la largeur de la cellule. Le remplacement par le signe # qui se produit dans l’interface utilisateur d’Excel n’a aucun effet sur la valeur de texte renvoyée par l’API. En lecture seule.|
|valueTypes|string|Représente le type de données de chaque cellule. En lecture seule. Les valeurs possibles sont les suivantes : Unknown (inconnu), Empty (vide), String (chaîne), Integer (entier), Double (double), Boolean (valeur booléenne), Error (erreur).|
|values|object[][]|Représente les valeurs brutes de la plage spécifiée. Les données renvoyées peuvent être des chaînes, des valeurs numériques ou des valeurs booléennes. Une cellule contenant une erreur renvoie une chaîne d’erreur.|

_Voir des [exemples](#property-access-examples) d’accès aux propriétés._

## Relations
| Relation | Type|Description|
|:---------------|:--------|:----------|
|format|[RangeFormat](rangeformat.md)|Renvoie un objet de format, qui comprend les propriétés de police, de remplissage, de bordures, d’alignement, etc. de la plage. En lecture seule.|
|worksheet|[Worksheet](worksheet.md)|Feuille de calcul contenant la plage. En lecture seule.|

## Méthodes

| Méthode   | Type renvoyé|Description|
|:---------------|:--------|:----------|
|[clear(applyTo: string)](#clearapplyto-string)|void|Supprime les valeurs et les propriétés de format, de remplissage, de bordure, etc. de la plage.|
|[delete(shift: string)](#deleteshift-string)|void|Supprime les cellules associées à la plage.|
|[getBoundingRect(anotherRange: Range or string)](#getboundingrectanotherrange-range-or-string)|[Range](range.md)|Renvoie le plus petit objet de plage qui englobe les plages données. Par exemple, la valeur getBoundingRect pour « B2:C5 » et « D10:E15 » est « B2:E15 ».|
|[getCell(row: number, column: number)](#getcellrow-number-column-number)|[Range](range.md)|Renvoie l’objet de plage qui contient une cellule donnée sur la base des numéros de ligne et de colonne. La cellule peut se trouver en dehors des limites de ses plages parent, pour peu qu’elle reste dans la grille de la feuille de calcul. L’emplacement de la cellule renvoyée est déterminé à partir de la cellule supérieure gauche de la plage.|
|[getColumn(column: number)](#getcolumncolumn-number)|[Range](range.md)|Obtient une colonne contenue dans la plage.|
|[getEntireColumn()](#getentirecolumn)|[Range](range.md)|Obtient un objet qui représente la colonne entière de la plage.|
|[getEntireRow()](#getentirerow)|[Range](range.md)|Obtient un objet qui représente la ligne entière de la plage.|
|[getIntersection(anotherRange: Range or string)](#getintersectionanotherrange-range-or-string)|[Range](range.md)|Obtient l’objet de plage qui représente l’intersection rectangulaire des plages données.|
|[getLastCell()](#getlastcell)|[Range](range.md)|Obtient la dernière cellule de la plage. Par exemple, la dernière cellule de la plage « B2:D5 » est « D5 ».|
|[getLastColumn()](#getlastcolumn)|[Range](range.md)|Obtient la dernière colonne de la plage. Par exemple, la dernière colonne de la plage « B2:D5 » est « D2:D5 ».|
|[getLastRow()](#getlastrow)|[Range](range.md)|Obtient la dernière ligne de la plage. Par exemple, la dernière ligne de la plage « B2:D5 » est « B5:D5 ».|
|[getOffsetRange(rowOffset: number, columnOffset: number)](#getoffsetrangerowoffset-number-columnoffset-number)|[Range](range.md)|Obtient un objet qui représente une plage décalée par rapport à la plage spécifiée. Les dimensions de la plage renvoyée correspondent à celle de la plage initiale. Si la plage obtenue se retrouve en dehors des limites de la grille de la feuille de calcul, une exception est déclenchée.|
|[getRow(row: number)](#getrowrow-number)|[Range](range.md)|Obtient une ligne contenue dans la plage.|
|[getUsedRange()](#getusedrange)|[Range](range.md)|Renvoie la plage utilisée d’un objet de plage donné.|
|[insert(shift: string)](#insertshift-string)|[Range](range.md)|Insère une cellule ou une plage de cellules dans la feuille de calcul à la place d’une plage donnée et décale les autres cellules pour libérer de l’espace. Renvoie un nouvel objet Range dans l’espace vide qui s’est créé.|
|[load(param: object)](#loadparam-object)|void|Remplit l’objet proxy créé dans le calque JavaScript avec des valeurs de propriété et d’objet spécifiées dans le paramètre.|
|[select()](#select)|void|Sélectionne la plage spécifiée dans l’interface utilisateur d’Excel.|

## Détails des méthodes

### clear(applyTo: string)
Supprime les valeurs et les propriétés de format, de remplissage, de bordure, etc. de la plage.

#### Syntaxe
```js
rangeObject.clear(applyTo);
```

#### Paramètres
| Paramètre   | Type|Description|
|:---------------|:--------|:----------|
|applyTo|string|Facultatif. Détermine le type d’action de suppression. Les valeurs possibles sont les suivantes : `All` (option par défaut),`Formats` ,`Contents`|

#### Retourne
void

#### Exemples

L’exemple ci-dessous efface le format et le contenu de la plage. 

```js
Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "D:F";
	var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
	range.clear();
	return ctx.sync(); 
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

### delete(shift: string)
Supprime les cellules associées à la plage.

#### Syntaxe
```js
rangeObject.delete(shift);
```

#### Paramètres
| Paramètre   | Type|Description|
|:---------------|:--------|:----------|
|Shift|string|Indique la façon dont les cellules doivent être décalées.  Les valeurs possibles sont les suivantes : Up (vers le haut), Left (vers la gauche)|

#### Retourne
void

#### Exemples

```js
Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "D:F";
	var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
	range.delete();
	return ctx.sync(); 
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

### getBoundingRect(anotherRange: Range or string)
Renvoie le plus petit objet de plage qui englobe les plages données. Par exemple, la valeur GetBoundingRect pour « B2:C5 » et « D10:E15 » est « B2:E15 ».

#### Syntaxe
```js
rangeObject.getBoundingRect(anotherRange);
```

#### Paramètres
| Paramètre   | Type|Description|
|:---------------|:--------|:----------|
|anotherRange|range ou string|Nom, adresse ou objet de plage.|

#### Retourne
[Range](range.md)

#### Exemples

```js

Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "D4:G6";
	var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
	var range = range.getBoundingRect("G4:H8");
	range.load('address');
	return ctx.sync().then(function() {
		console.log(range.address); // Prints Sheet1!D4:H8
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

### getCell(row: number, column: number)
Renvoie l’objet de plage qui contient une cellule donnée sur la base des numéros de ligne et de colonne. La cellule peut se trouver en dehors des limites de ses plages parent, pour peu qu’elle reste dans la grille de la feuille de calcul. L’emplacement de la cellule renvoyée est déterminé à partir de la cellule supérieure gauche de la plage.

#### Syntaxe
```js
rangeObject.getCell(row, column);
```

#### Paramètres
| Paramètre   | Type|Description|
|:---------------|:--------|:----------|
|row|number|Numéro de ligne de la cellule à récupérer. Avec indice zéro.|
|column|number|Numéro de colonne de la cellule à récupérer. Avec indice zéro.|

#### Retourne
[Range](range.md)

#### Exemples

```js
Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "A1:F8";
	var worksheet = ctx.workbook.worksheets.getItem(sheetName);
	var range = worksheet.getRange(rangeAddress);
	var cell = range.cell(0,0);
	cell.load('address');
	return ctx.sync().then(function() {
		console.log(cell.address);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

### getColumn(column: number)
Obtient une colonne contenue dans la plage.

#### Syntaxe
```js
rangeObject.getColumn(column);
```

#### Paramètres
| Paramètre   | Type|Description|
|:---------------|:--------|:----------|
|column|number|Numéro de colonne de la plage à récupérer. Avec indice zéro.|

#### Retourne
[Range](range.md)

#### Exemples

```js

Excel.run(function (ctx) { 
	var sheetName = "Sheet19";
	var rangeAddress = "A1:F8";
	var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress).getColumn(1);
	range.load('address');
	return ctx.sync().then(function() {
		console.log(range.address); // prints Sheet1!B1:B8
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

### getEntireColumn()
Obtient un objet qui représente la colonne entière de la plage.

#### Syntaxe
```js
rangeObject.getEntireColumn();
```

#### Paramètres
Aucun

#### Retourne
[Range](range.md)

#### Exemples

Remarque : les propriétés de grille de la plage (valeurs, format de nombre, formules) contiennent la valeur `null` car la plage en question est illimitée.

```js

Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "D:F";
	var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
	var rangeEC = range.getEntireColumn();
	rangeEC.load('address');
	return ctx.sync().then(function() {
		console.log(rangeEC.address);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```
### getEntireRow()
Obtient un objet qui représente la ligne entière de la plage.

#### Syntaxe
```js
rangeObject.getEntireRow();
```

#### Paramètres
Aucun

#### Retourne
[Range](range.md)

#### Exemples
```js

Excel.run(function (ctx) {
	var sheetName = "Sheet1";
	var rangeAddress = "D:F"; 
	var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
	var rangeER = range.getEntireRow();
	rangeER.load('address');
	return ctx.sync().then(function() {
		console.log(rangeER.address);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```
Les propriétés de grille de la plage (valeurs, format de nombre, formules) contiennent la valeur `null` car la plage en question est illimitée.

### getIntersection(anotherRange: Range or string)
Obtient l’objet de plage qui représente l’intersection rectangulaire des plages données.

#### Syntaxe
```js
rangeObject.getIntersection(anotherRange);
```

#### Paramètres
| Paramètre   | Type|Description|
|:---------------|:--------|:----------|
|anotherRange|range ou string|Objet de plage ou adresse de plage utilisé pour déterminer l’intersection des plages.|

#### Retourne
[Range](range.md)

#### Exemples

```js

Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "A1:F8";
	var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress).getIntersection("D4:G6");
	range.load('address');
	return ctx.sync().then(function() {
		console.log(range.address); // prints Sheet1!D4:F6
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

### getLastCell()
Obtient la dernière cellule de la plage. Par exemple, la dernière cellule de la plage « B2:D5 » est « D5 ».

#### Syntaxe
```js
rangeObject.getLastCell();
```

#### Paramètres
Aucun

#### Retourne
[Range](range.md)

#### Exemples

```js

Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "A1:F8";
	var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress).getLastCell();
	range.load('address');
	return ctx.sync().then(function() {
		console.log(range.address); // prints Sheet1!F8
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

### getLastColumn()
Obtient la dernière colonne de la plage. Par exemple, la dernière colonne de la plage « B2:D5 » est « D2:D5 ».

#### Syntaxe
```js
rangeObject.getLastColumn();
```

#### Paramètres
Aucun

#### Retourne
[Range](range.md)

#### Exemples

```js

Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "A1:F8";
	var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress).getLastColumn();
	range.load('address');
	return ctx.sync().then(function() {
		console.log(range.address); // prints Sheet1!F1:F8
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

### getLastRow()
Obtient la dernière ligne de la plage. Par exemple, la dernière ligne de la plage « B2:D5 » est « B5:D5 ».

#### Syntaxe
```js
rangeObject.getLastRow();
```

#### Paramètres
Aucun

#### Retourne
[Range](range.md)

#### Exemples

```js

Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "A1:F8";
	var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress).getLastRow();
	range.load('address');
	return ctx.sync().then(function() {
		console.log(range.address); // prints Sheet1!A8:F8
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```


### getOffsetRange(rowOffset: number, columnOffset: number)
Obtient un objet qui représente une plage décalée par rapport à la plage spécifiée. Les dimensions de la plage renvoyée correspondent à celle de la plage initiale. Si la plage obtenue se retrouve en dehors des limites de la grille de la feuille de calcul, une exception est déclenchée.

#### Syntaxe
```js
rangeObject.getOffsetRange(rowOffset, columnOffset);
```

#### Paramètres
| Paramètre   | Type|Description|
|:---------------|:--------|:----------|
|rowOffset|number|Nombre de lignes (positif, négatif ou nul) duquel décaler la plage. Les valeurs positives représentent un décalage vers le bas et les valeurs négatives un décalage vers le haut.|
|columnOffset|number|Nombre de colonnes (positif, négatif ou nul) duquel décaler la plage. Les valeurs positives représentent un décalage vers la droite et les valeurs négatives un décalage vers la gauche.|

#### Retourne
[Range](range.md)

#### Exemples

```js
Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "D4:F6";
	var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress).getOffsetRange(-1,4);
	range.load('address');
	return ctx.sync().then(function() {
		console.log(range.address); // prints Sheet1!H3:K5
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

### getRow(row: number)
Obtient une ligne contenue dans la plage.

#### Syntaxe
```js
rangeObject.getRow(row);
```

#### Paramètres
| Paramètre   | Type|Description|
|:---------------|:--------|:----------|
|row|number|Numéro de ligne de la plage à récupérer. Avec indice zéro.|

#### Retourne
[Range](range.md)

#### Exemples

```js

Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "A1:F8";
	var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress).getRow(1);
	range.load('address');
	return ctx.sync().then(function() {
		console.log(range.address); // prints Sheet1!A2:F2
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

### getUsedRange()
Renvoie la plage utilisée d’un objet de plage donné.

#### Syntaxe
```js
rangeObject.getUsedRange();
```

#### Paramètres
Aucun

#### Retourne
[Range](range.md)

#### Exemples

```js

Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "D:F";
	var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
	var rangeUR = range.getUsedRange();
	rangeUR.load('address');
	return ctx.sync().then(function() {
		console.log(rangeUR.address);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

### insert(shift: string)
Insère une cellule ou une plage de cellules dans la feuille de calcul à la place d’une plage donnée et décale les autres cellules pour libérer de l’espace. Renvoie un nouvel objet Range dans l’espace vide qui s’est créé.

#### Syntaxe
```js
rangeObject.insert(shift);
```

#### Paramètres
| Paramètre   | Type|Description|
|:---------------|:--------|:----------|
|Shift|string|Indique la façon dont les cellules doivent être décalées.  Les valeurs possibles sont les suivantes : Down (vers le bas), Right (vers la droite)|

#### Retourne
[Range](range.md)

#### Exemples

```js
	
Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "F5:F10";
	var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
	range.insert();
	return ctx.sync(); 
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
### select()
Sélectionne la plage spécifiée dans l’interface utilisateur d’Excel.

#### Syntaxe
```js
rangeObject.select();
```

#### Paramètres
Aucun

#### Retourne
void

#### Exemples

```js

Excel.run(function (ctx) {
	var sheetName = "Sheet1";
	var rangeAddress = "F5:F10"; 
	var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
	range.select();
	return ctx.sync(); 
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

### Exemples d’accès aux propriétés

Cet exemple utilise l’adresse de la plage pour obtenir l’objet de la plage.

```js

Excel.run(function (ctx) {
	var sheetName = "Sheet1";
	var rangeAddress = "A1:F8"; 
	var worksheet = ctx.workbook.worksheets.getItem(sheetName);
	var range = worksheet.getRange(rangeAddress);
	range.load('cellCount');
	return ctx.sync().then(function() {
		console.log(range.cellCount);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

Cet exemple utilise une plage nommée pour obtenir l’objet de la plage.

```js

Excel.run(function (ctx) { 
	var rangeName = 'MyRange';
	var range = ctx.workbook.names.getItem(rangeName).range;
	range.load('cellCount');
	return ctx.sync().then(function() {
		console.log(range.cellCount);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

L’exemple ci-dessous définit le format de nombre, les valeurs et les formules dans une grille 2x3.

```js
Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "F5:G7";
	var numberFormat = [[null, "d-mmm"], [null, "d-mmm"], [null, null]]
	var values = [["Today", 42147], ["Tomorrow", "5/24"], ["Difference in days", null]];
	var formulas = [[null,null], [null,null], [null,"=G6-G5"]];
	var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
	range.numberFormat = numberFormat;
	range.values = values;
	range.formulas= formulas;
	range.load('text');
	return ctx.sync().then(function() {
		console.log(range.text);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```
Obtenir la feuille de calcul contenant la plage 

```js
Excel.run(function (ctx) { 
	var names = ctx.workbook.names;
	var namedItem = names.getItem('MyRange');
	range = namedItem.range;
	var rangeWorksheet = range.worksheet;
	rangeWorksheet.load('name');
	return ctx.sync().then(function() {
			console.log(rangeWorksheet.name);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

