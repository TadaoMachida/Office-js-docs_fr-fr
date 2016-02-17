# Objet TableColumn (interface API JavaScript pour Excel)

_S’applique à : Excel 2016, Excel Online, Office 2016_

Cet objet représente une colonne dans un tableau.

## Propriétés

| Propriété   | Type|Description
|:---------------|:--------|:----------|
|id|int|Renvoie une clé unique qui identifie la colonne dans le tableau. En lecture seule.|
|index|int|Renvoie le numéro d’indice de la colonne dans la collection de colonnes du tableau. Avec indice zéro. En lecture seule.|
|name|string|Renvoie le nom de la colonne du tableau. En lecture seule.|
|values|object[][]|Représente les valeurs brutes de la plage spécifiée. Les données renvoyées peuvent être des chaînes, des valeurs numériques ou des valeurs booléennes. Une cellule contenant une erreur renvoie une chaîne d’erreur.|

_Voir des [exemples](#property-access-examples) d’accès aux propriétés._

## Relations
Aucun


## Méthodes

| Méthode   | Type renvoyé|Description|
|:---------------|:--------|:----------|
|[delete()](#delete)|void|Supprime la colonne du tableau.|
|[getDataBodyRange()](#getdatabodyrange)|[Range](range.md)|Obtient l’objet de plage associé au corps de données de la colonne.|
|[getHeaderRowRange()](#getheaderrowrange)|[Range](range.md)|Obtient l’objet de plage associé à la ligne d’en-tête de la colonne.|
|[getRange()](#getrange)|[Range](range.md)|Renvoie l’objet de plage associé à l’intégralité de la colonne.|
|[getTotalRowRange()](#gettotalrowrange)|[Range](range.md)|Obtient l’objet de plage associé à la ligne de total de la colonne.|
|[load(param: object)](#loadparam-object)|void|Remplit l’objet proxy créé dans le calque JavaScript avec des valeurs de propriété et d’objet spécifiées dans le paramètre.|

## Détails des méthodes

### delete()
Supprime la colonne du tableau.

#### Syntaxe
```js
tableColumnObject.delete();
```

#### Paramètres
Aucun

#### Retourne
void

#### Exemples

```js
Excel.run(function (ctx) { 
	var tableName = 'Table1';
	var column = ctx.workbook.tables.getItem(tableName).tableColumns.getItemAt(2);
	column.delete();
	return ctx.sync(); 
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

### getDataBodyRange()
Obtient l’objet de plage associé au corps de données de la colonne.

#### Syntaxe
```js
tableColumnObject.getDataBodyRange();
```

#### Paramètres
Aucun

#### Retourne
[Range](range.md)

#### Exemples

```js
Excel.run(function (ctx) { 
	var tableName = 'Table1';
	var column = ctx.workbook.tables.getItem(tableName).tableColumns.getItemAt(0);
	var dataBodyRange = column.getDataBodyRange();
	dataBodyRange.load('address');
	return ctx.sync().then(function() {
		console.log(dataBodyRange.address);
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

### getHeaderRowRange()
Obtient l’objet de plage associé à la ligne d’en-tête de la colonne.

#### Syntaxe
```js
tableColumnObject.getHeaderRowRange();
```

#### Paramètres
Aucun

#### Retourne
[Range](range.md)

#### Exemples

```js
Excel.run(function (ctx) { 
	var tableName = 'Table1';
	var columns = ctx.workbook.tables.getItem(tableName).tableColumns.getItemAt(0);
	var headerRowRange = columns.getHeaderRowRange();
	headerRowRange.load('address');
	return ctx.sync().then(function() {
		console.log(headerRowRange.address);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```
### getRange()
Renvoie l’objet de plage associé à l’intégralité de la colonne.

#### Syntaxe
```js
tableColumnObject.getRange();
```

#### Paramètres
Aucun

#### Retourne
[Range](range.md)

#### Exemples

```js
Excel.run(function (ctx) { 
	var tableName = 'Table1';
	var columns = ctx.workbook.tables.getItem(tableName).tableColumns.getItemAt(0);
	var columnRange = columns.getRange();
	columnRange.load('address');
	return ctx.sync().then(function() {
		console.log(columnRange.address);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

### getTotalRowRange()
Obtient l’objet de plage associé à la ligne de total de la colonne.

#### Syntaxe
```js
tableColumnObject.getTotalRowRange();
```

#### Paramètres
Aucun

#### Retourne
[Range](range.md)

#### Exemples

```js
Excel.run(function (ctx) { 
	var tableName = 'Table1';
	var columns = ctx.workbook.tables.getItem(tableName).tableColumns.getItemAt(0);
	var totalRowRange = columns.getTotalRowRange();
	totalRowRange.load('address');
	return ctx.sync().then(function() {
		console.log(totalRowRange.address);
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
	var tableName = 'Table1';
	var column = ctx.workbook.tables.getItem(tableName).tableColumns.getItem(0);
	column.load('index');
	return ctx.sync().then(function() {
		console.log(column.index);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

```js
Excel.run(function (ctx) { 
	var tables = ctx.workbook.tables;
	var newValues = [["New"], ["Values"], ["For"], ["New"], ["Column"]];
	var column = ctx.workbook.tables.getItem(tableName).tableColumns.getItemAt(2);
	column.values = newValues;
	column.load('values');
	return ctx.sync().then(function() {
		console.log(column.values);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```
