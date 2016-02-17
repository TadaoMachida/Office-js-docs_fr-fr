# Objet BindingCollection (interface API JavaScript pour Excel)

_S’applique à : Excel 2016, Excel Online, Office 2016_

Représente la collection de tous les objets de liaison qui font partie du classeur.

## Propriétés

| Propriété   | Type|Description
|:---------------|:--------|:----------|
|count|int|Renvoie le nombre de liaisons de la collection. En lecture seule.|
|Items|[binding[]](binding.md)|Collection d’objets de liaison. En lecture seule.|

_Voir des [exemples](#property-access-examples) d’accès aux propriétés._

## Relations
Aucun


## Méthodes

| Méthode   | Type renvoyé|Description|
|:---------------|:--------|:----------|
|[getItem(id: string)](#getitemid-string)|[binding](binding.md)|Obtient un objet de liaison par ID.|
|[getItemAt(index: number)](#getitematindex-number)|[binding](binding.md)|Obtient un objet de liaison en fonction de sa position dans le tableau d’éléments.|
|[load(param: object)](#loadparam-object)|void|Remplit l’objet proxy créé dans le calque JavaScript avec des valeurs de propriété et d’objet spécifiées dans le paramètre.|

## Détails des méthodes

### getItem(id: string)
Obtient un objet de liaison par ID.

#### Syntaxe
```js
bindingCollectionObject.getItem(id);
```

#### Paramètres
| Paramètre   | Type|Description|
|:---------------|:--------|:----------|
|id|string|ID de l’objet de liaison à récupérer.|

#### Retourne
[binding](binding.md)

#### Exemples

Créez une liaison de table pour contrôler les modifications apportées aux données de la table. Lorsque les données sont modifiées, la couleur d’arrière-plan du tableau devient orange.

```js
function addEventHandler() {
	//Create Table1
Excel.run(function (ctx) { 
	ctx.workbook.tables.add("Sheet1!A1:C4", true);
	return ctx.sync().then(function() {
			 console.log("My Diet Data Inserted!");
	})
	.catch(function (error) {
			 console.log(JSON.stringify(error));
	});
});
	//Create a new table binding for Table1
Office.context.document.bindings.addFromNamedItemAsync("Table1", Office.CoercionType.Table, { id: "myBinding" }, function (asyncResult) {
	if (asyncResult.status == "failed") {
		console.log("Action failed with error: " + asyncResult.error.message);
	}
	else {
		// If successful, add the event handler to the table binding.
		Office.select("bindings#myBinding").addHandlerAsync(Office.EventType.BindingDataChanged, onBindingDataChanged);
	}
});
}
	
// When data in the table is changed, this event is triggered.
function onBindingDataChanged(eventArgs) {
Excel.run(function (ctx) { 
	// Highlight the table in orange to indicate data changed.
	ctx.workbook.bindings.getItem(eventArgs.binding.id).getTable().getDataBodyRange().format.fill.color = "Orange";
	return ctx.sync().then(function() {
			console.log("The value in this table got changed!");
	})
	.catch(function (error) {
			console.log(JSON.stringify(error));
	});
});
}

```



#### Exemples
```js
Excel.run(function (ctx) { 
	var lastPosition = ctx.workbook.bindings.count - 1;
	var binding = ctx.workbook.bindings.getItemAt(lastPosition);
	binding.load('type')
	return ctx.sync().then(function() {
			console.log(binding.type); 
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

### getItemAt(index: number)
Obtient un objet de liaison en fonction de sa position dans le tableau d’éléments.

#### Syntaxe
```js
bindingCollectionObject.getItemAt(index);
```

#### Paramètres
| Paramètre   | Type|Description|
|:---------------|:--------|:----------|
|index|number|Valeur d’indice de l’objet à récupérer. Avec indice zéro.|

#### Retourne
[binding](binding.md)

#### Exemples
```js
Excel.run(function (ctx) { 
	var lastPosition = ctx.workbook.bindings.count - 1;
	var binding = ctx.workbook.bindings.getItemAt(lastPosition);
	binding.load('type')
	return ctx.sync().then(function() {
			console.log(binding.type); 
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
|param|object|Facultatif. Accepte les noms de paramètre et de relation sous forme de chaîne délimitée ou de tableau. Sinon, accepte un objet [loadOption](loadoption.md).|

#### Renvoie
void

### Exemples d’accès aux propriétés

```js
Excel.run(function (ctx) { 
	var bindings = ctx.workbook.bindings;
	bindings.load('items');
	return ctx.sync().then(function() {
		for (var i = 0; i < bindings.items.length; i++)
		{
			console.log(bindings.items[i].id);
			console.log(bindings.items[i].index);
		}
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```
Obtenir le nombre de liaisons

```js
Excel.run(function (ctx) { 
	var bindings = ctx.workbook.bindings;
	bindings.load('count');
	return ctx.sync().then(function() {
		console.log("Bindings: Count= " + bindings.count);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

