# Objet NamedItem (interface API JavaScript pour Excel)

_S’applique à : Excel 2016, Excel Online, Office 2016_

Représente un nom défini pour une plage de cellules ou une valeur. Les noms peuvent être des objets nommés primitifs (comme dans le type ci-dessous), un objet de plage et une référence à une plage. Cet objet peut être utilisé pour obtenir un objet de plage associé aux noms.

## Propriétés

| Propriété   | Type|Description
|:---------------|:--------|:----------|
|name|string|Nom de l’objet. En lecture seule.|
|type|string|Indique le type de référence associé au nom. En lecture seule. Les valeurs possibles sont les suivantes : String, Integer, Double, Boolean, Range.|
|value|object|Représente la formule à laquelle le nom doit faire référence. Par exemple, =Sheet14!$B$2:$H$12, =4.75, etc. En lecture seule.|
|visible|bool|Indique si l’objet est visible ou non.|

_Voir des [exemples](#property-access-examples) d’accès aux propriétés._

## Relations
Aucun


## Méthodes

| Méthode   | Type renvoyé|Description|
|:---------------|:--------|:----------|
|[getRange()](#getrange)|[Range](range.md)|Renvoie l’objet de plage qui est associé au nom. Renvoie une exception si le type de l’élément nommé n’est pas une plage.|
|[load(param: object)](#loadparam-object)|void|Remplit l’objet proxy créé dans le calque JavaScript avec des valeurs de propriété et d’objet spécifiées dans le paramètre.|

## Détails des méthodes

### getRange()
Renvoie l’objet de plage qui est associé au nom. Renvoie une exception si le type de l’élément nommé n’est pas une plage.

#### Syntaxe
```js
namedItemObject.getRange();
```

#### Paramètres
Aucun

#### Retourne
[Range](range.md)

#### Exemples

Renvoie l’objet de plage qui est associé au nom. Renvoie `null` si le nom n’est pas du type `Range`. Remarque : actuellement, cette API prend uniquement en charge les éléments de classeur inclus dans l’étendue.

```js
Excel.run(function (ctx) { 
	var names = ctx.workbook.names;
	var range = names.getItem('MyRange').getRange();
	range.load('address');
	return ctx.sync().then(function() {
			console.log(range.address);
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
	var names = ctx.workbook.names;
	var namedItem = names.getItem('MyRange');
	namedItem.load('type');
	return ctx.sync().then(function() {
			console.log(namedItem.type);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

