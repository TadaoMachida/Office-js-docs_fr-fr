# Objet TrackedObjectsCollection (interface API JavaScript pour Office 2016)

_S’applique à : Excel 2016, Excel Online, Office 2016_

Permet à des compléments de gérer des références d’objet de plage entre plusieurs lots sync(). En règle générale, la méthode Excel.run() permet de mettre à jour les références dans tous les lots de façon automatique, sans que vous ayez à effectuer de suivi explicitement. Toutefois, si un objet de plage doit être suivi et ajusté manuellement pour qu’il reflète l’état actuel de la plage Excel sous-jacente, cette collection peut être utilisée afin de marquer ces objets pour le suivi. Notez que si un objet de plage est marqué pour être suivi, il doit être explicitement supprimé lorsqu’il est utilisé, afin de libérer de la mémoire dans Excel, même en cas d’erreur.

## Propriétés
Aucune.

## Relations

Aucun

## Méthodes

Les méthodes suivantes définies pour l’objet trackedObjectsCollection :

| Méthode     | Type renvoyé    |Description|
|:-----------------|:--------|:----------|
|[add(rangeObject: Range)](#addrangeobject-range)| Null             |Crée une nouvelle référence sur une plage.|
|[remove(rangeObject: Range)](#removerangeobject-range)| Null             |Supprime une référence sur la plage.  |
|[removeAll()](#removeall)| Null|Supprime toutes les références créées par le complément sur l’appareil.|


## Spécification d’API 

### add(rangeObject: range)
Ajoute un objet de plage à la collection d’objets suivis. Les modifications sous-jacentes seront suivies pour toutes les demandes de traitement par lot et toutes les mises à jour de suivi seront appliquées à l’état actuel de l’objet de plage. 

#### Syntaxe
```js
trackedObjectsCollection.add(rangeObject);
```

#### Paramètres

Paramètre       | Type   | Description
--------------- | ------ | ------------
`rangeObject`  | [Range](range.md)| Objet de plage à ajouter à la collection d’objets suivis.

#### Renvoie
Null

#### Exemples

```js
var sheetName = "Sheet1";
var rangeAddress = "A1:B2";
var ctx = new Excel.RequestContext();
var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
ctx.trackedObjectsCollection.add(range);
ctx.load(range);

Excel.run(function (ctx) { 
	range.insert("Down");
	Console.log(range.address); // Address should be updated to A3:B4
	return ctx.sync(); 
});
```


### remove(rangeObject: range)

Supprime un objet de référence de la collection. Cette opération libère de la mémoire et des ressources nécessaires pour gérer l’état de l’objet suivi. Notez que si un objet de plage est marqué comme devant faire l’objet d’un suivi, il doit être explicitement supprimé, même en cas d’erreur.

#### Syntaxe
```js
trackedObjectsCollection.remove(rangeObject);
```

#### Paramètres

Paramètre       | Type   | Description
--------------- | ------ | ------------
`rangeObject`  | [Range](range.md)| Objet de plage à supprimer de la collection d’objets suivis.

#### Renvoie
Null

#### Exemples


```js
var sheetName = "Sheet1";
var rangeAddress = "A1:B2";
var ctx = new Excel.RequestContext();
var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
ctx.trackedObjectsCollection.add(range);
ctx.load(range);

Excel.run(function (ctx) { 
	range.insert("Down");
	Console.log(range.address); // Address should be updated to A3:B4
	ctx.trackedObjectsCollection.remove(range); 
	return ctx.sync(); 
});
```

### removeAll(rangeObject: range)

Supprime toutes les références créées par le complément sur l’appareil.

#### Syntaxe
```js
trackedObjectsCollection.removeAll();
```

#### Paramètres

Aucun

#### Retourne
Null

#### Exemples

```js
Excel.run(function (ctx) { 
	var sheetName = "Sheet1";
	var rangeAddress = "A1:B2";
	var ctx = new Excel.RequestContext();
	var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
	ctx.trackedObjectsCollection.add(range);
	ctx.load(range);
	range.insert("Down");
	Console.log(range.address); // Address should be updated to A3:B4
	ctx.trackedObjectsCollection.removeAll(); 
	return ctx.sync(); 
});
```

