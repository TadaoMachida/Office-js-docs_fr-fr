# Objet TableCollection (interface API JavaScript pour Excel)

Représente une collection de tous les tableaux du classeur.

## Propriétés

| Propriété     | Type   |Description
|:---------------|:--------|:----------|
|count|int|Renvoie le nombre de tableaux dans le classeur. En lecture seule.|
|Items|[Table[]](table.md)|Collection d’objets de tableau. En lecture seule.|

_Voir des [exemples](#exemples) d’accès aux propriétés._

## Relations
Aucun


## Méthodes

| Méthode           | Type renvoyé    |Description|
|:---------------|:--------|:----------|
|[add(address: string, hasHeaders: bool)](#addaddress-string-hasheaders-bool)|[Table](table.md)|Crée un tableau. L’adresse de la source de la plage détermine la feuille de calcul dans laquelle le tableau sera ajouté. Si l’ajout ne peut être effectué (par exemple, parce que l’adresse n’est pas valide, ou parce que le tableau empiéterait sur un autre tableau), un message d’erreur apparaît.|
|[getItem(key: number ou string)](#getitemkey-number-ou-string)|[Table](table.md)|Obtient un tableau à l’aide de son nom ou de son ID.|
|[getItemAt(index: number)](#getitematindex-number)|[Table](table.md)|Obtient un tableau en fonction de sa position dans la collection.|
|[load(param: object)](#loadparam-object)|void|Remplit l’objet proxy créé dans le calque JavaScript avec des valeurs de propriété et d’objet spécifiées dans le paramètre.|

## Détails des méthodes


### add(address: string, hasHeaders: bool)
Crée un tableau. L’adresse de la source de la plage détermine la feuille de calcul dans laquelle le tableau sera ajouté. Si l’ajout ne peut être effectué (par exemple, parce que l’adresse n’est pas valide, ou parce que le tableau empiéterait sur un autre tableau), un message d’erreur apparaît.

#### Syntaxe
```js
tableCollectionObject.add(address, hasHeaders);
```

#### Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|address|string|Adresse ou nom de l’objet de plage représentant la source de données. Si l’adresse ne contient pas de nom de feuille, la feuille ouverte est utilisée.|
|hasHeaders|bool|Valeur booléenne qui indique si les données importées disposent d’étiquettes de colonne. Si la source ne contient pas d’en-têtes (autrement dit, lorsque cette propriété est définie sur false), Excel génère automatiquement un en-tête et décale les données d’une ligne vers le bas.|

#### Retourne
[Table](table.md)

#### Exemples

```js
Excel.run(function (ctx) { 
    var table = ctx.workbook.tables.add('Sheet1!A1:E7', true);
    table.load('name');
    return ctx.sync().then(function() {
        console.log(table.name);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

### getItem(key: number or string)
Obtient un tableau à l’aide de son nom ou de son ID.

#### Syntaxe
```js
tableCollectionObject.getItem(key);
```

#### Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|Key|number or string|Nom ou ID du tableau à récupérer.|

#### Retourne
[Table](table.md)

#### Exemples

```js
Excel.run(function (ctx) { 
    var tableName = 'Table1';
    var table = ctx.workbook.tables.getItem(tableName);
    return ctx.sync().then(function() {
            console.log(table.index);
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
    var table = ctx.workbook.tables.getItemAt(0);
    return ctx.sync().then(function() {
            console.log(table.name);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```


### getItemAt(index: number)
Obtient un tableau en fonction de sa position dans la collection.

#### Syntaxe
```js
tableCollectionObject.getItemAt(index);
```

#### Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|index|number|Valeur d’indice de l’objet à récupérer. Avec indice zéro.|

#### Retourne
[Table](table.md)

#### Exemples

```js
Excel.run(function (ctx) { 
    var table = ctx.workbook.tables.getItemAt(0);
    return ctx.sync().then(function() {
            console.log(table.name);
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
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|param|object|Facultatif. Accepte les noms de paramètre et de relation sous forme de chaîne délimitée ou de tableau. Sinon, indiquez l’objet [loadOption](loadoption.md).|

#### Renvoie
void
### Exemples d’accès aux propriétés

```js
Excel.run(function (ctx) { 
    var tables = ctx.workbook.tables;
    tables.load('items');
    return ctx.sync().then(function() {
        console.log("tables Count: " + tables.count);
        for (var i = 0; i < tables.items.length; i++)
        {
            console.log(tables.items[i].name);
        }
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

Obtenir le nombre de tableaux

```js
Excel.run(function (ctx) { 
    var tables = ctx.workbook.tables;
    tables.load('count');
    return ctx.sync().then(function() {
        console.log(tables.count);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
