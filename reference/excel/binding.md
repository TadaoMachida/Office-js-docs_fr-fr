# Objet Binding (interface API JavaScript pour Excel)

Représente une liaison Office.js définie dans le classeur.

## Propriétés

| Propriété     | Type   |Description
|:---------------|:--------|:----------|
|id|string|Représente l’identificateur de liaison. En lecture seule.|
|type|string|Renvoie le type de la liaison. En lecture seule. Les valeurs possibles sont les suivantes : Range, Table, Text.|

_Voir des [exemples](#exemples) d’accès aux propriétés._

## Relations
Aucun


## Méthodes

| Méthode           | Type renvoyé    |Description|
|:---------------|:--------|:----------|
|[getRange()](#getrange)|[Range](range.md)|Renvoie la plage représentée par la liaison. Génère une erreur si la liaison n’est pas du type approprié.|
|[getTable()](#gettable)|[Table](table.md)|Renvoie le tableau représenté par la liaison. Génère une erreur si la liaison n’est pas du type approprié.|
|[getText()](#gettext)|string|Renvoie le texte représenté par la liaison. Génère une erreur si la liaison n’est pas du type approprié.|
|[load(param: object)](#loadparam-object)|void|Remplit l’objet proxy créé dans le calque JavaScript avec des valeurs de propriété et d’objet spécifiées dans le paramètre.|

## Détails des méthodes


### getRange()
Renvoie la plage représentée par la liaison. Génère une erreur si la liaison n’est pas du type approprié.

#### Syntaxe
```js
bindingObject.getRange();
```

#### Paramètres
Aucun

#### Retourne
[Range](range.md)

#### Exemples
L’exemple ci-dessous utilise un objet de liaison pour obtenir la plage associée.

```js
Excel.run(function (ctx) { 
    var binding = ctx.workbook.bindings.getItemAt(0);
    var range = binding.getRange();
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


### getTable()
Renvoie le tableau représenté par la liaison. Génère une erreur si la liaison n’est pas du type approprié.

#### Syntaxe
```js
bindingObject.getTable();
```

#### Paramètres
Aucun

#### Retourne
[Table](table.md)

#### Exemples
```js
Excel.run(function (ctx) { 
    var binding = ctx.workbook.bindings.getItemAt(0);
    var table = binding.getTable();
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


### getText()
Renvoie le texte représenté par la liaison. Génère une erreur si la liaison n’est pas du type approprié.

#### Syntaxe
```js
bindingObject.getText();
```

#### Paramètres
Aucun

#### Retourne
string

#### Exemples

```js
Excel.run(function (ctx) { 
    var binding = ctx.workbook.bindings.getItemAt(0);
    var text = binding.getText();
    ctx.load('text');
    return ctx.sync().then(function() {
        console.log(text);
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
|param|object|Facultatif. Accepte les noms de paramètre et de relation sous forme de chaîne délimitée ou de tableau. Sinon, accepte un objet [loadOption](loadoption.md).|

#### Renvoie
void
### Exemples d’accès aux propriétés

```js
Excel.run(function (ctx) { 
    var binding = ctx.workbook.bindings.getItemAt(0);
    binding.load('type');
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
