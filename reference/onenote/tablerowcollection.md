# Objet TableRowCollection (interface API JavaScript pour OneNote)

_S’applique à : OneNote Online_  


Contient une collection d’objets TableRow.

## Propriétés

| Propriété     | Type   |Description|Commentaires|
|:---------------|:--------|:----------|:-------|
|count|int|Renvoie le nombre de lignes de tableau dans cette collection. En lecture seule.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableRowCollection-count)|
|items|[TableRow[]](tablerow.md)|Collection d’objets tableRow. En lecture seule.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableRowCollection-items)|

_Voir des [exemples](#exemples) d’accès aux propriétés._

## Relations
Aucun


## Méthodes

| Méthode           | Type renvoyé    |Description| Commentaires|
|:---------------|:--------|:----------|:-------|
|[getItem(index: number or string)](#getitemindex-number-or-string)|[TableRow](tablerow.md)|Obtient un objet de ligne de tableau en fonction de son ID ou de son index dans la collection. En lecture seule.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableRowCollection-getItem)|
|[getItemAt(index: number)](#getitematindex-number)|[TableRow](tablerow.md)|Obtient une ligne de tableau au niveau de sa position dans la collection.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableRowCollection-getItemAt)|
|[load(param: object)](#loadparam-object)|void|Remplit l’objet proxy créé dans le calque JavaScript avec des valeurs de propriété et d’objet spécifiées dans le paramètre.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-tableRowCollection-load)|

## Détails des méthodes


### getItem(index: number or string)
Obtient un objet de ligne de tableau en fonction de son ID ou de son index dans la collection. En lecture seule.

#### Syntaxe
```js
tableRowCollectionObject.getItem(index);
```

#### Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|index|number or string|Nombre qui identifie l’emplacement associé à l’index d’un objet de ligne de tableau.|

#### Retourne
[TableRow](tablerow.md)

### getItemAt(index: number)
Obtient une ligne de tableau au niveau de sa position dans la collection.

#### Syntaxe
```js
tableRowCollectionObject.getItemAt(index);
```

#### Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|index|number|Valeur d’indice de l’objet à récupérer. Avec indice zéro.|

#### Retourne
[TableRow](tablerow.md)

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

#### Retourne
void
