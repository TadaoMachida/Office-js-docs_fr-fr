# Objet InkWordCollection (interface API JavaScript pour OneNote)

_S’applique à : OneNote Online_  


Représente une collection d’objets InkWord.

## Propriétés

| Propriété     | Type   |Description|Commentaires|
|:---------------|:--------|:----------|:-------|
|count|int|Renvoie le nombre d’objets InkWord dans la page. En lecture seule.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkWordCollection-count)|
|Items|[InkWord[]](inkword.md)|Collection d’objets inkWord. En lecture seule.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkWordCollection-items)|

_Voir des [exemples](#exemples) d’accès aux propriétés._

## Relations
Aucun


## Méthodes

| Méthode           | Type renvoyé    |Description| Commentaires|
|:---------------|:--------|:----------|:-------|
|[getItem(index: number or string)](#getitemindex-number-or-string)|[InkWord](inkword.md)|Obtient un objet InkWord en fonction de son ID ou de son index dans la collection. En lecture seule.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkWordCollection-getItem)|
|[getItemAt(index: number)](#getitematindex-number)|[InkWord](inkword.md)|Obtient un objet InkWord en fonction de sa position dans la collection.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkWordCollection-getItemAt)|
|[load(param: object)](#loadparam-object)|void|Remplit l’objet proxy créé dans le calque JavaScript avec des valeurs de propriété et d’objet spécifiées dans le paramètre.|[Activer](https://github.com/OfficeDev/office-js-docs/issues/new?title=OneNote-inkWordCollection-load)|

## Détails des méthodes


### getItem(index: number or string)
Obtient un objet InkWord en fonction de son ID ou de son index dans la collection. En lecture seule.

#### Syntaxe
```js
inkWordCollectionObject.getItem(index);
```

#### Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|index|number or string|ID de l’objet InkWord ou emplacement d’index de l’objet InkWord dans la collection.|

#### Retourne
[InkWord](inkword.md)

### getItemAt(index: number)
Obtient un objet InkWord en fonction de sa position dans la collection.

#### Syntaxe
```js
inkWordCollectionObject.getItemAt(index);
```

#### Paramètres
| Paramètre    | Type   |Description|
|:---------------|:--------|:----------|
|index|number|Valeur d’indice de l’objet à récupérer. Avec indice zéro.|

#### Retourne
[InkWord](inkword.md)

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
