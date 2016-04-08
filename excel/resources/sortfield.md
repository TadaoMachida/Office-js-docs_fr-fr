# Objet SortField (interface API JavaScript pour Excel)

_S’applique à : Excel 2016, Excel Online, Excel pour iOS, Office 2016_

Cet objet représente une condition dans une opération de tri.

## Propriétés

| Propriété	   | Type	|Description
| :---------------| :--------| :----------||ascending|bool|Indique si le tri est effectué dans l’ordre croissant.||color|string|Couleur de la cible de la condition si le tri s’effectue en fonction de la couleur de la police ou de la cellule.||dataOption|string|Options de tri supplémentaires pour ce champ. Les valeurs possibles sont les suivantes : Normal, TextAsNumber.||key|int|Colonne (ou ligne, selon l’orientation du tri) ciblée par la condition. Représentée sous forme d’un décalage par rapport à la première colonne (ou ligne).||sortOn|string|Type de tri pour cette condition. Les valeurs possibles sont les suivantes : Value, CellColor, FontColor, Icon.|_Consultez les [exemples](#property-access-examples) d’accès aux propriétés._

## Relations
| Relation | Type	|Description|| :---------------| :--------| :----------||icon|[Icône](icon.md)|Icône ciblée par la condition si le tri est appliqué à l’icône de la cellule.|

## Méthodes

| Méthode		   | Type de retour	|Description|| :---------------| :--------| :----------||[load(param: object)](#loadparam-object)|void|Remplit l’objet proxy créé dans le calque JavaScript avec des valeurs de propriété et d’objet spécifiées dans le paramètre.|

## Détails de la méthode


### load(param: object)
Remplit l’objet proxy créé dans le calque JavaScript avec des valeurs de propriété et d’objet spécifiées dans le paramètre.

#### Syntaxe
```js
object.load(param);
```

#### Paramètres
| Paramètre	   | Type	|Description|| :---------------| :--------| :----------||param|object|Facultatif. Accepte les noms de paramètre et de relation sous forme de chaîne délimitée ou de tableau. Sinon, indiquez l’objet [loadOption](loadoption.md).|

#### Renvoie
void

