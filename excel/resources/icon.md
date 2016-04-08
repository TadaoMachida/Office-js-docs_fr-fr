# Objet Icon (interface API JavaScript pour Excel)

_S’applique à : Excel 2016, Excel Online, Excel pour iOS, Office 2016_

Cet objet représente une icône de cellule.

## Propriétés

| Propriété	   | Type	|Description
| :---------------| :--------| :----------||index|int|Index de l’icône dans le jeu de données indiqué.||set|string|Jeu de données dont l’icône fait partie. Les valeurs possibles sont les suivantes : Invalid, ThreeArrows, ThreeArrowsGray, ThreeFlags, ThreeTrafficLights1, ThreeTrafficLights2, ThreeSigns, ThreeSymbols, ThreeSymbols2, FourArrows, FourArrowsGray, FourRedToBlack, FourRating, FourTrafficLights, FiveArrows, FiveArrowsGray, FiveRating, FiveQuarters, ThreeStars, ThreeTriangles, FiveBoxes.|_Consultez les [exemples](#property-access-examples) d’accès aux propriétés._

## Relations
Aucune


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
| Paramètre	   | Type	|Description|| :---| :---| :---||param|object|Facultatif. Accepte les noms de paramètre et de relation sous forme de chaîne délimitée ou de tableau. Sinon, indiquez l’objet [loadOption](loadoption.md).|

#### Renvoie
void

