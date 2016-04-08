# Objet FormatProtection (interface API JavaScript pour Excel)

_S’applique à : Excel 2016, Excel Online, Excel pour iOS, Office 2016_

Cet objet représente la protection du format d’un objet Range.

## Propriétés

| Propriété	   | Type	|Description
| :---------------| :--------| :----------||formulaHidden|bool|Indique si Excel masque la formule des cellules dans la plage. Une valeur Null indique que le paramètre masqué de formule n’est pas le même pour l’ensemble de la plage.||locked|bool|Indique si Excel verrouille les cellules de l’objet. Une valeur Null indique que le paramètre de verrouillage n’est pas le même pour l’ensemble de la plage.|_Consultez les [exemples](#property-access-examples) d’accès aux propriétés._

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

