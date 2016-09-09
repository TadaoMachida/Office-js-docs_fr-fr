
# Labs.Components.InputComponentResult

 _**S’applique à :** apps for Office | Office Add-ins | Office Mix | PowerPoint_

Résultat d’un envoi de composant de saisie.

```
class InputComponentResult
```


## Propriétés


|Propriété|Description|
|:-----|:-----|
| `public var score: any`|Note associée à l’envoi.|
| `public var complete: boolean`|Indique si le résultat envoyé a mis fin à la tentative.  Indique **True** si la tentative est terminée.|

## Méthodes




### constructeur

 `function constructor(score: any, complete: boolean)`

Crée une instance de la classe **InputComponentResult**.

 **Paramètres**


|Paramètre|Description|
|:-----|:-----|
| _employés_|Note associée au résultat.|
| _Intégration_|Indique l’expression booléenne **true** si le résultat a mis fin à la tentative.|
