
# Labs.Components.InputComponentInstance

 _**S’applique à :** apps for Office | Office Add-ins | Office Mix | PowerPoint_

Instance d’un composant de saisie.

```
class InputComponentInstance extends Labs.ComponentInstance<Components.InputComponentAttempt>
```


## Propriétés


|Propriété|Description|
|:-----|:-----|
| `public var component: Components.IInputComponentInstance`|Objet [Labs.Components.IInputComponentInstance](../../reference/office-mix/labs.components.iinputcomponentinstance.md) sous-jacent représenté par cette classe.|

## Méthodes




### constructeur

 `function constructor(component: Components.IInputComponentInstance)`

Crée une instance [Labs.Components.IInputComponentInstance](../../reference/office-mix/labs.components.iinputcomponentinstance.md).

 **Paramètres**


|Paramètre|Description|
|:-----|:-----|
| _composant_|Objet [Labs.Components.IInputComponentInstance](../../reference/office-mix/labs.components.iinputcomponentinstance.md) à utiliser pour créer cette classe.|

### buildAttempt

 `public function buildAttempt(createAttemptAction: Labs.Core.IAction): Components.InputComponentAttempt`

Crée une instance [Labs.Components.InputComponentAttempt](../../reference/office-mix/labs.components.inputcomponentattempt.md). Implémente la méthode abstraite définie dans la classe de base.

 **Paramètres**


|Paramètre|Description|
|:-----|:-----|
| _createAttemptResult_|Résultat d’une tentative de création.|
