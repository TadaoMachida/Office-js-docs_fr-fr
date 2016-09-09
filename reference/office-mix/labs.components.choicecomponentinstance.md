
# Labs.Components.ChoiceComponentInstance

 _**S’applique à :** apps for Office | Office Add-ins | Office Mix | PowerPoint_

Instance d’un composant de choix.

```
class ChoiceComponentInstance extends Labs.ComponentInstance<Components.ChoiceComponentAttempt>
```


## Propriétés


|Propriété|Description|
|:-----|:-----|
| `public var component: Components.IChoiceComponentInstance`|Instance du composant [Labs.Components.IChoiceComponentInstance](../../reference/office-mix/labs.components.ichoicecomponentinstance.md) sous-jacent représentée par cette classe.|

## Méthodes




### constructeur

 `function constructor(component: Components.IChoiceComponentInstance)`

Crée une instance de la classe **ChoiceComponentInstance**.

 **Paramètres**


|Paramètre|Description|
|:-----|:-----|
| _composant_|Objet [Labs.Components.IChoiceComponentInstance](../../reference/office-mix/labs.components.ichoicecomponentinstance.md) à utiliser pour créer cette classe.|

### buildAttempt

 `public function buildAttempt(createAttemptAction: Labs.Core.IAction): Components.ChoiceComponentAttempt`

Crée une instance **ChoiceComponentAttempt** et implémente la méthode abstraite définie dans la classe de base.

 **Paramètres**


|Paramètre|Description|
|:-----|:-----|
| _createAttemptResult_|Résultat de l’action Tentative de création.|
