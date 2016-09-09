
# Labs.Components.ActivityComponentInstance

 _**S’applique à :** apps for Office | Office Add-ins | Office Mix | PowerPoint_

Instance actuelle d’un composant d’activité.

```
class ActivityComponentInstance extends Labs.ComponentInstance<Components.ActivityComponentAttempt>
```


## Propriétés


|**Nom**|**Description**|
|:-----|:-----|
| `public var component: Components.IActivityComponentInstance`|Instance du composant d’activité [Labs.Components.IActivityComponentInstance](../../reference/office-mix/labs.components.iactivitycomponentinstance.md) sous-jacent représenté par cette classe.|

## Méthodes




### constructeur

 `function constructor(component: Components.IActivityComponentInstance)`

Crée une instance de la classe [Labs.Components.IActivityComponentInstance](../../reference/office-mix/labs.components.iactivitycomponentinstance.md).

 **Paramètres**


|**Name**|**Description**|
|:-----|:-----|
| _composant_|Instance de composant  **IActivityComponentInstance** permettant de créer cette classe à partir de cette classe.|

### buildAttempt

 `public function buildAttempt(createAttemptAction: Labs.Core.IAction): Components.ActivityComponentAttempt`

Génère une instance **ActivityComponentAttempt** et implémente la méthode abstraite définie dans la classe de base.

 **Paramètres**


|**Name**|**Description**|
|:-----|:-----|
| _createAttemptResult_|Résultat d’une tentative de création.|
