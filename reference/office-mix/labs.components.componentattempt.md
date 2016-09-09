
# Labs.Components.ComponentAttempt

 _**S’applique à :** applications pour Office |Compléments Office | Office Mix | PowerPoint_

Classe de base pour essayer des composants.

```
class ComponentAttempt
```


## Propriétés


|**Nom**|**Description**|
|:-----|:-----|
| `public var _componentId: string`|ID du composant spécifié.|
| `public var _id: string`|ID de l’atelier associé.|
| `public var _labs: Labs.LabsInternal`|Objet de l’atelier ([Labs.LabsInternal](http://msdn.microsoft.com/library/599fb2c4-bb16-4422-84ad-10ed85a14018.aspx)) utilisé pour interagir avec l’instance [Labs.Core.ILabHost](../../reference/office-mix/labs.core.ilabhost.md) sous-jacente.|
| `public var _resumed: boolean`|Indique **True** si l’atelier a repris l’avancement de la tentative donnée.|
| `public var _state: Labs.ProblemState`|État actuel de la tentative indiqué par l’énumération [Labs.ProblemState](../../reference/office-mix/labs.problemstate.md).|
| `public var _values: { [type:string]: Labs.ValueHolder<any>[]}`|Valeurs éventuelles associées à la tentative, telle qu’elles figurent dans l’objet [Labs.ValueHolder](../../reference/office-mix/labs.valueholder.md).|

## Méthodes




### constructeur

 `(labs: Labs.LabsInternal, componentId: string, attemptId: string, values: {[type:string]: Labs.Core.IValueInstance[]})`

Crée une instance de la classe ComponentAttempt et fournit les valeurs des paramètres d’entrée.

 **Paramètres**


|**Name**|**Description**|
|:-----|:-----|
| _labs_|Instance [Labs.LabsInternal](http://msdn.microsoft.com/library/599fb2c4-bb16-4422-84ad-10ed85a14018.aspx) à utiliser avec la tentative.|
| _attemptId_|ID associé à la tentative.|
| _values_|Tableau de valeurs ([Labs.Core.IValueInstance](../../reference/office-mix/labs.core.ivalueinstance.md)) associé à la tentative.|

### isResumed

 `public function isResumed(): boolean`

Fonction booléenne indiquant si l’atelier a repris.  Indique **True** si l’atelier a repris.

 **Paramètres**

Aucun.


### resume

 `public function resume(callback: Labs.Core.ILabCallback<void>): void`

Indique si l’atelier a repris l’avancement de la tentative donnée et s’il charge les données existantes dans le cadre de ce processus. Une tentative doit être relancée avant de pouvoir être utilisée.

 **Paramètres**


|**Name**|**Description**|
|:-----|:-----|
| _callback_|Fonction de rappel qui se déclenche lors de la reprise de la tentative.|

### getState

 `public function getState(): Labs.ProblemState`

Récupère l’état de l’atelier.

 **Paramètres**

Aucun.


### processAction

 `public function processAction(action: Labs.Core.IAction): void`

Exécute l’action associée à la tentative.

 **Paramètres**

Aucun.


### getValues

 `public function getValues(key: string): Labs.ValueHolder<any>[]`

Récupère les valeurs associées à la tentative.

 **Paramètres**


|**Name**|**Description**|
|:-----|:-----|
| _Key_|Clé associée à la valeur dans le mappage des propriétés.|
