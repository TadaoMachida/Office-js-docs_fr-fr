
# Labs.Components.DynamicComponentInstance

 _**S’applique à :** applications pour Office |Compléments Office | Office Mix | PowerPoint_

Instance d’un composant dynamique.

```
class DynamicComponentInstance extends Labs.ComponentInstanceBase
```


## Propriétés


|Propriété|Description|
|:-----|:-----|
| `public var component: Components.IDynamicComponentInstance`|Définition de l’instance du composant.|

## Méthodes




### constructeur

 `function constructor(component: Components.IDynamicComponentInstance)`

Crée une instance de composant dynamique à l’aide de la définition [Labs.Components.IDynamicComponentInstance](../../reference/office-mix/labs.components.idynamiccomponentinstance.md).


### getComponents

 `public function getComponents(callback: Labs.Core.ILabCallback<Labs.ComponentInstanceBase[]>): void`

Récupère tous les composants créés par ce composant dynamique.

 **Paramètres**


|Paramètre|Description|
|:-----|:-----|
| _callback_|Fonction de rappel qui se déclenche une fois tous les composants récupérés.|

### createComponent

 `public function createComponent(component: Labs.Core.IComponent, callback: Labs.Core.ILabCallback<Labs.ComponentInstanceBase>): void`

Crée un composant en utilisant le composant dynamique comme composant de base.

 **Paramètres**


|Paramètre|Description|
|:-----|:-----|
| _composant_|Composant ([Labs.Core.IComponent](../../reference/office-mix/labs.core.icomponent.md)) à utiliser pour créer l’instance.|
| _callback_|Fonction de rappel qui se déclenche une fois le composant créé.|

### fermer

 `public function close(callback: Labs.Core.ILabCallback<void>): void`

Indique qu’il n’y aura pas d’autres envois associés à cette instance du composant.

 **Paramètres**


|Paramètre|Description|
|:-----|:-----|
| _callback_|Fonction de rappel qui se déclenche une fois l’instance fermée.|

### isClosed

 `public function isClosed(callback: Labs.Core.ILabCallback<boolean>): void`

Indique si le composant dynamique est fermé. Renvoie **True** si le composant dynamique est fermé.

