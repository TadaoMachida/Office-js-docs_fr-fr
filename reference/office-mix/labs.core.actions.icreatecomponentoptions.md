
# Labs.Core.Actions.ICreateComponentOptions

 _**S’applique à :** apps for Office | Office Add-ins | Office Mix | PowerPoint_

Crée un composant.

```
interface ICreateComponentOptions extends Core.IActionOptions
```


## Propriétés


|||
|:-----|:-----|
| `componentId: string`|Composant d’appel de l’action Création de composant.|
| `component: Core.IComponent`|Composant [Labs.Core.IComponent](../../reference/office-mix/labs.core.icomponent.md) à créer.|
| `correlationId?: string`|Champ facultatif correspondant à ce composant dans toutes les instances d’un atelier. Permet à l’hôte d’identifier les différentes tentatives du même composant.|
