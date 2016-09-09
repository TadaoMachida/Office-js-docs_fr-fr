
# Labs.Core.IComponent

 _**S’applique à :** apps for Office | Office Add-ins | Office Mix | PowerPoint_

Classe de base pour la représentation des composants d’un laboratoire.

```
interface IComponent extends Core.ILabObject, Core.IUserData
```


## Propriétés


|||
|:-----|:-----|
| `name: string`|Nom du composant.|
| `values: {[type:string]: Core.IValue[]}`|Mappage des propriétés de valeur associées au composant.|
