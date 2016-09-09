
# Labs.Core.IValueInstance

 _**S’applique à :** apps for Office | Office Add-ins | Office Mix | PowerPoint_

Instance d’objet [Labs.Core.IValue](../../reference/office-mix/labs.core.ivalue.md) qui contient les éventuelles données de la valeur.

```
interface IValueInstance
```


## Propriétés


|||
|:-----|:-----|
| `valueId: string`|ID de la valeur représentée par cette instance.|
| `isHint: boolean`|Expression booléenne **true** si cette valeur est un conseil.|
| `hasValue: boolean`|Expression booléenne  **true** si les informations de l’instance contiennent la valeur.|
| `value?: any`|Valeur. Ce paramètre peut être défini ou non s’il a été masqué ou non.|
