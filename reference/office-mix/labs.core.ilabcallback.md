
# Labs.Core.ILabCallback

 _**S’applique à :** apps for Office | Office Add-ins | Office Mix | PowerPoint_

Interface de gestion des méthodes de rappel Labs.js.

```
interface ILabCallback<T>
```


## Signature de rappel

 `(err: any, data: T): void`

 **Paramètres de rappel**


|||
|:-----|:-----|
| _err_|**Null** si aucune erreur ne se produit. Autre réponse que **null** si une erreur s’est produite.|
| _data_|Données renvoyées avec le rappel.|
