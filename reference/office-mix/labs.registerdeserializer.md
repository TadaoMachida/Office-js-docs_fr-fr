
# Labs.registerDeserializer

 _**S’applique à :** apps for Office | Office Add-ins | Office Mix | PowerPoint_

Désérialise un objet JSON spécifié en un objet. Seuls les auteurs de composant doivent l’utiliser.

```
function registerDeserializer(type: string, deserialize: (json: Core.ILabObject): any): void
```


## Paramètres


|**Name**|**Description**|
|:-----|:-----|
|json|Instance [Labs.Core.ILabObject](../../reference/office-mix/labs.core.ilabobject.md) à désérialiser.|

## Valeur renvoyée

Renvoie une instance [Labs.Core.ILabObject](../../reference/office-mix/labs.core.ilabobject.md).

