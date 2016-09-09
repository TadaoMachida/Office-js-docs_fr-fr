
# Labs.Command

 _**S’applique à :** apps for Office | Office Add-ins | Office Mix | PowerPoint_

Commande générale permettant de transmettre des messages entre le client et l’hôte.

```
class Command
```


## Propriétés


|**Nom**|**Description**|
|:-----|:-----|
| `public var type: string`|Type de la commande.|
| `public var commandData: any`|Données facultatives associées à la commande.|

## Méthodes




### constructeur

 `function constructor(type: string, commandData?: any)`

Description

 **Paramètres**


|||
|:-----|:-----|
| `type`|Type de la commande.|
| `commandData`|Données facultatives associées à la commande.|
