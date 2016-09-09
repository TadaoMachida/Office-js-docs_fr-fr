
# Labs.LabInstance

 _**S’applique à :** applications pour Office |Compléments Office | Office Mix | PowerPoint_

Instance d’un atelier configurée pour l’utilisateur actuel. Cet objet permet d’enregistrer et de récupérer des données relatives à l’atelier pour l’utilisateur.

```
class LabInstance
```


## Variables


|||
|:-----|:-----|
| `public var data: any`|Variable de conteneur pour conserver les données utilisateur.|
| `public var components: Labs.ComponentInstanceBase[]`|Composants qui constituent l’instance de l’atelier.|

## Méthodes




### getState

 `public function getState(callback: Labs.Core.ILabCallback<any>): void`

Récupère l’état actuel de l’atelier pour un utilisateur donné.

 **Paramètres**


|||
|:-----|:-----|
| _callback_|Fonction de rappel qui se déclenche une fois l’état de l’atelier récupéré.|

### setState

 `public function setState(state: any, callback: Labs.Core.ILabCallback<void>): void`

Définit l’état de l’atelier pour un utilisateur donné.

 **Paramètres**


|||
|:-----|:-----|
| _state_|État à définir.|
| _callback_|Fonction de rappel qui se déclenche une fois l’état défini.|

### Terminé

 `public function done(callback: Labs.Core.ILabCallback<void>): void`

Fonction indiquant que l’utilisateur a terminé d’utiliser l’atelier.

 **Paramètres**


|||
|:-----|:-----|
| _callback_|Fonction de rappel qui se déclenche une fois l’atelier terminé.|
