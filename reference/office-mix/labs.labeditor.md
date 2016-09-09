
# Labs.LabEditor

 _**S’applique à :** applications pour Office |Compléments Office | Office Mix | PowerPoint_

L’objet **LabEditor** vous permet de modifier un atelier donné, ainsi que d’obtenir et de définir des données de configuration associées à l’atelier.

```
class LabEditor
```


## Méthodes


### getConfiguration

 `public function getConfiguration(callback: Labs.Core.ILabCallback<Labs.Core.IConfiguration>): void`

Récupère la configuration de l’atelier en cours.

 **Paramètres**


|**Name**|**Description**|
|:-----|:-----|
| _callback_|Fonction de rappel déclenchée une fois la configuration récupérée.|

### setConfiguration

 `public function getConfiguration(callback: Labs.Core.ILabCallback<Labs.Core.IConfiguration>): void`

Définit une nouvelle configuration pour l’atelier.

 **Paramètres**


|**Name**|**Description**|
|:-----|:-----|
| _configuration_|Configuration à définir.|
| _callback_|Fonction de rappel déclenchée une fois la configuration définie.|

### fait

 `public function done(callback: Labs.Core.ILabCallback<void>): void`

Indique que l’utilisateur a fini de modifier l’atelier.

 **Paramètres**


|**Name**|**Description**|
|:-----|:-----|
| _callback_|Fonction de rappel qui se déclenche quand l’éditeur de l’atelier a terminé.|
