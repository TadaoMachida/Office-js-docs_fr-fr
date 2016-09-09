
# Labs.Core.IConnectionResponse

 _**S’applique à :** apps for Office | Office Add-ins | Office Mix | PowerPoint_

Informations de la réponse renvoyées par l’appel de connexion.

```
interface IConnectionResponse
```


## Propriétés


|||
|:-----|:-----|
| `initializationInfo: Core.IConfigurationInfo`|Informations sur la configuration de l’initialisation ou **null** si l’application n’a pas été initialisée.|
| `mode: Core.LabMode`|Mode utilisé pour exécuter l’atelier.|
| `hostVersion: Core.IVersion`|Informations sur la version ([Labs.Core.IVersion](../../reference/office-mix/labs.core.iversion.md)) du serveur.|
| `userInfo: Core.IUserInfo`|Informations de l’utilisateur ([Labs.Core.IUserInfo](../../reference/office-mix/labs.core.iuserinfo.md)).|
