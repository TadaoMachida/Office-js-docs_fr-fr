
# Labs.Core.Actions.ISubmitAnswerResult

 _**S’applique à :** apps for Office | Office Add-ins | Office Mix | PowerPoint_

Résultat de l’envoi d’une réponse pour une tentative.

```
interface ISubmitAnswerResult extends Core.IActionResult
```


## Propriétés


|||
|:-----|:-----|
| `submissionId: string`|ID associé à l’envoi. Fourni par le serveur.|
| `complete: boolean`|Renvoie  **true** si la tentative est terminée grâce à l’envoi en cours.|
| `score: any`|Informations sur la note associée à l’envoi.|
