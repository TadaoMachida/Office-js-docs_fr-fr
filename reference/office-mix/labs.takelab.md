
# Labs.takeLab

 _**S’applique à :** apps for Office | Office Add-ins | Office Mix | PowerPoint_

Exécute l’atelier spécifié et active l’envoi de résultats de l’atelier au serveur. Un atelier ne peut pas être exécuté lorsqu’il est en cours de modification.

```
function takeLab(callback: Core.ILabCallback<LabInstance>): void
```


## Paramètres


|**Name**|**Description**|
|:-----|:-----|
| _callback_|Méthode de rappel déclenchée une fois l’objet [Labs.LabInstance](../../reference/office-mix/labs.labinstance.md) créé.|
