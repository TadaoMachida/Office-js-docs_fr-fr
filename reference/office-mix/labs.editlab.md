
# Labs.editLab

 _**S’applique à :** apps for Office | Office Add-ins | Office Mix | PowerPoint_

Ouvre l’atelier spécifié pour le modifier. Il est possible d’indiquer des données de configuration de l’atelier en mode Édition. Toutefois, il est impossible de modifier un atelier lors de son exécution.

```
function editLab(callback: Core.ILabCallback<LabEditor>): void
```


## Paramètres


|**Name**|**Description**|
|:-----|:-----|
| _callback_|Méthode de rappel qui se déclenche une fois l’objet [Labs.LabInstance](../../reference/office-mix/labs.labinstance.md) créé.|
