
# Élément Host
Spécifie le type d’application hôte Office pris en charge par votre complément Office.

 **Type de complément :** Application de contenu, de volet Office, de messagerie


## Syntaxe :


```XML
<Host Name= ["Document" | "Database" | "Mailbox" | "Presentation" | "Project" | "Workbook"] />
```


## Attributs



|**Attribut**|**Type**|**Requis**|**Description**|
|:-----|:-----|:-----|:-----|
|Nom|string|obligatoire|Nom du type d’application hôte Office.|

## Remarques

Vous pouvez spécifier les valeurs suivantes dans l’attribut **Name** d’un élément **Host**. Chaque valeur correspond à l’ensemble d’une ou plusieurs applications hôtes Office prises en charge par votre complément.



|**Name**|**Applications hôtes Office**|
|:-----|:-----|
| `"Document"`|Word, Word Online, Word sur iPad|
| `"Database"`|applications web Access|
| `"Mailbox"`|Outlook, Outlook Web App, OWA pour les périphériques|
| `"Notebook"`|OneNote Online|
| `"Presentation"`|PowerPoint, PowerPoint Online, PowerPoint sur iPad|
| `"Project"`|Projet|
| `"Workbook"`|Excel, Excel Online, Excel sur iPad|

## Remarques

Pour plus d’informations sur la spécification de prise en charge d’hôtes, voir l’article relatif à la [spécification d’hôtes Office et de conditions requises d’API](../../docs/overview/specify-office-hosts-and-api-requirements.md).

