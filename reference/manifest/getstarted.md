# Élément GetStarted

Fournit des informations utilisées par la légende qui s’affiche lorsque le complément est installé dans des hôtes Word, Excel, PowerPoint et OneNote. L’élément **GetStarted** est un élément enfant de [FormFactor](./formfactor.md).

## Éléments enfants

| Élément                       | Obligatoire | Description                                        |
|:------------------------------|:--------:|:---------------------------------------------------|
| [Titre](#titre)               | Oui      | Définit l’emplacement où se trouvent les fonctionnalités d’un complément     |
| [Description](#description)   | Oui      | URL pointant vers un fichier qui contient les fonctions JavaScript.|
| [LearnMoreUrl](#learnmoreurl) | Non       | URL vers une page qui décrit le complément de façon plus détaillée.   |


## Titre 
Obligatoire. Le titre est utilisé pour la partie supérieure de la légende. L’attribut **resid** fait référence à un ID valide de l’élément [ShortStrings](./resources.md#shortstrings) dans la section [Resources](./resources.md).

## Description
Obligatoire. Description/Contenu du corps de la légende. L’attribut **resid** fait référence à un ID valide de l’élément [LongStrings](./resources.md#longstrings) dans la section [Resources](./resources.md).

## LearnMoreUrl
Obligatoire. URL vers une page dans laquelle l’utilisateur peut obtenir des informations sur votre complément. L’attribut **resid** fait référence à un ID valide de l’élément [Urls](./resources.md#urls) dans la section [Resources](./resources.md).

> **REMARQUE :** **LearnMoreUrl** n’est pas actuellement restitué dans les clients Word, Excel ou PowerPoint. Nous vous recommandons d’ajouter cette URL pour tous les clients afin que l’URL soit restituée lorsqu’elle est disponible. 
