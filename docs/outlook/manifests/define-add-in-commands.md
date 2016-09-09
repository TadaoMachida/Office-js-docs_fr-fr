# Définir des commandes de complément dans votre manifeste

Les commandes de complément sont un moyen de personnaliser facilement l’interface utilisateur d’Office par défaut en y ajoutant des éléments d’interface qui exécutent des actions, tels que des boutons personnalisés ajoutés au ruban. Pour créer des commandes, ajoutez un nœud **[VersionOverrides](../../../reference/manifest/versionoverrides.md)** à un manifeste existant du volet Office. 

Lorsqu’un manifeste contient l’élément **VersionOverrides**, les versions de Word, Excel, Outlook et PowerPoint prenant en charge les commandes de complément utiliseront les informations de cet élément pour charger le complément. Les versions antérieures des produits Office qui ne prennent pas en charge les commandes de complément ignoreront l’élément.

Lorsque les applications clientes reconnaissent le nœud **VersionOverrides**, le nom du complément s’affiche dans le ruban, et non dans un volet Office ou un volet de lecture/composition. Le complément n’apparaîtra pas dans les deux emplacements.
 

## Nœud VersionOverrides

L’élément [VersionOverrides](../../../reference/manifest/versionoverrides.md) est l’élément racine qui contient des informations pour les commandes de complément implémentées par le complément. Il est pris en charge dans la version 1.1 du schéma de manifeste et les versions ultérieures, mais il est défini dans la version 1.0 du schéma VersionOverrides. 

L’élément VersionOverrides inclut les éléments enfants suivants :

- [Description](../../../reference/manifest/description.md)
- [Configuration requise](../../../reference/manifest/requirements.md)
- [Hôtes](../../../reference/manifest/hosts.md)
- [Ressources](../../../reference/manifest/resources.md)

Le diagramme suivant illustre la hiérarchie des éléments utilisés pour définir des commandes de complément. 

![Hiérarchie des éléments de commandes de complément dans le manifeste](../../../images/080da303-51c4-4882-b74a-7ba11517c0ad.png)

## Modifications des règles pour les commandes de complément Outlook

Les modifications suivantes affectent les règles du manifeste :

- Les règles d’activation sont désormais à l’intérieur de chaque point d’entrée.
    
- L’attribut **ItemIs** de l’élément [Rule](../../../reference/manifest/rule.md) a été modifié. **ItemType** peut correspondre à Message ou à AppointmentAttendee. L’attribut **FormType** a été supprimé.
    
- L’attribut **ItemHasKnownEntity** de l’élément [Rule](../../../reference/manifest/rule.md) a été mis à jour afin d’accepter une chaîne pour EntityType.
    

## Exemple de manifestes

Pour un exemple de manifeste qui implémente les commandes de complément pour Word, Excel et PowerPoint, voir l’article sur l’[exemple de commandes de complément simples](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/tree/master/Simple).

Pour un exemple de manifeste qui implémente des commandes de complément pour Outlook, voir l’article sur l’[exemple de fichier de manifeste pour un complément Outlook](https://gist.github.com/mlafleur/95b7ac030bb7a7ae742527e85a36b095).


## Ressources supplémentaires


- [Commandes de complément pour Outlook](../../outlook/add-in-commands-for-outlook.md)
    
- [Manifestes des compléments Outlook](../../outlook/manifests/manifests.md)
    
- [Démonstration de la commande du complément Outlook](https://github.com/jasonjoh/command-demo)
