
# Commandes de complément pour Excel, Word et PowerPoint

Les commandes de complément sont des éléments qui étendent l’interface utilisateur d’Office et qui lancent des actions dans votre complément. Vous pouvez ajouter un bouton sur le ruban ou un élément à un menu contextuel. Lorsque les utilisateurs sélectionnent une commande de complément, ils lancent des actions telles que l’exécution de code JavaScript ou l’affichage d’une page de complément dans un volet Office. Les commandes de complément permettent aux utilisateurs de trouver et d’utiliser votre complément, ce qui contribue à augmenter l’adoption et la réutilisation de votre complément, ainsi qu’à améliorer la fidélisation des clients.

Pour en savoir plus sur les fonctionnalités, regardez la vidéo sur les [commandes de complément du ruban Office](https://channel9.msdn.com/events/Build/2016/P551).


**Complément incluant des commandes en cours d’exécution dans Excel (version Bureau)**
![Commandes de complément](../../images/addincommands1.png)

**Complément incluant des commandes en cours d’exécution dans Excel (version Online)**
![Commandes de complément](../../images/addincommands2.png)

## Fonctionnalités de commande
Les fonctionnalités de commande suivantes sont actuellement prises en charge.

**Points d’extension**

- Onglets de ruban - Permet d’étendre les onglets prédéfinis ou de créer un onglet personnalisé.
- Menus contextuels - Permet d’étendre les menus contextuels sélectionnés. 

**Types de contrôles**

- Boutons simples - Permettent de déclencher des actions spécifiques.
- Menus - Contiennent plusieurs boutons qui déclenchent des actions.

**Actions**

- ShowTaskpane - Affiche un ou plusieurs volets où sont chargées des pages HTML personnalisées.
- ExecuteFunction - Charge une page HTML invisible, puis y exécute une fonction JavaScript. Pour afficher l’interface utilisateur au sein de votre fonction (par exemple, erreurs, avancement, entrées supplémentaires), vous pouvez utiliser l’API [displayDialog](http://dev.office.com/reference/add-ins/shared/officeui).  

## Plateformes prises en charge
Les commandes de complément sont actuellement prises en charge sur les plateformes suivantes :

- Office pour Windows 2016 pour ordinateur de bureau (version 16.0.6769.0000 ou ultérieure)
- Office Online avec comptes personnels
- Office Online avec comptes professionnels ou scolaires (Aperçu)

D’autres plateformes seront bientôt disponibles.

## Prise en main des commandes de complément

Pour obtenir des informations sur la façon de spécifier des commandes de complément dans votre manifeste, consultez la page concernant [la définition des commandes de complément dans votre manifeste](http://dev.office.com/docs/add-ins/outlook/manifests/define-add-in-commands).

Pour commencer à utiliser des commandes de complément, consultez la page relative aux [exemples de commandes de complément Office](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/) sur GitHub.





