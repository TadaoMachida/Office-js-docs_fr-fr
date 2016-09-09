
# Compléments Outlook contextuels

Les compléments contextuels sont des compléments Outlook qui s’activent en fonction du texte d’un message ou d’un rendez-vous. Grâce aux compléments contextuels, vous pouvez initier des tâches associées à un message sans avoir à quitter le message lui-même. L’expérience utilisateur en est ainsi facilitée et enrichie.

Les compléments contextuels sont différents des compléments relatifs aux pièces jointes ou propres à certains types de messages. Voici des exemples d’applications contextuelles :


- Choisir une adresse ouvre une carte de l’emplacement.
    
- Choisir une chaîne ouvre un complément qui suggère une réunion.
    
- Choisir un numéro de téléphone permet de l’ajouter à vos contacts.
    
Actuellement, les compléments contextuels sont limités à Outlook Web App.

## Création d’un complément contextuel

Pour créer un complément contextuel, le manifeste du complément doit spécifier l’entité ou l’expression régulière qui peut l’activer. L’entité peut être l’une des propriétés de l’objet [Entities](../../reference/outlook/simple-types.md). Par conséquent, le manifeste du complément doit contenir une règle de type  **ItemHasKnownEntity** ou **ItemHasRegularExpressionMatch**. L’exemple suivant montre comment spécifier un numéro de téléphone en tant qu’entité :


```XML
<Rule xsi:type="ItemHasKnownEntity" EntityType="PhoneNumber"/>

```

Lorsqu’un complément contextuel est associé à un compte, il démarre automatiquement quand l’utilisateur clique sur une entité en surbrillance ou une expression régulière. Pour plus d’informations sur les expressions régulières des compléments Outlook, voir [Utiliser des règles d’activation d’expression régulière pour afficher un complément Outlook](../outlook/use-regular-expressions-to-show-an-outlook-add-in.md).

Il existe plusieurs restrictions sur les compléments contextuels :


- Un complément contextuel ne peut exister que dans des compléments de lecture (pas dans des compléments de composition).
    
- Vous ne pouvez pas spécifier la couleur de l’entité en surbrillance.
    
- Si une entité n’est pas en surbrillance, elle ne lancera pas de complément contextuel dans une carte.
    
- La carte mesure 140 à 450 pixels (limite recommandée : 300 pixels) de hauteur et de 570 pixels de largeur.
    
- Vous ne pouvez pas spécifier si le complément s’affichera dans la carte ou dans la barre de complément.
    

## Lancement d’un complément contextuel

Le lancement d’un complément contextuel se fait par le biais de texte (soit une entité connue, soit l’expression régulière d’un développeur) ou par la barre du complément. En règle générale, l’utilisateur identifie le complément contextuel, car l’entité est en surbrillance. L’exemple suivant illustre la mise en surbrillance dans un message. Dans cette image, l’entité (une adresse) est de couleur bleue et est soulignée par une ligne bleue en pointillés. Pour lancer le complément contextuel, l’utilisateur clique sur l’entité en surbrillance. 


**Exemple de texte avec l’entité (une adresse) en surbrillance**

![Indique l’entité en surbrillance dans un paragraphe](../../images/828175bb-4579-4454-abbd-1987fffe5052.jpg)

Bien que la mise en surbrillance soit la meilleure indication des compléments contextuels, dans certains cas, le complément contextuel s’affiche dans la barre du complément :

- Si l’entité est une URL ou une adresse de messagerie.
    
- Si le manifeste du complément a une règle dont le type et l’une des propriétés sont les suivants : type="ItemHasRegularExpressionMatch" et PropertyName="BodyAsHTML" ou PropertyName="SenderSMTPAddress".
    
- Lorsque le manifeste du complément contient une règle d’activation qui utilise une collection de règles OU où la première règle est de type = « ItemIs » avec itemType = « Rendez-vous » ou « Message » et où la deuxième règle est de type = « ItemHasKnownEntity » ou « ItemHasRegularExpressionMatch »
    
- Si la complexité du corps du message électronique a une incidence sur le client de messagerie.
    
Lorsque plusieurs entités ou compléments contextuels sont présents dans un message, l’interaction avec l’utilisateur a lieu comme suit :



- S’il existe plusieurs entités, l’utilisateur doit cliquer sur une autre entité pour lancer le complément pour celle-ci.
    
- Si une entité active plusieurs compléments, chacun d’entre eux s’ouvre dans un nouvel onglet. L’utilisateur bascule entre les onglets, comme dans la barre du complément, pour changer de complément. Par exemple, un nom et une adresse peuvent déclencher un complément téléphonique et un complément de carte géographique.
    
- Si une chaîne unique contient plusieurs entités qui activent plusieurs compléments, la chaîne entière est mise en surbrillance et lorsque l’utilisateur clique sur cette chaîne, tous les compléments concernés par la chaîne s’affichent dans des onglets distincts. Par exemple, une chaîne qui décrit une proposition de réunion dans un restaurant peut activer le complément de suggestion de réunion et un complément d’avis sur des restaurants.
    

## Affichage des compléments contextuels

Un complément contextuel activé s’affiche à l’un des deux emplacements suivants :


- Dans la carte, c’est-à-dire une fenêtre séparée, près de l’entité
    
- Dans la barre du complément, c’est-à-dire la ligne entre l’expéditeur et le corps d’un message
    
La carte s’affiche généralement en dessous de l’entité et est centrée, autant que possible, par rapport à l’entité. Si l’espace n’est pas suffisant en dessous de l’entité, la carte est placée au-dessus. La capture d’écran suivante montre l’entité en surbrillance et, en dessous, un complément activé (Bing Cartes) dans une carte.


**Exemple d’un complément affiché dans une carte**

![Présente une application contextuelle dans une carte](../../images/59bcabc2-7cb0-4b9b-bb9f-06089dca9c31.png)

Remarques :

- L’onglet « Bing Cartes » s’affiche sous forme de texte blanc sur fond bleu. Si un nouveau complément est sélectionné, le texte de l’onglet devient bleu et le fond blanc.
    
- Les onglets des autres compléments, le cas échéant, apparaissent dans un onglet à droite de « Bing Cartes » et le texte est bleu sur fond blanc. Lorsque l’utilisateur clique sur un onglet, le texte de cet onglet devient blanc sur fond bleu et le nouveau complément se charge.
    
- Si l’utilisateur clique sur le bouton « + Plus de compléments », il est redirigé vers l’Office Store.
    
- Si le nom du complément est trop long et ne tient pas dans l’espace disponible, il est remplacé par « ... » à gauche du bouton « + Plus de compléments ». Pour voir les compléments dont le nom ne tient pas dans la barre, l’utilisateur peut cliquer sur ces points de suspension et ainsi afficher ces compléments dans une liste déroulante.
    
- Pour fermer la carte et quitter le complément, il suffit de cliquer en dehors de la carte.
    
La capture d’écran suivante montre comment le même complément (dans ce cas, Bing Cartes) apparaît dans la barre si le texte n’a pas pu être mis en surbrillance (par exemple s’il était contenu dans un lien hypertexte).


**Exemple d’une barre de complément et d’un complément dans un iframe**

![Affiche la barre de l’application au-dessus de l’iframe affichant l’application](../../images/4adce8d2-6957-4d80-b365-7a36dc3cef11.jpg)

Remarques :

- Dans cette capture d’écran, la barre du complément affiche le nom du complément qui a été lancé et le bouton « + Plus de compléments » au-dessus de l’iframe. Si d’autres compléments (contextuels ou non) sont lancés à partir de la barre du complément, ils apparaissent également.
    
- L’iframe affiche le complément. Le développeur peut définir la hauteur de l’iframe mais la largeur est une valeur fixe. La hauteur est la même pour le complément lancé à partir de la barre du complément et la carte ; il n’est pas nécessaire que le développeur spécifie deux hauteurs distinctes.
    

## Affichage des compléments contextuels en fonction des appareils

Sur un ordinateur de bureau, un complément contextuel s’affiche généralement dans une carte ; si plusieurs compléments sont présents, ils apparaissent dans des onglets distincts. Sur une tablette, le même complément s’affiche comme s’il était « au verso » et, si plusieurs compléments sont présents, ils apparaissent dans plusieurs onglets. Sur un téléphone, le complément apparaît sous la forme d’une expérience immersive. Dans le cas où plusieurs compléments sont activés sur l’entité, des points de suspension « ... » apparaissent dans le coin supérieur droit pour permettre aux utilisateurs de naviguer entre les différents compléments sur l’entité spécifique.


## Compléments contextuels actuels

Les compléments contextuels suivants sont installés par défaut pour les utilisateurs qui utilisent des compléments Outlook :


- Cartes Bing 
    
- Réunions suggérées
    
De plus, le complément contextuel [Package Tracker](https://store.office.com/package-tracker-WA104162083.aspx?assetid=WA104162083.aspx) est disponible dans l’Office Store.


## Ressources supplémentaires



- [Prise en main des compléments Outlook pour Office 365](https://dev.outlook.com/MailAppsGettingStarted/GetStarted.aspx)
    
- [Utiliser des règles d’activation d’expression régulière pour afficher un complément Outlook](../outlook/use-regular-expressions-to-show-an-outlook-add-in.md)

- [Objet Entities](../../reference/outlook/simple-types.md)
    
