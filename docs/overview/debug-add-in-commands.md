# Utilisation de la journalisation runtime pour déboguer les commandes de complément

Les clients Office 16 pour ordinateur de bureau disposent d’une nouvelle fonctionnalité permettant de consigner des informations utiles. Entre autres, cet outil peut vous aider à diagnostiquer des erreurs dans votre manifeste de complément, ce qui est particulièrement utile si vous créez des manifestes incluant des commandes de complément. 

La documentation complète concernant la fonctionnalité est en cours de préparation mais, en attendant, découvrez comment vous pouvez l’utiliser pour déboguer les problèmes lors de l’analyse de manifestes incluant des commandes de complément.

##Activation de la journalisation runtime

**Important** : la journalisation runtime possède un **gain de performances**. Activez-la uniquement lorsque vous avez besoin de déboguer des problèmes avec vos compléments.

1. Vérifiez que votre version prend en charge la journalisation runtime. La version des clients **Office 16 pour ordinateur de bureau** doit être supérieure ou égale à **16.0.7019**
2. Ajoutez la clé de registre `RuntimeLogging` sous `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Wef\Developer\`. 
3. Définissez la valeur par défaut de la clé pour le chemin d’accès complet du fichier dans lequel écrire le journal. Voir un [exemple de clé de registre](RuntimeLogging/EnableRuntimeLogging.zip) (non décompressé)

Votre registre doit ressembler à ceci : ![](http://i.imgur.com/Sa9TyI6.png)

Si vous devez désactiver la fonctionnalité, supprimez simplement la clé de registre. 

##Diagnostic des problèmes liés aux commandes
La journalisation runtime est utile pour détecter les **problèmes avec votre manifeste** qui sont difficiles à identifier, par exemple, les incohérences entre les ID de ressource et les longueurs non valides non détectées par la validation de schéma XSD. 

Voici les étapes pour effectuer des essais :
 
1. Suivez les instructions du fichier [LisezMoi](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/blob/master/README.md) afin de charger une version test de votre complément. 
2. Si vous ne voyez pas votre projet de boutons du ruban et que rien ne s’affiche dans la boîte de dialogue des compléments, vérifiez les journaux.
3. Rechercher l’ID de votre complément, que vous définissez dans votre manifeste, pour rechercher des messages appartenant à ce complément. Les journaux signalent cet ID comme `SolutionId`. Il est recommandé que vous chargiez uniquement une version test à ce moment-là pour éviter qu’un trop grand nombre de messages n’appartenant pas à votre complément s’affichent. 

Dans l’exemple ci-dessous, la journalisation runtime a permis d’identifier un contrôle qui pointe vers un fichier de ressources inexistant. La solution consiste à corriger la faute de frappe (le cas échéant) ou à ajouter la ressource manquante.

![](http://i.imgur.com/f8bouLA.png) 

##Problèmes connus relatifs à la journalisation
La journalisation runtime a toujours présenté des bogues. Vous verrez peut-être plusieurs messages pas clairs ou classés de manière inappropriée. Par exemple :

- Les messages `Medium  Current host not in add-in's host list` suivis de `Unexpected Parsed manifest targeting different host` sont classés de manière incorrecte. Il ne s’agit pas d’erreurs, vous pouvez les ignorer.
- Le message `Unexpected   Add-in is missing required manifest fields  DisplayName` ne contient pas l’élément SolutionId du complément posant problème. Toutefois, cela n’est probablement PAS lié au complément que vous déboguez. 
- Tous les messages `Monitorable` sont des erreurs attendues du point de vue du système. Parfois, ils indiquent un problème avec votre manifeste (par exemple, un élément mal orthographié qui a été ignoré, mais n’a pas provoqué l’échec du manifeste). 

