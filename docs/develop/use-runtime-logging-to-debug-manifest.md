# Utilisation de la journalisation runtime pour déboguer le manifeste pour votre complément Office

Vous pouvez utiliser la journalisation runtime pour déboguer le manifeste de votre complément. Cette fonctionnalité peut vous aider à identifier et à résoudre les problèmes avec votre manifeste qui ne sont pas détectés par la validation de schéma XSD, comme une incompatibilité entre les ID de ressources. La journalisation runtime est particulièrement utile pour le débogage des compléments implémentant des commandes de complément.  

>**Remarque :** La fonctionnalité de journalisation runtime est actuellement disponible pour Office 2016 pour ordinateur de bureau.

## Activation de la journalisation runtime

>**Important** : La journalisation runtime affecte les performances. Activez-la uniquement lorsque vous avez besoin de déboguer des problèmes avec votre manifeste de complément.

1. Vérifiez que vous exécutez la version Bureau d’Office 2016 **16.0.7019** ou une version ultérieure. 
2. Ajoutez la clé de registre `RuntimeLogging` sous 'HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Wef\Developer\'. 
3. Définissez la valeur par défaut de la clé pour le chemin d’accès complet du fichier dans lequel écrire le journal. Pour obtenir un exemple, voir [EnableRuntimeLogging.zip](RuntimeLogging/EnableRuntimeLogging.zip). 

 > **Remarque :** Le répertoire dans lequel le fichier journal sera écrit doit déjà exister et vous devez disposer des autorisations d’écriture correspondantes. 
 
L’image suivante indique à quoi doit ressembler le registre.
![Capture d’écran de l’Éditeur du registre avec une clé de registre RuntimeLogging](http://i.imgur.com/Sa9TyI6.png)

Pour désactiver la fonctionnalité, supprimez la clé de registre `RuntimeLogging`. 

## Résolution des problèmes avec votre manifeste

Pour utiliser la journalisation runtime pour résoudre les problèmes de chargement d’un complément, procédez comme suit :
 
1. [Chargez une version test de votre complément](../testing/sideload-office-add-ins-for-testing.md). 

    >Remarque : Nous vous recommandons de charger uniquement une version test du complément que vous testez pour réduire le nombre de messages dans le fichier journal.
2. Si rien ne se produit et que votre complément n’apparaît pas (et ne s’affiche pas dans la boîte de dialogue des compléments), ouvrez le fichier journal.
3. Recherchez le fichier journal pour l’ID de votre complément, que vous définissez dans votre manifeste. Dans le fichier journal, cet ID est intitulé `SolutionId`. 

Dans l’exemple suivant, le fichier journal identifie un contrôle qui pointe vers un fichier de ressources qui n’existe pas. Pour cet exemple, la correction consistera à corriger la faute de frappe dans le manifeste ou à ajouter la ressource manquante.

![Capture d’écran d’un fichier journal avec une entrée qui spécifie un ID de ressource qui est introuvable](http://i.imgur.com/f8bouLA.png) 

##Problèmes connus avec la journalisation runtime
Vous pouvez afficher des messages dans le fichier journal qui sont source de confusion ou classés de façon incorrecte. Par exemple :

- Le message `Medium   Current host not in add-in's host list` suivi de `Unexpected Parsed manifest targeting different host` est classé incorrectement en tant qu’erreur.
- Si vous voyez le message `Unexpected    Add-in is missing required manifest fields  DisplayName` et qu’il ne contient pas de SolutionId, l’erreur n’est probablement pas liée au complément que vous déboguez. 
- Tous les messages `Monitorable` sont des erreurs attendues du point de vue du système. Parfois, ils indiquent un problème avec votre manifeste, comme un élément mal orthographié qui a été ignoré, mais n’a pas provoqué l’échec du manifeste. 

##Ressources supplémentaires

- [Chargement de version test des compléments Office](../testing/sideload-office-add-ins-for-testing.md)
- [Débogage des compléments Office](../testing/debug-add-ins-using-f12-developer-tools-on-windows-10.md)
