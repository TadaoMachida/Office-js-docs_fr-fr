# Élément FunctionFile

Spécifie le fichier de code source pour les opérations qu’un complément expose via les commandes de complément qui exécutent une fonction JavaScript au lieu d’afficher l’interface utilisateur. L’élément **FunctionFile** est un élément enfant de [FormFactor](./formfactor). L’attribut **resid** de l’élément **FunctionFile** est défini sur la valeur de l’attribut **id** d’un élément **Url** dans l’élément **Resources** contenant l’URL d’un fichier HTML qui contient ou charge toutes les fonctions JavaScript utilisées par les boutons de commande de complément sans interface utilisateur, telles que définies par l’[élément Control](control.md).

Vous trouverez ci-dessous un exemple de l’élément **FunctionFile**.


```XML
<DesktopFormFactor>
          <FunctionFile resid="residDesktopFuncUrl" />
          <ExtensionPoint xsi:type="PrimaryCommandSurface">
            <!-- information about this extension point -->
          </ExtensionPoint>

          <!-- You can define more than one ExtensionPoint element as needed -->

        </DesktopFormFactor>
```

Le code JavaScript dans le fichier HTML indiqué par l’élément **FunctionFile** doit appeler `Office.initialize` et définir les fonctions nommées qui prennent un paramètre unique : `event`. Les fonctions doivent utiliser l’API [item.notificationMessages](../../../reference/outlook/Office.context.mailbox.item.md) pour indiquer l’avancement, la réussite ou l’échec à l’utilisateur. Il doit appeler également [event.completed](../../../reference/shared/event.completed.md) lorsque l’exécution est terminée. Le nom de ces fonctions est utilisé dans l’élément **FunctionName** pour les boutons sans interface utilisateur.

Vous trouverez ci-dessous un exemple d’un fichier HTML définissant une fonction **trackMessage**.

```js
Office.intialize = function () {
    doAuth();
}

function trackMessage (event) {
    var buttonId = event.source.id;    
    var itemId = Office.context.mailbox.item.id;
    // save this message
    event.completed();
}
```

Le code suivant montre comment implémenter la fonction utilisée par **FunctionName**.




```js
        // The initialize function must be run each time a new page is loaded.
        (function () {
            Office.initialize = function (reason) {
               // If you need to initialize something you can do so here.
            };
        })();

            // Your function must be in the global namespace.
        function writeText(event) {

            // Implement your custom code here. The following code is a simple example.

            Office.context.document.setSelectedDataAsync("ExecuteFunction works. Button ID=" + event.source.id,
                function (asyncResult) {
                    var error = asyncResult.error;
                    if (asyncResult.status === "failed") {
                        // Show error message.
                    }
                    else {
                        // Show success message.
                    }
                });
           // Calling event.completed is required. event.completed lets the platform know that processing has completed.
       event.completed();
        }
```


 >**Important**  L’appel de l’élément **event.completed** indique que vous avez correctement géré l’événement. Lorsqu’une fonction est appelée à plusieurs reprises, par exemple lorsque l’utilisateur clique plusieurs fois sur une même commande de complément, tous les événements sont automatiquement mis en file d’attente. Le premier événement s’exécute automatiquement, tandis que les autres événements restent dans la file d’attente. Lorsque votre fonction appelle **event.completed**, l’appel suivant de cette fonction s’exécute. Vous devez implémenter **event.completed** pour que votre fonction s’exécute correctement.
