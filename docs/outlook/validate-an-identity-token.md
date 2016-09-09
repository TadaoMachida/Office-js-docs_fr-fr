
# Valider un jeton d’identité Exchange

Votre complément Outlook peut vous envoyer un jeton d’identité, mais avant d’approuver la demande, vous devez valider le jeton pour garantir qu’il provient du serveur Exchange attendu. Les exemples de cet article montrent comment valider le jeton d’identité Exchange en utilisant un objet de validation écrit en C# ; cependant, vous pouvez utiliser n’importe quel langage de programmation pour effectuer la validation. Les opérations requises pour valider le jeton sont décrites dans le [Document préliminaire Internet JWT (JSON Web Token)](http://self-issued.info/docs/draft-goland-json-web-token-00.mdl). 

Nous vous suggérons d’utiliser un processus en quatre étapes pour valider le jeton d’identité et obtenir l’identificateur unique de l’utilisateur. Dans un premier temps, extrayez le jeton JWT (JSON Web Token) à partir d’une chaîne d’URL encodée au format base64. Dans un deuxième temps, assurez-vous que le jeton est bien formé, c’est-à-dire qu’il est adapté à votre complément Outlook, qu’il n’a pas expiré et que vous pouvez extraire une URL valide pour le document de métadonnées d’authentification. Dans un troisième temps, récupérez le document de métadonnées d’authentification sur le serveur Exchange et validez la signature jointe au jeton d’identité. Dans un quatrième temps, calculez un identificateur unique pour l’utilisateur en hachant l’ID Exchange de l’utilisateur avec l’URL du document de métadonnées d’authentification. Le processus peut sembler globalement complexe, mais chaque étape individuelle est relativement simple. Vous pouvez télécharger la solution contenant ces exemples sur le web à partir de la page  [Outlook-Add-in-JavaScript-ValidateIdentityToken](https://github.com/OfficeDev/Outlook-Add-in-JavaScript-ValidateIdentityToken).
 




## Configuration pour valider votre jeton d’identité


Les exemples de code dans cet article dépendent de Windows Identity Foundation (WIF), ainsi que d’une DLL qui prolonge le WIF avec des gestionnaires de jetons JSON. Vous pouvez télécharger les assemblys requis sur les sites suivants :


- [Windows Identity Foundation](http://msdn.microsoft.com/en-us/security/aa570351)
    
- [Windows.IdentityModel.Extensions.dll pour les applications 32 bits](http://download.microsoft.com/download/0/1/D/01D06854-CA0C-46F1-ADBA-EBF86010DCC6/MicrosoftIdentityExtensions-32.msi)
    
- [Windows.IdentityModel.Extensions.dll pour les applications 64 bits](http://download.microsoft.com/download/0/1/D/01D06854-CA0C-46F1-ADBA-EBF86010DCC6/MicrosoftIdentityExtensions-64.msi)
    

## Extraction du jeton Web JSON


La méthode de fabrique  **Decode** fractionne le JWT du serveur Exchange en trois chaînes qui composent le jeton, puis utilise la méthode **Base64Decode** (présentée dans le deuxième exemple) pour décoder l’en-tête et la charge JWT en chaînes JSON. Les chaînes sont passées au constructeur **JsonToken**, où le contenu du JWT est validé et une nouvelle instance de l’objet **JsonToken** est renvoyée.


```C#
    public static JsonToken Decode(string rawToken)
    {
      string[] tokenParts = rawToken.Split('.');

      if (tokenParts.Length != 3)
      {
        throw new ApplicationException("Token must have three parts separated by '.' characters.");
      }

      string encodedHeader = tokenParts[0];
      string encodedPayload = tokenParts[1];
      string signature = tokenParts[2];

      string decodedHeader = Base64Decode(encodedHeader);
      string decodedPayload = Base64Decode(encodedPayload);

      JavaScriptSerializer serializer = new JavaScriptSerializer();

      Dictionary<string, string> header = serializer.Deserialize<Dictionary<string, string>>(decodedHeader);
      Dictionary<string, string> payload = serializer.Deserialize<Dictionary<string, string>>(decodedPayload);

      return new JsonToken(header, payload, signature);
    }
```

La méthode **Base64Decode** implémente la logique de décodage qui est décrite dans l’Annexe « Remarques sur l’implémentation du codage base64url sans remplissage » dans le [Document préliminaire Internet sur JWT (JSON Web Token)](http://self-issued.info/docs/draft-goland-json-web-token-00.mdl).




```C#
    public static Encoding TextEncoding = Encoding.UTF8;

    private static char Base64PadCharacter = '=';
    private static char Base64Character62 = '+';
    private static char Base64Character63 = '/';
    private static char Base64UrlCharacter62 = '-';
    private static char Base64UrlCharacter63 = '_';

    private static byte[] DecodeBytes(string arg)
    {
      if (String.IsNullOrEmpty(arg))
      {
        throw new ApplicationException("String to decode cannot be null or empty.");
      }

      StringBuilder s = new StringBuilder(arg);
      s.Replace(Base64UrlCharacter62, Base64Character62);
      s.Replace(Base64UrlCharacter63, Base64Character63);

      int pad = s.Length % 4;
      s.Append(Base64PadCharacter, (pad == 0) ? 0 : 4 - pad);

      return Convert.FromBase64String(s.ToString());
    }

    private static string Base64Decode(string arg)
    {
      return TextEncoding.GetString(DecodeBytes(arg));
    }
```


## Analyse du JWT


Le constructeur de l’objet  **JsonToken** vérifie la structure et le contenu du JWT pour déterminer s’il est valide. Il convient de le faire avant que vous demandiez le document de métadonnées d’authentification. Si le JWT ne contient pas les revendications appropriées, ou s’il est périmé, vous pouvez éviter un appel au serveur Exchange et le retard associé.

Le constructeur appelle des méthodes utilitaires pour déterminer si les différentes revendications sont présentes et comprises dans l’étendue. En cas de problème, la méthode utilitaire déclenche une exception d’application. Si aucune exception n’est déclenchée, la propriété  **IsValid** est définie à **true** et le jeton est prêt pour la validation de signature.

Les diverses méthodes utilitaires sont décrites plus loin dans cet article.




```C#
    public JsonToken(Dictionary<string, string> header, Dictionary<string, string> payload, string signature)
    {

      // Assume that the token is invalid to start out.
      this.IsValid = false;

      // Set the private dictionaries that contain the claims.
      this.headerClaims = header;
      this.payloadClaims = payload;
      this.signature = signature;

      // If there is no "appctx" claim in the token, throw an ApplicationException.
      if (!this.payloadClaims.ContainsKey(AuthClaimTypes.AppContext))
      {
        throw new ApplicationException(String.Format("The {0} claim is not present.", AuthClaimTypes.AppContext));
      }

      appContext = new JavaScriptSerializer().Deserialize<Dictionary<string, string>>(payload[AuthClaimTypes.AppContext]);


      // Validate the header fields.
      this.ValidateHeader();

      // Determine whether the token is within its valid time.
      this.ValidateLifetime();

      // Validate that the token was sent to the correct URL.
      this.ValidateAudience();

      // Validate the token version.
      this.ValidateVersion();

      // Make sure that the appctx contains an authentication
      // metadata location.
      this.ValidateMetadataLocation();

      // If the token passes all the validation checks, we
      // can assume that it is valid.
      this.IsValid = true;
    }
```


### Méthode ValidateHeader

La méthode  **ValidateHeader** vérifie que les revendications requises figurent bien dans l’en-tête du jeton et que les revendications ont les valeurs appropriées. L’en-tête doit correspondre à celui présenté ci-dessous ; sinon, la méthode déclenche une exception et se termine.

```js
{ "typ" : "JWT", "alg" : "RS256", "x5t" : "<thumbprint>" }
```

```C#
    private void ValidateHeaderClaim(string key, string value)
    {
      if (!this.headerClaims.ContainsKey(key))
      {
        throw new ApplicationException(String.Format("Header does not contain \"{0}\" claim.", key));
      }

      if (!value.Equals(this.headerClaims[key]))
      {
        throw new ApplicationException(String.Format("\"{0}\" claim must be \"{0}\".", key, value));
      }
    }

    private void ValidateHeader()
    {
      ValidateHeaderClaim(AuthClaimTypes.TokenType, Config.TokenType);
      ValidateHeaderClaim(AuthClaimTypes.Algorithm, Config.Algorithm);
    
      if (!this.headerClaims.ContainsKey(AuthClaimTypes.x509Thumprint))
      {
        throw new ApplicationException(String.Format("Header does not contain \"{0}\" claim.", AuthClaimTypes.x509Thumprint));
      }
    }


```


### Méthode ValidateLifetime

Deux dates sont fournies dans le jeton JWT : « nbf » (« not before ») indique la date et l’heure marquant le début de validité du jeton, et « exp » indique l’heure d’expiration du jeton. Seuls les jetons présentés entre ces deux dates doivent être considérés valides. Pour tenir compte des différences mineures au niveau des paramètres d’horloge entre le serveur et le client, cette méthode valide les jetons jusqu’à 5 minutes avant et 5 minutes après les heures figurant dans le jeton.


```C#
    private void ValidateLifetime()
    {
      if (!this.payloadClaims.ContainsKey(AuthClaimTypes.ValidFrom))
      {
        throw new ApplicationException(
          String.Format("The \"{0}\" claim is missing from the token.", AuthClaimTypes.ValidFrom));
      }

      if (!this.payloadClaims.ContainsKey(AuthClaimTypes.ValidTo))
      {
        throw new ApplicationException(
          String.Format("The \"{0}\" claim is missing from the token.", AuthClaimTypes.ValidTo));
      }

      DateTime unixEpoch = new DateTime(1970, 1, 1, 0, 0, 0,DateTimeKind.Utc);

      TimeSpan padding = new TimeSpan(0, 5, 0);

      DateTime validFrom = unixEpoch.AddSeconds(int.Parse(this.payloadClaims[AuthClaimTypes.ValidFrom]));
      DateTime validTo = unixEpoch.AddSeconds(int.Parse(this.payloadClaims[AuthClaimTypes.ValidTo]));

      DateTime now = DateTime.UtcNow;

      if (now < (validFrom - padding))
      {
        throw new ApplicationException(String.Format("The token is not valid until {0}.", validFrom));
      }

      if (now > (validTo + padding))
      {
        throw new ApplicationException(String.Format("The token is not valid after {0}.", validFrom));
      }
    }
```

Les dates  **validFrom** (« nbf ») et **validTo** (« exp ») sont envoyées comme un nombre de secondes depuis Epoch Unix, 1er janvier 1970. Les dates et heures sont calculées à l’aide d’UTC pour éviter tout problème lié à des différences de fuseau horaire entre le serveur Exchange et le serveur exécutant le code de validation.


### Méthode ValidateAudience

Le jeton d’identité est uniquement valide pour le complément qui l’a demandé. La méthode  **ValidateAudience** vérifie la revendication d’audience dans le jeton pour s’assurer qu’elle correspond à l’URL attendue pour le complément Outlook.


```C#
    private void ValidateAudience()
    {
      if (!this.payloadClaims.ContainsKey(AuthClaimTypes.Audience))
      {
        throw new ApplicationException(String.Format("The \"{0}\" claim is missing from the application context.", AuthClaimTypes.Audience));
      }

      string location = Config.Audience.Replace("/", "-").Replace("\\", "-");
      string audience = this.payloadClaims[AuthClaimTypes.Audience].Replace("/", "-").Replace("\\", "-");

      if (!location.Equals(audience))
      {
        throw new ApplicationException(String.Format(
          "The audience URL does not match. Expected {0}; got {1}.",
          Config.Audience, this.payloadClaims[AuthClaimTypes.Audience]));
      }
    }

```


### Méthode ValidateVersion

La méthode  **ValidateVersion** vérifie la version du jeton d’identité et s’assure qu’elle est bien la version attendue. Différentes versions du jeton peuvent apporter différentes revendications. La vérification de la version garantit que les revendications attendues seront contenues dans le jeton d’identité.


```js
    private void ValidateVersion()
    {
      if (!this.appContext.ContainsKey(AuthClaimTypes.MsExchExtensionVersion))
      {
        throw new ApplicationException(String.Format("The \"{0}\" claim is missing from the token.", AuthClaimTypes.MsExchExtensionVersion));
      }

      if (!Config.Version.Equals(this.appContext[AuthClaimTypes.MsExchExtensionVersion]))
      {
        throw new ApplicationException(String.Format(
          "The version does not match. Expected {0}; got {1}.",
          Config.Version, this.appContext[AuthClaimTypes.MsExchExtensionVersion]));
      }
    }

```


### Méthode ValidateMetadataLocation

L’objet de métadonnées d’authentification qui est stocké dans le serveur Exchange contient les informations qui sont requises pour valider la signature incluse dans le jeton d’identité. La méthode  **ValidateMetadataLocation** s’assure de la présence d’une revendication URL de métadonnées d’authentification dans le jeton d’identité, garantissant ainsi que la validation de la signature est effectuée à l’étape suivante.


```C#
    private void ValidateMetadataLocation()
    {
      if (!this.appContext.ContainsKey(AuthClaimTypes.MsExchAuthMetadataUrl))
      {
        throw new ApplicationException(String.Format("The \"{0}\" claim is missing from the token.", AuthClaimTypes.MsExchAuthMetadataUrl));
      }
    }

```


## Validation de la signature du jeton d’identité


Une fois que vous savez que le JWT contient les revendications dont vous avez besoin pour valider la signature, vous pouvez utiliser WIF (Windows Identity Foundation) et les extensions WIF pour valider la signature sur le jeton. Vous avez besoin des informations suivantes pour valider la signature :


- La chaîne de jeton d’identité codée en URL base-64 envoyée par le serveur Exchange.
    
- L’emplacement du document de métadonnées d’authentification provenant du JWT.
    
- L’URL d’audience provenant du JWT.
    
Dans cet exemple, le constructeur pour un objet  **IdentityToken** obtient le document de métadonnées d’authentification du serveur Exchange et valide la signature sur le jeton d’identité. Si le jeton d’identité est valide, vous pouvez utiliser l’instance de l’objet **IdentityToken** pour obtenir l’identifiant utilisateur unique inclus dans le jeton d’identité.




```C#
    public IdentityToken(string rawToken, string audience, string authMetadataEndpoint)
    {
      X509Certificate2 currentCertificate = null;

      currentCertificate = AuthMetadata.GetSigningCertificate(new Uri(authMetadataEndpoint));

      JsonWebSecurityTokenHandler jsonTokenHandler =
          GetSecurityTokenHandler(audience, authMetadataEndpoint, currentCertificate);

      SecurityToken jsonToken = jsonTokenHandler.ReadToken(rawToken);
      JsonWebSecurityToken webToken = (JsonWebSecurityToken)jsonToken;

      SigningCertificateThumbprint = currentCertificate.Thumbprint;
      Issuer = webToken.Issuer;
      Audience = webToken.Audience;
      ValidTo = webToken.ValidTo;
      ValidFrom = webToken.ValidFrom;
      foreach (JsonWebTokenClaim claim in webToken.Claims)
      {
        if (claim.ClaimType.Equals(AuthClaimTypes.AppContextSender))
        {
          ApplicationContextSender = claim.Value;
        }

        if (claim.ClaimType.Equals(AuthClaimTypes.IsBrowserHostedApp))
        {
          IsBrowserHostedApp = claim.Value == "true";
        }

        if (claim.ClaimType.Equals(AuthClaimTypes.AppContext))
        {
          string[] appContextClaims = claim.Value.Split(',');
          Dictionary<string, string> appContext =
              new JavaScriptSerializer().Deserialize<Dictionary<string, string>>(claim.Value);
          AuthenticationMetaDataUrl = appContext[AuthClaimTypes.MsExchAuthMetadataUrl];
          ExchangeID = appContext[AuthClaimTypes.MsExchImmutableId];
          TokenVersion = appContext[AuthClaimTypes.MsExchTokenVersion];
        }
      }
    }


```

La majeure partie du code dans le constructeur de l’objet  **IdentityToken** définit les propriétés sur l’instance avec les revendications provenant du serveur Exchange. Le constructeur appelle la méthode **GetSecurityTokenHandler** pour obtenir un gestionnaire de jetons qui valide le jeton d’identité Exchange. La méthode **GetSecurityTokenHandler** appelle deux méthodes utilitaires, **GetMetadataDocument** et **GetSigningCertificate**, qui se chargent d’obtenir le certificat de signature à partir du serveur Exchange. Ces méthodes sont décrites dans les sections suivantes.


### Méthode GetSecurityTokenHandler

La méthode  **GetSecurityTokenHandler** renvoie un gestionnaire de jetons WIF qui valide le jeton d’identité. La majeure partie du code dans la méthode initialise le gestionnaire de jetons pour effectuer la validation ; cependant, cette méthode appelle la méthode **GetSigningCertificate** pour récupérer le certificat X.509 utilisé pour signer le jeton provenant du serveur Exchange.


```C#
    private JsonWebSecurityTokenHandler GetSecurityTokenHandler(string audience,
        string authMetadataEndpoint,
        X509Certificate2 currentCertificate)
    {
      JsonWebSecurityTokenHandler jsonTokenHandler = new JsonWebSecurityTokenHandler();
      jsonTokenHandler.Configuration = new SecurityTokenHandlerConfiguration();

      jsonTokenHandler.Configuration.AudienceRestriction = new AudienceRestriction(AudienceUriMode.Always);
      jsonTokenHandler.Configuration.AudienceRestriction.AllowedAudienceUris.Add(
        new Uri(audience, UriKind.RelativeOrAbsolute));

      jsonTokenHandler.Configuration.CertificateValidator = X509CertificateValidator.None;

      jsonTokenHandler.Configuration.IssuerTokenResolver =
        SecurityTokenResolver.CreateDefaultSecurityTokenResolver(
          new ReadOnlyCollection<SecurityToken>(new List<SecurityToken>(
            new SecurityToken[]
            {
              new X509SecurityToken(currentCertificate)
            })), false);

      ConfigurationBasedIssuerNameRegistry issuerNameRegistry = new ConfigurationBasedIssuerNameRegistry();
      issuerNameRegistry.AddTrustedIssuer(currentCertificate.Thumbprint, Config.ExchangeApplicationIdentifier);
      jsonTokenHandler.Configuration.IssuerNameRegistry = issuerNameRegistry;

      return jsonTokenHandler;
    }
```


### Méthode GetSigningCertificate

La méthode  **GetSigningCertificate** appelle la méthode **GetMetadataDocument** pour récupérer les métadonnées d’authentification du serveur Exchange, puis renvoie le premier certificat X.509 dans le document de métadonnées d’authentification. Si le document n’existe pas, la méthode génère une exception d’application.


```C#
    private X509Certificate2 GetSigningCertificate(Uri authMetadataEndpoint)
    {
      JsonAuthMetadataDocument document = GetMetadataDocument(authMetadataEndpoint);

      if (null != document.keys &amp;&amp; document.keys.Length > 0)
      {
        JsonKey signingKey = document.keys[0];

        if (null != signingKey &amp;&amp; null != signingKey.keyValue)
        {
          return new X509Certificate2(Encoding.UTF8.GetBytes(signingKey.keyValue.value));
        }
      }

      throw new ApplicationException("The metadata document does not contain a signing certificate.");
    }

```


### Méthode GetMetadataDocument

Le document de métadonnées d’authentification contient les informations dont vous avez besoin pour valider la signature sur le jeton d’identité Exchange. Le document est envoyé sous la forme d’une chaîne JSON. La méthode  **GetMetatDataDocument** demande le document à l’emplacement spécifié dans le jeton d’identité Exchange et renvoie un objet qui encapsule la chaîne JSON sous la forme d’un objet. Si l’URL ne contient pas de document de métadonnées d’authentification, la méthode génère une exception d’application.


```C#
    private JsonAuthMetadataDocument GetMetadataDocument(Uri authMetadataEndpoint)
    {
      // Uncomment the next line if your Exchange server uses the default
      // self-signed certificate.
      // ServicePointManager.ServerCertificateValidationCallback = Config.CertificateValidationCallback;

      byte[] acsMetadata;
      using (WebClient webClient = new WebClient())
      {
        acsMetadata = webClient.DownloadData(authMetadataEndpoint);
      }
      string jsonResponseString = Encoding.UTF8.GetString(acsMetadata);

      JsonAuthMetadataDocument document = new JavaScriptSerializer().Deserialize<JsonAuthMetadataDocument>(jsonResponseString);

      if (null == document)
      {
        throw new ApplicationException(String.Format("No authentication metadata document found at {0}.", authMetadataEndpoint));
      }

      return document;
    }
```

Par défaut, le serveur utilise un certificat X.509 autosigné pour authentifier les demandes du document de métadonnées d’authentification. Sauf si vous avez installé un certificat qui retrace un serveur racine, vous devez créer une méthode de rappel de validation de certificat, sinon la demande du document de métadonnées d’authentification échoue. 

La classe  **ServicePointManager** dans l’espace de noms .NET Framework System.Net vous permet de joindre une méthode de rappel de validation en définissant la propriété **ServerCertificateValidationCallback**. Vous pouvez voir un exemple de méthode de rappel de validation de certificat adaptée au développement et au test dans l’article [Validation de certificats X509](http://msdn.microsoft.com/en-us/library/dd633677%28EXCHG.80%29.aspx).


 **Remarque de sécurité**  Si vous utilisez une méthode de rappel de validation de certificat, vous devez vous assurer qu’elle répond aux exigences de sécurité de votre organisation.


## Calculer l’ID unique d’un compte Exchange


Vous pouvez créer un identificateur unique pour un compte Exchange en hachant l’URL du document de métadonnées d’authentification avec l’identificateur Exchange du compte. Lorsque vous avez cet identificateur unique, vous pouvez l’utiliser pour créer un système d’authentification unique destiné à votre service web de complément Outlook. Pour plus d’informations sur l’utilisation de l’identificateur unique pour l’authentification unique, voir [Authentifier un utilisateur avec un jeton d’identité pour Exchange](../outlook/authenticate-a-user-with-an-identity-token.md).

La propriété  **UniqueUserIdentification** crée un hachage SHA256 basé sur une valeur salt de l’ID Exchange et l’URL de métadonnées d’authentification via un fournisseur SHA256 standard à partir de l’espace de noms **System.Security.Cryptography**.


 **Remarque de sécurité**  Vous devez hacher le document de métadonnées d’authentification avec l’ID Exchange pour créer l’identificateur unique d’un compte. L’utilisation de l’ID Exchange uniquement peut exposer votre service aux utilisateurs non autorisés. Et comme toujours, quand il s’agit d’authentification et de sécurité, vous devez vous assurer que l’utilisation de l’identificateur unique créé à l’aide de cette méthode répond aux exigences de sécurité de votre application.




```C#
    // Salt to apply when creating unique ID.
    private byte[] Salt = new byte[] {<Provide random salt bytes here };

    private string ComputeUniqueIdentification()
    {
      byte[] inputBytes = Encoding.ASCII.GetBytes(string.Concat(ExchangeID, AuthenticationMetaDataUrl));

      // Combine input bytes and salt.
      byte[] saltedInput = new byte[Salt.Length + inputBytes.Length];
      Salt.CopyTo(saltedInput, 0);
      inputBytes.CopyTo(saltedInput, Salt.Length);

      // Compute the unique key.
      byte[] hashedBytes = SHA256CryptoServiceProvider.Create().ComputeHash(saltedInput);

      // Convert the hashed value to a string and return.
      return BitConverter.ToString(hashedBytes);
    }

    public string UniqueUserIdentification
    {
      get { return ComputeUniqueIdentification(); }
    }


```


## Objets utilitaires


Les exemples de code contenus dans cet article dépendent d’objets utilitaires qui fournissent des noms conviviaux aux constantes utilisées. Le tableau suivant répertorie ces objets utilitaires.


**Tableau 1 : Objets utilitaires**


|**Objet**|**Description**|
|:-----|:-----|
|**AuthClaimsType**|Collecte en un emplacement unique les identificateurs de revendications qui sont utilisés par le code de validation de jeton.|
|**Config**|Fournit les constantes pour valider le jeton d’identité. |
|**JsonAuthMetadataDocument**|Encapsule le document de métadonnées d’authentification JSON envoyé par le serveur Exchange.|

### Objet AuthClaimTypes

L’objet  **AuthClaimTypes** collecte en un emplacement unique les identificateurs de revendications qui sont utilisés par le code de validation de jeton. Il inclut des revendications JWT standard mais aussi les revendications spécifiques contenues dans le jeton d’identité Exchange.


```C#
  public class AuthClaimTypes
  {
    public const string NameIdentifier =
        JsonWebTokenConstants.ReservedClaims.NameIdentifier;
    public const string MsExchImmutableId = "msexchuid";
    public const string MsExchTokenVersion = "version";
    public const string MsExchAuthMetadataUrl = "amurl";

    public const string AppContext =
        JsonWebTokenConstants.ReservedClaims.AppContext;
    public const string Audience =
        JsonWebTokenConstants.ReservedClaims.Audience;
    public const string Issuer =
        JsonWebTokenConstants.ReservedClaims.Issuer;
    public const string ValidFrom =
        JsonWebTokenConstants.ReservedClaims.NotBefore;
    public const string ValidTo =
        JsonWebTokenConstants.ReservedClaims.ExpiresOn;

    public const string AppContextSender = "appctxsender";
    public const string IsBrowserHostedApp = "isbrowserhostedapp";

    public const string TokenType = "typ";
    public const string Algorithm = "alg";
    public const string x509Thumbprint = "x5t";      
  }
```


### Objet Config

L’objet  **Config** contient les constantes qui sont utilisées pour valider le jeton d’identité, ainsi qu’une méthode de rappel de validation de certificat que vous pouvez utiliser si votre serveur n’a pas de certificat X509 qui retrace un certificat racine.


 
  **Remarque de sécurité**  La méthode de rappel de certificat de sécurité est uniquement requise si votre serveur utilise le certificat autosigné par défaut. La méthode de rappel dans cet exemple renvoie  **false** lorsque le certificat est autosigné, vous devrez donc le remplacer par une méthode de rappel répondant aux exigences de sécurité de votre organisation. Pour un exemple d’une méthode de rappel de validation de certificat adaptée au développement et au test, voir [Validation de certificats X509](http://msdn.microsoft.com/en-us/library/dd633677%28EXCHG.80%29.aspx).


```C#
  public static class Config
  {
    public static string Algorithm = "RS256";
    public static string Audience = @"https:\\localhost:44300\Pages\IdentityTest.html";
    public static string TokenType = "JWT";
    public static string Version = "ExIdTok.V1";

    public static string ExchangeApplicationIdentifier = "Exchange";

    internal static bool CertificateValidationCallback(
    object sender,
    System.Security.Cryptography.X509Certificates.X509Certificate certificate,
    System.Security.Cryptography.X509Certificates.X509Chain chain,
    System.Net.Security.SslPolicyErrors sslPolicyErrors)
    {
      // If the certificate is a valid, signed certificate, return true.
      if (sslPolicyErrors == System.Net.Security.SslPolicyErrors.None)
      {
        return true;
      }

      // If there are errors in the certificate chain, look at each error to determine the cause.
      else
      {
        return false;
      }
    }
  }
```


### Objet JsonAuthMetadataDocument

L’objet  **JsonAuthMetadataDocument** expose le contenu du document de métadonnées d’authentification au moyen de propriétés.


```C#
using System;

namespace IdentityTest
{
  public class JsonAuthMetadataDocument
  {
    public string id { get; set; }
    public string version { get; set; }
    public string name { get; set; }
    public string realm { get; set; }
    public string serviceName { get; set; }
    public string issuer { get; set; }
    public string [] allowedAudiences { get; set; }
    public JsonKey[] keys;
    public JsonEndpoint[] endpoints;
  }

  public class JsonEndpoint
  {
    public string location { get; set; }
    public string protocol { get; set; }
    public string usage { get; set; }
  }

  public class JsonKey
  {
    public string usage { get; set; }
    public JsonKeyValue keyValue { get; set; }
  }

  public class JsonKeyValue
  {
    public string type { get; set; }
    public string value { get; set; }
  }
}

```


## Ressources supplémentaires



- [Authentifier un complément Outlook à l’aide de jetons d’identité Exchange](../outlook/authentication.md)
    
- [Présentation du jeton d’identité Exchange](../outlook/inside-the-identity-token.md)
    
