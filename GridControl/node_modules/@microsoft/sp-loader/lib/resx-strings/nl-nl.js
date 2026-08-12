define([], function() {
  var strings = {
    "_nQNACBeQ34aV6bVwtFBayA": {
      "loaderUserFriendlyError": "De toepassing kan niet worden geladen op deze pagina. Gebruik de knop Terug van de browser om het opnieuw te proberen. Stuur de beheerder van de site de informatie in Technische details als het probleem zich blijft voordoen.",
      "platformFailedToLoadError": "Laden van platform is mislukt. Id: {0}, naam: {1}",
      "platformFailedToLoadWithMessageError": "Laden van platform is mislukt. Id: {0}, naam: {1}.\r\nFout: {2}",
      "applicationFailedToInitializeError": "Er is een fout opgetreden bij het initialiseren van de toepassing. Fout: {0}",
      "invalidPreloadedDataError": "Ongeldige vooraf geladen gegevens.",
      "manifestNotFoundByIdError": "Manifest voor onderdeel-id {0} is niet gevonden.",
      "manifestNotFoundError": "Manifest is niet gevonden voor onderdeel-id {0} en versie {1}.",
      "systemConfigDisabledError": "System.config() wordt niet ondersteund. Gebruik een manifest om de configuratie op te geven.",
      "loadComponentLog": "Onderdeel {0} ({1}) laden.",
      "loadComponentEndLog": "Onderdeel {0} ({1}) is geladen.",
      "loadComponentRetryLog": "Onderdeel {0} ({1}) laden. Poging {2} van {3}.",
      "loadComponentError": "Laden van onderdeel {0} ({1}) is mislukt.\r\nOorspronkelijke fout: {2}",
      "loadComponentMaxRetriesError": "Er is {2} keer tevergeefs geprobeerd om onderdeel {0} ({1}) te laden.",
      "loadComponentDependencyError": "Laden van onderdeelafhankelijkheid {0} van onderdeel {1} ({2}) is mislukt.\r\nOorspronkelijke fout: {3}",
      "loadComponentDependencyFailoverPathError": "Laden van onderdeelafhankelijkheid {0} met failover-pad {1} van onderdeel {2} ({3}) is mislukt.\r\nOorspronkelijke fout: {4}",
      "loadPathDependencyLog": "Laden van padafhankelijkheid {0} van onderdeel {1} ({2}) is mislukt.",
      "loadPathDependencyError": "Laden van padafhankelijkheid {0} van onderdeel {1} ({2}) is mislukt.\r\nOorspronkelijke fout: {3}",
      "loadPathDependencyBlockedByAnotherDependencyError": "Laden van padafhankelijkheid {0} van onderdeel {1} ({2}) is mislukt omdat het laden van een andere afhankelijkheid eveneens is mislukt.",
      "loadEntryPointError": "Laden van invoerpunt van onderdeel {0} ({1}) is mislukt.\r\nOorspronkelijke fout: {2}",
      "loadComponentReturnsEmptyError": "loadComponent() heeft een leeg object geretourneerd voor onderdeel {0} ({1}).",
      "loadComponentReturnsDefaultEmptyError": "loadComponent() heeft een object geretourneerd met een lege standaardeigenschap voor onderdeel {0} ({1}).",
      "moduleHasUndeclaredDependencyError": "Het invoerpunt voor onderdeel {0} ({1}) heeft een afhankelijkheid van {2} die niet is gedeclareerd in het manifest.",
      "loadScriptWithStringError": "Voor de functie loadScript is geen tekenreeks als 2e parameter toegestaan. Gebruik in plaats daarvan ILoadScriptOptions.",
      "tooManyManifestsError": "Er zijn {0} manifesten (versies {1}) gevonden voor onderdeel {2}.",
      "tooManyCompatibleVersionsError": "Er zijn {0} compatibele versies ({1}) gevonden voor onderdeel {2} versie {3}.",
      "tooManyComponentsError": "Er zijn te veel onderdelen gevonden voor id {0}.",
      "noComponentFoundError": "Geen onderdeel gevonden voor id {0}.",
      "deleteComponentLog": "Onderdeel {0} versie {1} verwijderen uit het archief.",
      "browserNotSupportedError": "Deze versie van de browser wordt niet ondersteund.\r\nWerk uw browser bij naar de nieuwste versie.",
      "ie9OrOlderNotSupportedError": "Deze pagina biedt geen ondersteuning voor Internet Explorer-versies die ouder zijn dan versie 10. Werk de webbrowser bij.",
      "firefox43OrOlderNotSupportedError": "Deze pagina biedt geen ondersteuning voor Mozilla Firefox-versies die ouder zijn dan versie 44. Werk de webbrowser bij.",
      "resourceNotFoundError": "Resource {0} is niet gevonden in de laderconfiguratie van het manifest voor onderdeel {1} ({2}).",
      "noFailoverPathError": "Kan resolveAddress() niet aanroepen voor een onderdeel dat geen failover-pad heeft",
      "urlStatusLocalhostFileNotFoundError": "Laden van URL {3} voor resource {2} in onderdeel {0} ({1}) is mislukt. Het bestand is niet gevonden op de server.\r\nControleer of gulp serve wordt uitgevoerd.",
      "urlStatusFileNotFoundError": "Laden van URL {3} voor resource {2} in onderdeel {0} ({1}) is mislukt. Het bestand is niet gevonden op de server.",
      "urlStatusForbiddenError": "Laden van URL {3} voor resource {2} in onderdeel {0} ({1}) is mislukt. Toegang tot het bestand is niet toegestaan.",
      "urlStatusClientErrorError": "Laden van URL {3} voor resource {2} in onderdeel {0} ({1}) is mislukt. Er is een fout opgetreden bij het aanvragen van het bestand.",
      "urlStatusServerErrorError": "Laden van URL {3} voor resource {2} in onderdeel {0} ({1}) is mislukt. Er doet zich een probleem voor op de server.",
      "urlStatusLocalhostNetworkErrorError": "Laden van URL {3} voor resource {2} in onderdeel {0} ({1}) is mislukt. Er is een probleem opgetreden met het netwerk.\r\nControleer of gulp serve en gulp trust-dev-cert worden uitgevoerd.",
      "urlStatusHttpsNetworkErrorError": "Laden van URL {3} voor resource {2} in onderdeel {0} ({1}) is mislukt. Er doet zich een probleem voor met het netwerk.\r\nMogelijk gaat het hier om een probleem met een HTTPS-certificaat. Controleer of u over het juiste certificaat beschikt.",
      "urlStatusNetworkErrorError": "Laden van URL {3} voor resource {2} in onderdeel {0} ({1}) is mislukt. Er doet zich een probleem voor met het netwerk.",
      "urlStatusUndefinedError": "Laden van URL {3} voor resource {2} in onderdeel {0} ({1}) is mislukt vanwege onbekende problemen.",
      "isUndefinedValidateError": "De waarde voor \"{0}\" mag niet onbepaald zijn",
      "failedToCreateGlobalVariableError": "Het maken van globale variabele {0} vanuit script {1} is mislukt",
      "dependencyLoadError": "Laden van module {0} is mislukt omdat afhankelijkheid {1} niet is geladen",
      "missingPathDependencyError": "Ontbrekende padafhankelijkheid {0} uit onderdeel {1} ({2}). Bestaande padafhankelijkheden: {3}",
      "listSeparator": ", "
    },
    "_FmFyAWZ1md7Z1R+V8t2S2Q": {
      "errorLoadingDebugScriptHTTPS": "Fout bij het laden van script voor foutopsporing. Controleer of de server actief is en de URL voor de parameter {0} juist is.",
      "errorLoadingDebugScriptHTTP": "Fout bij het laden van script voor foutopsporing. Controleer of de server actief is, de URL voor de parameter {0} juist is en of het laden van onveilige scripts is toegestaan. Overweeg ook om een ontwikkelingscertificaat te gebruiken en scripts voor foutopsporing via HTTPS te verwerken.",
      "errorLoadingDebugScriptMalformed": "Fout bij het laden van script voor foutopsporing. De foutopsporings-URL ({0}) is waarschijnlijk onjuist.",
      "errorLoadingDebugScriptUnknown": "Onbekende fout bij het laden van een script voor foutopsporing.",
      "errorLoadingDebugLoaderTitle": "Fout bij het laden van lader voor foutopsporing.",
      "errorLoadingDebugManifestTitle": "Fout bij het laden van foutopsporingsmanifesten.",
      "errorLoadingUnknownTitle": "Fout bij het laden van scripts voor foutopsporing."
    },
    "_RPELcTeq3ZByqi3N5dt18w": {
      "missingDeveloperToolsTabInitFunctionError": "Onderdeel of initialisatiefunctie ontbreekt.",
      "closeDeveloperToolsAriaLabel": "Sluit de ontwikkelhulpprogramma's."
    },
    "_fwMQe6Xe08yEeCPNxngd+g": {
      "warningHeading": "Waarschuwing!",
      "warningLine1": "Als u dit hulpprogramma gebruikt, wordt u blootgesteld aan potentiële beveiligingsrisico's die ertoe kunnen leiden dat anderen toegang krijgen tot uw persoonlijke Office 365-gegevens (documenten, e-mailberichten, gesprekken en meer). Zorg ervoor dat u de persoon of organisatie vertrouwt die toegang tot dit hulpprogramma vraagt, voordat u doorgaat.",
      "warningLine2": "Meer informatie: {0}"
    },
    "_mraBnnuq2J9WjrAcnw9QNA": {
      "debugManifestErrorDetail": "Er is een fout opgetreden bij het laden van foutopsporingsmanifesten.",
      "debugManifestErrorDismissButtonText": "Verwijderen"
    },
    "_upo3vfLFBbnbzl2hKy2TwA": {
      "allowDebugManifestsTitle": "Scripts voor foutopsporing toestaan?",
      "allowDebugLoaderTitle": "Lader voor foutopsporing toestaan?",
      "allowDebugLoaderAndManifestsTitle": "Lader voor foutopsporing en scripts voor foutopsporing toestaan?",
      "debugManifestLoadingWarning": "WAARSCHUWING: Deze pagina bevat onveilige scripts die de computer kunnen beschadigen indien ze worden geladen. Ga niet door tenzij u de ontwikkelaar vertrouwt en de risico's begrijpt.",
      "debugManifestLoadingWarning2": "Klik op {0} als u het niet zeker weet.",
      "debugManifestLoadingConfirm": "Scripts voor foutopsporing laden",
      "debugManifestLoadingCancel": "Scripts voor foutopsporing niet laden",
      "debugManifestLoadingCalloutText": "Klik hier als u niet weet wat u moet doen."
    },
    "_SxImp5ewsUToxeAHBkB+pw": {
      "developerToolsTabLoadingText": "Laden...",
      "developerToolsTabLoadingUnknownError": "Er is een onbekende fout opgetreden bij het laden van de module ontwikkelhulpprogramma's."
    },
    "_sovI4qDAUPMnD4jg3Vsyfg": {
      "tabTitle": "Manifesten",
      "noManifestSelected": "Er is geen manifest geselecteerd"
    },
    "_g7G0QHJ5bQYlxe+lk+DcxA": {
      "TabTitle": "Prestaties",
      "ErrorAccessingPerfDataErrorMessage": "Kan de prestatiegegevens niet ophalen: object is null of is niet gedefinieerd.",
      "ErrorAccessingRedirectDataErrorMessage": "Er is een probleem opgetreden bij het openen van de prestatiegegevens van de HTTP-omleiding.",
      "ErrorParsingPercievedLatencyErrorMessage": "Er is een fout aangetroffen bij het parseren van de waargenomen latentiegegevens.",
      "ErrorParsingApiDataErrorMessage": "Er is een fout aangetroffen bij het parseren van de API-gegevens.",
      "UnkownPerformanceDataErrorMessage": "Er is een onbekende fout opgetreden: {0}",
      "DefaultWebPartName": "Webonderdeel",
      "ServerResponseLabel": "Reactietijd van server",
      "ApplicationInitializationLabel": "Initialisatietijd voor toepassing",
      "ScriptFetchEvalLabel": "Script ophalen en evalueren",
      "SpLoaderStartLabel": "SPFx initialiseren",
      "PageRenderLabel": "Weergavetijd voor pagina",
      "LeftNavRenderLabel": "Weergavetijd voor linkernavigatiegedeelte",
      "CanvasRenderLabel": "Weergavetijd voor canvas",
      "LayoutRenderLabel": "Weergavetijd voor indeling",
      "RedirectResponseLabel": "Laadtijd voor omleiding van antwoord",
      "AppLoadLabel": "Laadtijd voor toepassing",
      "RenderWebPartsLabel": "Weergavetijd voor webonderdelen",
      "TotalRenderTimeLabel": "Totaal",
      "GeneralErrorMessage": "Er is iets misgegaan bij het ophalen van de prestatiegegevens.",
      "ErrorMessagePrefix": "Foutbericht: {0}",
      "PerformanceDataHint": "Opmerking: vernieuw de pagina na het toevoegen of verwijderen van een webonderdeel om de bijgewerkte prestatiegegevens te zien.",
      "ModulesLoadedLegendLabel": "Geladen modules",
      "InitializationLegendLabel": "Initialisatie",
      "RenderTimeLegendLabel": "Weergavetijd",
      "InitializationTimeLabel": "Initialisatietijd",
      "ModuleLoadingTimeLabel": "Laadtijd voor modules",
      "ModuleLazyLoadingDelayLabel": "Laden van module is vertraagd",
      "DataFetchTimeLabel": "Tijd voor ophalen van gegevens",
      "DataFetchLegendLabel": "Ophalen van gegevens",
      "ItemsColumnHeader": "Items",
      "DurationColumnHeader": "Duur",
      "MillisecondsUnitLabel": "{0} ms",
      "NAPlaceholder": "N.v.t."
    },
    "_gqinlPQb8HZprTeCpwNz2w": {
      "TabTitle": "Tracering",
      "EmptyTraceData": "Geen traceringen geladen.",
      "ExportCSVButtonLabel": "Exporteren als CSV",
      "LevelHeaderLabel": "Niveau",
      "MessageHeaderLabel": "Bericht",
      "ScopeHeaderLabel": "Bereik",
      "SourceHeaderLabel": "Bron",
      "TimestampHeaderLabel": "Timestamp",
      "TimestampFormat": "{2}-{1}-{0} {3}:{4}:{5},{6}",
      "ErrorAccessingTraceDataErrorMessage": "Er is een probleem opgetreden bij het openen van de traceergegevens."
    }
  };

  strings.default = strings;
  return strings;
});