define([], function() {
  var strings = {
    "_nQNACBeQ34aV6bVwtFBayA": {
      "loaderUserFriendlyError": "Nevar ielādēt lietojumprogrammu šajā lapā. Izmantojiet pārlūkprogrammas pogu pāriešanai atpakaļ, lai mēģinātu vēlreiz. Ja problēma atkārtojas, sazinieties ar vietnes administratoru un sniedziet viņam informāciju, kas redzama sadaļā Tehniskā informācija.",
      "platformFailedToLoadError": "Platformu neizdevās ielādēt. ID: {0}, nosaukums: {1}",
      "platformFailedToLoadWithMessageError": "Platformu neizdevās ielādēt. ID: {0}, nosaukums: {1}.\r\nKļūda: {2}",
      "applicationFailedToInitializeError": "Inicializējot lietojumprogrammu, radās kļūda: {0}",
      "invalidPreloadedDataError": "Nederīgi iepriekš ielādēti dati.",
      "manifestNotFoundByIdError": "Komponentam ID {0} nav atrasts manifests.",
      "manifestNotFoundError": "Manifests komponentam ar ID {0} un versiju {1} nav atrasts.",
      "systemConfigDisabledError": "Metode System.config() netiek atbalstīta. Lai norādītu konfigurāciju, izmantojiet manifestu.",
      "loadComponentLog": "Tiek ielādēts komponents {0} ({1}).",
      "loadComponentEndLog": "Komponents {0} ({1}) ielādēts.",
      "loadComponentRetryLog": "Tiek ielādēts komponents {0} ({1}). {2}. mēģinājums no {3}.",
      "loadComponentError": "Neizdevās ielādēt komponentu {0} ({1}).\r\nSākotnējā kļūda: {2}",
      "loadComponentMaxRetriesError": "{2} reizes bez panākumiem notika mēģinājums ielādēt komponentu {0} ({1}).",
      "loadComponentDependencyError": "Neizdevās ielādēt komponentu atkarību {0} no komponenta {1} ({2}).\r\nSākotnējā kļūda: {3}",
      "loadComponentDependencyFailoverPathError": "Neizdevās ielādēt komponentu atkarību {0} ar kļūmjpārlēces ceļu {1} no komponenta {2} ({3}).\r\nSākotnējā kļūda: {4}",
      "loadPathDependencyLog": "Tiek ielādēta ceļa atkarība {0} no komponenta {1} ({2})",
      "loadPathDependencyError": "Neizdevās ielādēt ceļa atkarību {0} no komponenta {1} ({2}).\r\nSākotnējā kļūda: {3}",
      "loadPathDependencyBlockedByAnotherDependencyError": "Neizdevās ielādēt ceļa atkarību {0} no komponenta {1} ({2}), jo neizdevās ielādēt citu atkarību.",
      "loadEntryPointError": "Neizdevās ielādēt ieejas punktu no komponenta {0} ({1}).\r\nSākotnējā kļūda: {2}",
      "loadComponentReturnsEmptyError": "loadComponent() atgrieza tukšu objektu komponentam {0} ({1}).",
      "loadComponentReturnsDefaultEmptyError": "loadComponent() atgrieza objektu ar tukšu noklusējuma rekvizītu komponentam {0} ({1}).",
      "moduleHasUndeclaredDependencyError": "Komponenta {0} ({1}) ieejas punktam ir atkarība no objekta {2}, kas manifestā nav deklarēts.",
      "loadScriptWithStringError": "Funkcija loadScript neļauj izmantot virkni kā 2. parametru. Tā vietā izmantojiet ILoadScriptOptions.",
      "tooManyManifestsError": "Komponentam {2} atrasti {0} manifesti (versijas {1}).",
      "tooManyCompatibleVersionsError": "Komponenta {2} versijai {3} atrastas {0} saderīgas versijas ({1}).",
      "tooManyComponentsError": "Identifikatoram {0} atrasts pārāk daudz komponentu.",
      "noComponentFoundError": "Nav atrasts neviens komponents ar ID {0}.",
      "deleteComponentLog": "Krātuvē notiek komponenta {0} versijas {1} dzēšana.",
      "browserNotSupportedError": "Šī pārlūkprogrammas versija netiek atbalstīta.\r\nLūdzu, atjauniniet savu pārlūkprogrammu uz jaunāko versiju.",
      "ie9OrOlderNotSupportedError": "Šī lapā neatbalsta Internet Explorer laidienus, kas ir vecāki par versiju 10. Lūdzu, atjauniniet savu tīmekļa pārlūkprogrammu.",
      "firefox43OrOlderNotSupportedError": "Šī lapā neatbalsta Mozilla Firefox laidienus, kas ir vecāki par versiju 44. Lūdzu, atjauniniet savu tīmekļa pārlūkprogrammu.",
      "resourceNotFoundError": "Resurss {0} komponenta {1} ({2}) manifesta ielādētāja konfigurācijā nav atrasts.",
      "noFailoverPathError": "Nevar izsaukt resolveAddress() komponentam bez kļūmjpārlēces ceļa",
      "urlStatusLocalhostFileNotFoundError": "Neizdevās ielādēt vietrādi URL {3} resursam {2} komponentā {0} ({1}). Fails serverī netika atrasts.\r\nPārliecinieties, vai izmantojat 'gulp serve'.",
      "urlStatusFileNotFoundError": "Neizdevās ielādēt vietrādi URL {3} resursam {2} komponentā {0} ({1}). Fails serverī netika atrasts.",
      "urlStatusForbiddenError": "Neizdevās ielādēt vietrādi URL {3} resursam {2} komponentā {0} ({1}). Piekļuve failam ir aizliegta.",
      "urlStatusClientErrorError": "Neizdevās ielādēt vietrādi URL {3} resursam {2} komponentā {0} ({1}). Pieprasot failu, radās kļūda.",
      "urlStatusServerErrorError": "Neizdevās ielādēt vietrādi URL {3} resursam {2} komponentā {0} ({1}). Serverī radās problēma.",
      "urlStatusLocalhostNetworkErrorError": "Neizdevās ielādēt vietrādi URL {3} resursam {2} komponentā {0} ({1}). Radās tīkla problēma.\r\nPārliecinieties, vai palaižat gulp serve and esat palaidis gulp trust-dev-cert.",
      "urlStatusHttpsNetworkErrorError": "Neizdevās ielādēt vietrādi URL {3} resursam {2} komponentā {0} ({1}). Radās tīkla problēma.\r\nTā var būt problēma ar HTTPS sertifikātu. Pārliecinieties, vai jums ir pareizais sertifikāts.",
      "urlStatusNetworkErrorError": "Neizdevās ielādēt vietrādi URL {3} resursam {2} komponentā {0} ({1}). Radās tīkla problēma.",
      "urlStatusUndefinedError": "Nezināmu problēmu dēļ neizdevās ielādēt vietrādi URL {3} resursam {2} komponentā {0} ({1}).",
      "isUndefinedValidateError": "Mainīgā {0} vērtība nedrīkst būt nedefinēta",
      "failedToCreateGlobalVariableError": "No skripta {1} neizdevās izveidot globālo mainīgo {0}",
      "dependencyLoadError": "Neizdevās ielādēt moduli {0}, jo atkarība {1} netika ielādēta",
      "missingPathDependencyError": "Komponentā {1} ({2}) trūkst ceļa atkarības {0}. Esošās ceļa atkarības: {3}",
      "listSeparator": ", "
    },
    "_FmFyAWZ1md7Z1R+V8t2S2Q": {
      "errorLoadingDebugScriptHTTPS": "Ielādējot atkļūdošanas skriptu, radās kļūda. Pārliecinieties, vai serveris darbojas un vai parametra {0} vietrādis URL ir pareizs.",
      "errorLoadingDebugScriptHTTP": "Ielādējot atkļūdošanas skriptu, radās kļūda. Pārliecinieties, vai serveris darbojas, vai parametra {0} vietrādis URL ir pareizs un vai nedrošu skriptu ielāde ir atļauta. Apsveriet arī iespēju izmantot izstrādes sertifikātu un apkalpot atkļūdošanas skriptus, izmantojot HTTPS.",
      "errorLoadingDebugScriptMalformed": "Ielādējot atkļūdošanas skriptu, radās kļūda. Atkļūdošanas vietrādis URL ({0}), šķiet, ir nepareizi veidots.",
      "errorLoadingDebugScriptUnknown": "Ielādējot atkļūdošanas skriptu, radās nezināma kļūda.",
      "errorLoadingDebugLoaderTitle": "Ielādējot atkļūdošanas ielādētāju, radās kļūda.",
      "errorLoadingDebugManifestTitle": "Ielādējot atkļūdošanas manifestus, radās kļūda.",
      "errorLoadingUnknownTitle": "Ielādējot atkļūdošanas skriptus, radās kļūda."
    },
    "_RPELcTeq3ZByqi3N5dt18w": {
      "missingDeveloperToolsTabInitFunctionError": "Trūkst komponenta vai inicializētāja funkcijas.",
      "closeDeveloperToolsAriaLabel": "Aizvērt izstrādātāju rīkus."
    },
    "_fwMQe6Xe08yEeCPNxngd+g": {
      "warningHeading": "Uzmanību!",
      "warningLine1": "Šī rīka lietošana pakļauj jūs iespējamiem drošības apdraudējumiem, kas var izraisīt citu personu piekļuvi jūsu personiskajiem Office 365 datiem (dokumentiem, e-pasta ziņojumiem, sarunām un citiem). Pirms turpināt, pārliecinieties, vai uzticaties personai vai organizācijai, kas lūdza jums piekļūt šim rīkam.",
      "warningLine2": "Uzziniet vairāk šeit: {0}"
    },
    "_mraBnnuq2J9WjrAcnw9QNA": {
      "debugManifestErrorDetail": "Ielādējot atkļūdošanas manifestus, radās kļūda.",
      "debugManifestErrorDismissButtonText": "Vairs nerādīt"
    },
    "_upo3vfLFBbnbzl2hKy2TwA": {
      "allowDebugManifestsTitle": "Vai atļaut skriptu atkļūdošanu?",
      "allowDebugLoaderTitle": "Vai atļaut atkļūdošanas ielādētāju?",
      "allowDebugLoaderAndManifestsTitle": "Vai atļaut atkļūdošanas ielādētāju un atkļūdot skriptus?",
      "debugManifestLoadingWarning": "BRĪDINĀJUMS: šajā lapā ir nedroši skripti, kas, tos ielādējot, var kaitēt jūsu datoram. Neturpiniet, ja vien neuzticaties izstrādātājam un neizprotat riskus.",
      "debugManifestLoadingWarning2": "Ja īsti nezināt, noklikšķiniet uz {0}.",
      "debugManifestLoadingConfirm": "Ielādēt atkļūdošanas skriptus",
      "debugManifestLoadingCancel": "Neielādēt atkļūdošanas skriptus",
      "debugManifestLoadingCalloutText": "Ja nezināt, ko darīt, noklikšķiniet šeit."
    },
    "_SxImp5ewsUToxeAHBkB+pw": {
      "developerToolsTabLoadingText": "Notiek ielāde...",
      "developerToolsTabLoadingUnknownError": "Ielādējot izstrādātāju rīku moduli, radās nezināma kļūda."
    },
    "_sovI4qDAUPMnD4jg3Vsyfg": {
      "tabTitle": "Manifesti",
      "noManifestSelected": "Nav atlasīts manifests"
    },
    "_g7G0QHJ5bQYlxe+lk+DcxA": {
      "TabTitle": "Veiktspēja",
      "ErrorAccessingPerfDataErrorMessage": "Nevar izgūt veiktspējas datus: objekts ir Null vai nedefinēts.",
      "ErrorAccessingRedirectDataErrorMessage": "Piekļūstot HTTP novirzīšanas veiktspējas datiem, radās problēma.",
      "ErrorParsingPercievedLatencyErrorMessage": "Parsējot uztvertos latentuma datus, tika atrasta kļūda.",
      "ErrorParsingApiDataErrorMessage": "Parsējot API datus, tika atrasta kļūda.",
      "UnkownPerformanceDataErrorMessage": "Radās nezināma kļūda: {0}",
      "DefaultWebPartName": "Tīmekļa daļa",
      "ServerResponseLabel": "Servera atbilde",
      "ApplicationInitializationLabel": "Lietojumprogrammas inicializācija",
      "ScriptFetchEvalLabel": "Skriptu ienese un novērtēšana",
      "SpLoaderStartLabel": "SPFx inicializācija",
      "PageRenderLabel": "Lapas atveide",
      "LeftNavRenderLabel": "Kreisās navigācijas atveide",
      "CanvasRenderLabel": "Pamatnes atveidošana",
      "LayoutRenderLabel": "Izkārtojuma atveide",
      "RedirectResponseLabel": "Novirzīt atbildi",
      "AppLoadLabel": "Lietojumprogrammas ielāde",
      "RenderWebPartsLabel": "Tīmekļa daļu atveide",
      "TotalRenderTimeLabel": "Kopsumma",
      "GeneralErrorMessage": "Diemžēl, izgūstot veiktspējas datus, radās problēma.",
      "ErrorMessagePrefix": "Kļūdas ziņojums: {0}",
      "PerformanceDataHint": "Piezīme. Pēc tīmekļa daļas pievienošanas vai noņemšanas atsvaidziniet lapu, lai redzētu atjauninātos veiktspējas datus.",
      "ModulesLoadedLegendLabel": "Ielādētie moduļi",
      "InitializationLegendLabel": "Inicializācija",
      "RenderTimeLegendLabel": "Atveides laiks",
      "InitializationTimeLabel": "Inicializācijas laiks",
      "ModuleLoadingTimeLabel": "Moduļa ielādes laiks",
      "ModuleLazyLoadingDelayLabel": "Moduļa ielāde aizkavēta",
      "DataFetchTimeLabel": "Datu ieneses laiks",
      "DataFetchLegendLabel": "Datu ienese",
      "ItemsColumnHeader": "Vienumi",
      "DurationColumnHeader": "Ilgums",
      "MillisecondsUnitLabel": "{0} ms",
      "NAPlaceholder": "Nav datu"
    },
    "_gqinlPQb8HZprTeCpwNz2w": {
      "TabTitle": "Trasējums",
      "EmptyTraceData": "Trasējumi nav ielādēti.",
      "ExportCSVButtonLabel": "Eksportēt CSV",
      "LevelHeaderLabel": "Līmenis",
      "MessageHeaderLabel": "Ziņojums",
      "ScopeHeaderLabel": "Tvērums",
      "SourceHeaderLabel": "Avots",
      "TimestampHeaderLabel": "Laikspiedols",
      "TimestampFormat": "{0}/{1}/{2} {3}:{4}:{5}.{6}",
      "ErrorAccessingTraceDataErrorMessage": "Piekļūstot trasēšanas datiem, radās problēma."
    }
  };

  strings.default = strings;
  return strings;
});