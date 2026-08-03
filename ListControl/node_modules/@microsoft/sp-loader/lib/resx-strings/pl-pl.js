define([], function() {
  var strings = {
    "_nQNACBeQ34aV6bVwtFBayA": {
      "loaderUserFriendlyError": "Nie można załadować aplikacji na tej stronie. Użyj przycisku Wstecz przeglądarki, aby ponowić próbę. Jeśli problem będzie nadal występować, skontaktuj się z administratorem witryny i przekaż mu informacje w sekcji Szczegóły techniczne.",
      "platformFailedToLoadError": "Nie można załadować platformy. Identyfikator: „{0}”. Nazwa: „{1}”",
      "platformFailedToLoadWithMessageError": "Nie można załadować platformy. Identyfikator: „{0}”, nazwa: „{1}”.\r\nBłąd: {2}",
      "applicationFailedToInitializeError": "Błąd podczas inicjowania aplikacji. Błąd: {0}",
      "invalidPreloadedDataError": "Nieprawidłowe dane załadowane wstępnie.",
      "manifestNotFoundByIdError": "Nie znaleziono manifestu dla identyfikatora składnika „{0}”.",
      "manifestNotFoundError": "Nie można odnaleźć manifestu składnika o identyfikatorze „{0}” i wersji „{1}”.",
      "systemConfigDisabledError": "Metoda System.config() nie jest obsługiwana. Określ konfigurację przy użyciu manifestu.",
      "loadComponentLog": "Ładowanie składnika „{0}” ({1}).",
      "loadComponentEndLog": "Załadowano składnik „{0}” ({1}).",
      "loadComponentRetryLog": "Ładowanie składnika „{0}” ({1}). Próba {2} z {3}.",
      "loadComponentError": "Nie można załadować składnika „{0}” ({1}).\r\nOryginalny błąd: {2}",
      "loadComponentMaxRetriesError": "Próbowano załadować składnik „{0}” ({1}) {2} razy bez powodzenia.",
      "loadComponentDependencyError": "Nie można załadować zależności składnika „{0}” ze składnika „{1}” ({2}).\r\nOryginalny błąd: {3}",
      "loadComponentDependencyFailoverPathError": "Nie można załadować zależności składnika „{0}” ze ścieżką trybu failover „{1}” ze składnika „{2}” ({3}).\r\nOryginalny błąd: {4}",
      "loadPathDependencyLog": "Ładowanie zależności ścieżki „{0}” ze składnika „{1}” ({2})",
      "loadPathDependencyError": "Nie można załadować zależności ścieżki „{0}” ze składnika „{1}” ({2}).\r\nOryginalny błąd: {3}",
      "loadPathDependencyBlockedByAnotherDependencyError": "Nie można załadować zależności ścieżki „{0}” ze składnika „{1}” ({2}) ze względu na inną zależność, której nie można załadować.",
      "loadEntryPointError": "Nie można załadować punktu wejścia ze składnika „{0}” ({1}).\r\nOryginalny błąd: {2}",
      "loadComponentReturnsEmptyError": "Metoda loadComponent() zwróciła pusty obiekt składnika „{0}” ({1}).",
      "loadComponentReturnsDefaultEmptyError": "Metoda loadComponent() zwróciła obiekt z pustą właściwością domyślną składnika „{0}” ({1}).",
      "moduleHasUndeclaredDependencyError": "Punkt wejścia dla składnika „{0}” ({1}) ma zależność „{2}”, która nie została zadeklarowana w manifeście.",
      "loadScriptWithStringError": "Funkcja loadScript nie zezwala na użycie ciągu jako drugiego parametru. Zamiast tego użyj interfejsu ILoadScriptOptions.",
      "tooManyManifestsError": "Znaleziono manifesty ({0}, wersje {1}) dla składnika „{2}”.",
      "tooManyCompatibleVersionsError": "Znaleziono zgodne wersje ({0}: {1}) dla składnika „{2}” w wersji „{3}”.",
      "tooManyComponentsError": "Znaleziono zbyt wiele składników identyfikatora „{0}”.",
      "noComponentFoundError": "Nie znaleziono składnika o identyfikatorze „{0}”.",
      "deleteComponentLog": "Usuwanie wersji „{1}” składnika „{0}” z magazynu.",
      "browserNotSupportedError": "Ta wersja przeglądarki nie jest obsługiwana.\r\nZaktualizuj przeglądarkę do najnowszej wersji.",
      "ie9OrOlderNotSupportedError": "Ta strona nie obsługuje wersji programu Internet Explorer starszych niż wersja 10. Zaktualizuj przeglądarkę internetową.",
      "firefox43OrOlderNotSupportedError": "Ta strona nie obsługuje wersji programu Mozilla Firefox starszych niż wersja 44. Zaktualizuj przeglądarkę internetową.",
      "resourceNotFoundError": "Nie znaleziono zasobu „{0}” w konfiguracji modułu ładującego manifestu składnika „{1}” ({2}).",
      "noFailoverPathError": "Nie można wywołać elementu resolveAddress() dla składnika bez ścieżki trybu failover",
      "urlStatusLocalhostFileNotFoundError": "Nie można załadować adresu URL „{3}” dla zasobu „{2}” w składniku „{0}” ({1}). Nie odnaleziono pliku na serwerze.\r\nUpewnij się, że uruchamiasz polecenie „gulp serve”.",
      "urlStatusFileNotFoundError": "Nie można załadować adresu URL „{3}” dla zasobu„{2}” w składniku „{0}” ({1}). Nie odnaleziono pliku na serwerze.",
      "urlStatusForbiddenError": "Nie można załadować adresu URL „{3}” dla zasobu„{2}” w składniku „{0}” ({1}). Dostęp do pliku jest zabroniony.",
      "urlStatusClientErrorError": "Nie można załadować adresu URL „{3}” dla zasobu„{2}” w składniku „{0}” ({1}). Wystąpił błąd podczas żądania pliku.",
      "urlStatusServerErrorError": "Nie można załadować adresu URL „{3}” dla zasobu„{2}” w składniku „{0}” ({1}). Wystąpił problem na serwerze.",
      "urlStatusLocalhostNetworkErrorError": "Nie można załadować adresu URL „{3}” dla zasobu„{2}” w składniku „{0}” ({1}). Wystąpił problem z siecią.\r\nUpewnij się, że zostały uruchomione polecenia „gulp serve” i „gulp trust-dev-cert”.",
      "urlStatusHttpsNetworkErrorError": "Nie można załadować adresu URL „{3}” dla zasobu„{2}” w składniku „{0}” ({1}). Wystąpił problem z siecią.\r\nMoże to być problem z certyfikatem HTTPS. Upewnij się, że masz odpowiedni certyfikat.",
      "urlStatusNetworkErrorError": "Nie można załadować adresu URL „{3}” dla zasobu„{2}” w składniku „{0}” ({1}). Wystąpił problem z siecią.",
      "urlStatusUndefinedError": "Nie można załadować adresu URL „{3}” dla zasobu„{2}” w składniku „{0}” ({1}) z powodu nieznanych problemów.",
      "isUndefinedValidateError": "Wartość zmiennej „{0}” nie może być niezdefiniowana",
      "failedToCreateGlobalVariableError": "Nie można utworzyć zmiennej globalnej „{0}” ze skryptu „{1}”",
      "dependencyLoadError": "Nie można załadować modułu „{0}”, ponieważ nie załadowano zależności {1}",
      "missingPathDependencyError": "Brak zależności ścieżki „{0}” ze składnika „{1}” ({2}). Istniejące zależności ścieżki: {3}",
      "listSeparator": ", "
    },
    "_FmFyAWZ1md7Z1R+V8t2S2Q": {
      "errorLoadingDebugScriptHTTPS": "Wystąpił błąd podczas ładowania skryptu debugowania. Upewnij się, że serwer jest uruchomiony oraz że parametr adresu URL „{0}” jest poprawny.",
      "errorLoadingDebugScriptHTTP": "Wystąpił błąd podczas ładowania skryptu debugowania. Upewnij się, że serwer jest uruchomiony, parametr adresu URL „{0}” jest poprawny, a ładowanie niebezpiecznych skryptów jest dozwolone. Rozważ również użycie certyfikatu opracowywania i obsługiwanie skryptów debugowania za pomocą protokołu HTTPS.",
      "errorLoadingDebugScriptMalformed": "Wystąpił błąd podczas ładowania skryptu debugowania. Adres URL debugowania ({0}) prawdopodobnie jest źle sformułowany.",
      "errorLoadingDebugScriptUnknown": "Wystąpił nieznany błąd podczas ładowania skryptu debugowania.",
      "errorLoadingDebugLoaderTitle": "Wystąpił błąd podczas ładowania modułu ładującego debugowania.",
      "errorLoadingDebugManifestTitle": "Błąd podczas ładowania manifestów debugowania.",
      "errorLoadingUnknownTitle": "Wystąpił błąd podczas ładowania skryptów debugowania."
    },
    "_RPELcTeq3ZByqi3N5dt18w": {
      "missingDeveloperToolsTabInitFunctionError": "Brak składnika lub funkcji inicjatora.",
      "closeDeveloperToolsAriaLabel": "Zamknij narzędzia deweloperskie."
    },
    "_fwMQe6Xe08yEeCPNxngd+g": {
      "warningHeading": "Ostrzeżenie!",
      "warningLine1": "Korzystanie z tego narzędzia naraża Cię na potencjalne zagrożenia bezpieczeństwa, które mogą doprowadzić do uzyskania przez inne osoby dostępu do Twoich osobistych danych w usłudze Office 365 (dokumentów, wiadomości e-mail, konwersacji i nie tylko). Przed kontynuowaniem upewnij się, że ufasz osobie lub organizacji, która poprosiła Cię o dostęp do tego narzędzia.",
      "warningLine2": "Więcej informacji znajdziesz tutaj: {0}"
    },
    "_mraBnnuq2J9WjrAcnw9QNA": {
      "debugManifestErrorDetail": "Wystąpił błąd podczas ładowania manifestów debugowania.",
      "debugManifestErrorDismissButtonText": "Odrzuć"
    },
    "_upo3vfLFBbnbzl2hKy2TwA": {
      "allowDebugManifestsTitle": "Zezwolić na skrypty debugowania?",
      "allowDebugLoaderTitle": "Zezwolić na moduł ładujący debugowania?",
      "allowDebugLoaderAndManifestsTitle": "Zezwolić na moduł ładujący debugowania i skrypty debugowania?",
      "debugManifestLoadingWarning": "OSTRZEŻENIE: ta strona zawiera niebezpieczne skrypty, których załadowanie może uszkodzić komputer. Nie kontynuuj, chyba że ufasz deweloperowi i rozumiesz istotę zagrożenia.",
      "debugManifestLoadingWarning2": "Jeśli nie masz pewności, kliknij przycisk {0}.",
      "debugManifestLoadingConfirm": "Załaduj skrypty debugowania",
      "debugManifestLoadingCancel": "Nie ładuj skryptów debugowania",
      "debugManifestLoadingCalloutText": "Jeśli nie wiesz, co zrobić, kliknij tutaj."
    },
    "_SxImp5ewsUToxeAHBkB+pw": {
      "developerToolsTabLoadingText": "Trwa ładowanie...",
      "developerToolsTabLoadingUnknownError": "Nieznany błąd podczas ładowania modułu narzędzi deweloperskich."
    },
    "_sovI4qDAUPMnD4jg3Vsyfg": {
      "tabTitle": "Manifesty",
      "noManifestSelected": "Nie wybrano żadnego manifestu"
    },
    "_g7G0QHJ5bQYlxe+lk+DcxA": {
      "TabTitle": "Wydajność",
      "ErrorAccessingPerfDataErrorMessage": "Nie można pobrać danych dotyczących wydajności: obiekt ma wartość null lub jest niezdefiniowany.",
      "ErrorAccessingRedirectDataErrorMessage": "Wystąpił problem podczas uzyskiwania dostępu do danych dotyczących wydajności przekierowywania HTTP.",
      "ErrorParsingPercievedLatencyErrorMessage": "Wystąpił błąd podczas analizowania danych dotyczących postrzeganych opóźnień.",
      "ErrorParsingApiDataErrorMessage": "Wystąpił błąd podczas analizowania danych interfejsu API.",
      "UnkownPerformanceDataErrorMessage": "Wystąpił nieznany błąd: {0}",
      "DefaultWebPartName": "Składnik Web Part",
      "ServerResponseLabel": "Odpowiedź serwera",
      "ApplicationInitializationLabel": "Inicjowanie aplikacji",
      "ScriptFetchEvalLabel": "Pobieranie i ocena skryptu",
      "SpLoaderStartLabel": "Inicjowanie struktury SPFx",
      "PageRenderLabel": "Renderowanie strony",
      "LeftNavRenderLabel": "Renderowanie lewego obszaru nawigacji",
      "CanvasRenderLabel": "Renderowanie kanwy",
      "LayoutRenderLabel": "Renderowanie układu",
      "RedirectResponseLabel": "Odpowiedź przekierowywania",
      "AppLoadLabel": "Ładowanie aplikacji",
      "RenderWebPartsLabel": "Renderowanie składników Web Part",
      "TotalRenderTimeLabel": "Suma",
      "GeneralErrorMessage": "Niestety, wystąpił problem podczas pobierania danych dotyczących wydajności.",
      "ErrorMessagePrefix": "Komunikat o błędzie: {0}",
      "PerformanceDataHint": "Uwaga: po dodaniu lub usunięciu składnika Web Part odśwież stronę, aby wyświetlić zaktualizowane dane dotyczące wydajności.",
      "ModulesLoadedLegendLabel": "Załadowane moduły",
      "InitializationLegendLabel": "Inicjowanie",
      "RenderTimeLegendLabel": "Czas renderowania",
      "InitializationTimeLabel": "Czas inicjowania",
      "ModuleLoadingTimeLabel": "Czas ładowania modułu",
      "ModuleLazyLoadingDelayLabel": "Opóźnione ładowanie modułu",
      "DataFetchTimeLabel": "Czas pobierania danych",
      "DataFetchLegendLabel": "Pobieranie danych",
      "ItemsColumnHeader": "Elementy",
      "DurationColumnHeader": "Czas trwania",
      "MillisecondsUnitLabel": "{0} ms",
      "NAPlaceholder": "Brak"
    },
    "_gqinlPQb8HZprTeCpwNz2w": {
      "TabTitle": "Śledzenie",
      "EmptyTraceData": "Nie załadowano żadnych danych śledzenia.",
      "ExportCSVButtonLabel": "Eksportuj plik CSV",
      "LevelHeaderLabel": "Poziom",
      "MessageHeaderLabel": "Komunikat",
      "ScopeHeaderLabel": "Zakres",
      "SourceHeaderLabel": "Źródło",
      "TimestampHeaderLabel": "Sygnatura czasowa",
      "TimestampFormat": "{0}/{1}/{2} {3}:{4}:{5}.{6}",
      "ErrorAccessingTraceDataErrorMessage": "Wystąpił problem podczas uzyskiwania dostępu do danych śledzenia."
    }
  };

  strings.default = strings;
  return strings;
});