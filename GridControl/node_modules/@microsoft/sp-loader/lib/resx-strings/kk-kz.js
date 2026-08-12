define([], function() {
  var strings = {
    "_nQNACBeQ34aV6bVwtFBayA": {
      "loaderUserFriendlyError": "Осы беттегі бағдарламаны жүктеу мүмкін емес. Әрекетті қайталау үшін, браузердің \"Артқа\" түймешігін пайдаланыңыз. Егер мәселе шешілмесе, веб-сайттың әкімшісіне хабарласыңыз және оған \"Техникалық мәліметтер\" бөліміндегі ақпаратты ұсыныңыз.",
      "platformFailedToLoadError": "Платформаны жүктеу сәтсіз аяқталды. Ид.: \"{0}\", аты: \"{1}\"",
      "platformFailedToLoadWithMessageError": "Платформаны жүктеу мүмкін болмады. Идентификатор: \"{0}\". Аты: \"{1}\".\r\nҚате: {2}",
      "applicationFailedToInitializeError": "Бағдарламаны іске қосу кезінде қате пайда болды. Қате: {0}",
      "invalidPreloadedDataError": "Алдын ала жүктелген деректер жарамсыз.",
      "manifestNotFoundByIdError": "\"{0}\" идентификаторы үшін манифест табылмады.",
      "manifestNotFoundError": "\"{0}\" компонент идентификаторы және \"{1}\" нұсқасы үшін манифест табылмады.",
      "systemConfigDisabledError": "System.config() үшін қолдау көрсетілмейді. Конфигурацияны көрсету үшін, манифестті пайдаланыңыз.",
      "loadComponentLog": "\"{0}\" ({1}) компоненті жүктелуде.",
      "loadComponentEndLog": "\"{0}\" ({1}) компоненті жүктелді.",
      "loadComponentRetryLog": "\"{0}\" ({1}) компоненті жүктелуде. {2}/{3}-талпыныс.",
      "loadComponentError": "\"{0}\" ({1}) компонентін жүктеу сәтсіз аяқталды.\r\nБастапқы қате: {2}",
      "loadComponentMaxRetriesError": "\"{0}\" ({1}) компонентін жүктеу әрекеті {2} рет сәтсіз орындалды.",
      "loadComponentDependencyError": "\"{1}\" ({2}) компонентіндегі \"{0}\" тәуелділігін жүктеу сәтсіз аяқталды.\r\nБастапқы қате: {3}",
      "loadComponentDependencyFailoverPathError": "\"{2}\" ({3}) компонентіндегі \"{1}\" ақаулығын өңдеу жолы арқылы \"{0}\" компонент тәуелділігін жүктеу сәтсіз аяқталды.\r\nБастапқы қате: {4}",
      "loadPathDependencyLog": "\"{1}\" ({2}) компонентінен \"{0}\" жол тәуелділігі жүктелуде",
      "loadPathDependencyError": "\"{1}\" ({2}) компонентіндегі \"{0}\" жол тәуелділігін жүктеу сәтсіз аяқталды.\r\nБастапқы қате: {3}",
      "loadPathDependencyBlockedByAnotherDependencyError": "Басқа тәуелділікті жүктеу сәтсіз аяқталғандықтан, \"{1}\" ({2}) компонентінен \"{0}\" жол тәуелділігін жүктеу сәтсіз аяқталды.",
      "loadEntryPointError": "\"{0}\" ({1}) компонентінен кіріс нүктесін жүктеу сәтсіз аяқталды.\r\nБастапқы қате: {2}",
      "loadComponentReturnsEmptyError": "loadComponent() \"{0}\" ({1}) компоненті үшін бос нысанды қайтарды.",
      "loadComponentReturnsDefaultEmptyError": "loadComponent() \"{0}\" ({1}) компоненті үшін бос әдепкі сипаты бар нысанды қайтарды.",
      "moduleHasUndeclaredDependencyError": "\"{0}\" ({1}) компонентіне арналған кіріс нүктесі манифестте берілмеген \"{2}\" тәуелділігін қамтиды.",
      "loadScriptWithStringError": "loadScript функциясында 2-ші параметр ретінде жолды пайдалануға тыйым салынады. Оның орнына ILoadScriptOptions функциясын пайдаланыңыз.",
      "tooManyManifestsError": "\"{2}\" компоненті үшін {0} манифест ({1} нұсқалары) табылды.",
      "tooManyCompatibleVersionsError": "\"{2}\" компонентінің \"{3}\" нұсқасы үшін {0} үйлесімді нұсқа ({1}) табылды.",
      "tooManyComponentsError": "\"{0}\" идентификаторы үшін өте көп компонент табылды.",
      "noComponentFoundError": "\"{0}\" идентификаторы бойынша компонент табылмады.",
      "deleteComponentLog": "\"{0}\" компонентінің \"{1}\" нұсқасы дүкеннен жойылуда.",
      "browserNotSupportedError": "Браузеріңіздің нұсқасына қолдау көрсетілмейді.\r\nБраузеріңізді соңғы нұсқаға жаңартыңыз.",
      "ie9OrOlderNotSupportedError": "Бұл бетте Internet Explorer браузерінің 10 нұсқасынан ескі нұсқаларына қолдау көрсетілмейді. Веб-браузеріңізді жаңартыңыз.",
      "firefox43OrOlderNotSupportedError": "Бұл бетте Mozilla Firefox браузерінің 44 нұсқасынан ескі нұсқаларына қолдау көрсетілмейді. Веб-браузеріңізді жаңартыңыз.",
      "resourceNotFoundError": "\"{1}\" ({2}) компонентіне арналған манифестті жүктеп алу құралының конфигурациясында \"{0}\" ресурсы табылмады.",
      "noFailoverPathError": "Басқа ресурсқа ауысу жолы жоқ құрамдаста resolveAddress() шақыру мүмкін емес",
      "urlStatusLocalhostFileNotFoundError": "\"{0}\" ({1}) компонентіндегі \"{2}\" ресурсына арналған \"{3}\" URL мекенжайын жүктеу сәтсіз аяқталды. Файл серверде табылмады.\r\n\"Gulp serve\" іске қосылғанын тексеріңіз.",
      "urlStatusFileNotFoundError": "\"{0}\" ({1}) компонентіндегі \"{2}\" ресурсына арналған \"{3}\" URL мекенжайын жүктеу сәтсіз аяқталды. Файл серверде табылмады.",
      "urlStatusForbiddenError": "\"{0}\" ({1}) компонентіндегі \"{2}\" ресурсына арналған \"{3}\" URL мекенжайын жүктеу сәтсіз аяқталды. Файлға қатынасу мүмкіндігіне тыйым салынды.",
      "urlStatusClientErrorError": "\"{0}\" ({1}) компонентіндегі \"{2}\" ресурсына арналған \"{3}\" URL мекенжайын жүктеу сәтсіз аяқталды. Файлды сұрау кезінде қате пайда болды.",
      "urlStatusServerErrorError": "\"{0}\" ({1}) компонентіндегі \"{2}\" ресурсына арналған \"{3}\" URL мекенжайын жүктеу сәтсіз аяқталды. Серверде ақаулық пайда болды.",
      "urlStatusLocalhostNetworkErrorError": "\"{0}\" ({1}) компонентіндегі \"{2}\" ресурсына арналған \"{3}\" URL мекенжайын жүктеу сәтсіз аяқталды. Желі ақаулығы орын алды.\r\n\"Gulp serve\" опциясының пайдаланылып жатқанын және \"gulp trust-dev-cert\" іске қосылғанын тексеріңіз.",
      "urlStatusHttpsNetworkErrorError": "\"{0}\" ({1}) компонентіндегі \"{2}\" ресурсына арналған \"{3}\" URL мекенжайын жүктеу сәтсіз аяқталды. Желі ақаулығы пайда болды.\r\nHTTPS сертификатына қатысты ақау орын алған болуы мүмкін. Сертификаттың дұрыстығына көз жеткізіңіз.",
      "urlStatusNetworkErrorError": "\"{0}\" ({1}) компонентіндегі \"{2}\" ресурсына арналған \"{3}\" URL мекенжайын жүктеу сәтсіз аяқталды. Желі ақаулығы пайда болды.",
      "urlStatusUndefinedError": "Белгісіз қате орын алғандықтан, \"{0}\" ({1}) компонентіндегі \"{2}\" ресурсына арналған \"{3}\" URL мекенжайын жүктеу сәтсіз аяқталды.",
      "isUndefinedValidateError": "\"{0}\" мәні анықталмаған болмауы тиіс",
      "failedToCreateGlobalVariableError": "\"{1}\" сценарийінен \"{0}\" айнымалысын жасау сәтсіз аяқталды",
      "dependencyLoadError": "\"{0}\" модулін жүктеу мүмкін болмады, себебі {1} тәуелділігі жүктелмеді",
      "missingPathDependencyError": "\"{1}\" ({2}) компонентінде \"{0}\" жол тәуелділігі жоқ. Бар жол тәуелділіктері: {3}",
      "listSeparator": ", "
    },
    "_FmFyAWZ1md7Z1R+V8t2S2Q": {
      "errorLoadingDebugScriptHTTPS": "Күйін келтіру сценарийін жүктеу кезінде қате пайда болды. Сервердің жұмыс істеп тұрғанына және \"{0}\" параметрінің URL мекенжайы дұрыс екеніне көз жеткізіңіз.",
      "errorLoadingDebugScriptHTTP": "Күйін келтіру сценарийін жүктеу кезінде қате пайда болды. Сервердің жұмыс істеп тұрғанына, \"{0}\" параметрінің URL мекенжайы дұрыс екеніне және қауіпті сценарийлерді жүктеуге рұқсат етілгеніне көз жеткізіңіз. Сондай-ақ HTTPS протоколдары бойынша әзірлеу сертификатын пайдалану және күйін келтіру сценарийлеріне қызмет көрсету мүмкіндігін қараңыз.",
      "errorLoadingDebugScriptMalformed": "Күйін келтіру сценарийін жүктеу кезінде қате пайда болды. Күйін келтірудің URL мекенжайының ({0}) пішімі қате болу мүмкін.",
      "errorLoadingDebugScriptUnknown": "Күйін келтіру сценарийін жүктеу кезінде белгісіз қате пайда болды.",
      "errorLoadingDebugLoaderTitle": "Күйін келтіруге арналған жүктеу құралын жүктеу кезінде қате пайда болды.",
      "errorLoadingDebugManifestTitle": "Күйін келтіру манифестерін жүктеу кезінде қате пайда болды.",
      "errorLoadingUnknownTitle": "Күйін келтіру сценарийлерін жүктеу кезінде қате пайда болды."
    },
    "_RPELcTeq3ZByqi3N5dt18w": {
      "missingDeveloperToolsTabInitFunctionError": "Компонент немесе инициализатор функциясы жоқ.",
      "closeDeveloperToolsAriaLabel": "Әзірлеуші құралдарын жабу."
    },
    "_fwMQe6Xe08yEeCPNxngd+g": {
      "warningHeading": "Ескерту!",
      "warningLine1": "Бұл құралдарды пайдалану қауіпсіздікке қауіп-қатер төндіруі және басқа пайдаланушылардың жеке Office 365 деректеріңізге (құжаттар, электрондық хаттар, сөйлесулер, т.б.) қатынасу мүмкіндігіне ие болуына алып келуі мүмкін. Жалғастырмас бұрын, осы құралға қатынасуды сұраған адамның немесе ұйымның сенімді екеніне көз жеткізіңіз.",
      "warningLine2": "Қосымша мәлімет: {0}"
    },
    "_mraBnnuq2J9WjrAcnw9QNA": {
      "debugManifestErrorDetail": "Күйін келтіру манифестерін жүктеу кезінде қате пайда болды.",
      "debugManifestErrorDismissButtonText": "Жабу"
    },
    "_upo3vfLFBbnbzl2hKy2TwA": {
      "allowDebugManifestsTitle": "Күйін келтіру сценарийлеріне рұқсат ету қажет пе?",
      "allowDebugLoaderTitle": "Күйін келтіруге арналған жүктеу құралына рұқсат ету қажет пе?",
      "allowDebugLoaderAndManifestsTitle": "Күйін келтіруге арналған жүктеу құралына және күйін келтіру сценарийлеріне рұқсат ету қажет пе?",
      "debugManifestLoadingWarning": "ЕСКЕРТУ: Бұл бет жүктелген жағдайда компьютеріңізге зиян келтіруі мүмкін қауіпті сценарийлерді қамтиды. Әзірлеушіге сенбесеңіз және төнетін қатерлердің қаншалықты қауіпті екенін білмесеңіз, процесті жалғастырмаңыз.",
      "debugManifestLoadingWarning2": "Сенімді болмасаңыз, {0} түймешігін басыңыз.",
      "debugManifestLoadingConfirm": "Реттеу сценарийлерін жүктеу",
      "debugManifestLoadingCancel": "Реттеу сценарийін жүктемеу",
      "debugManifestLoadingCalloutText": "Егер не істеу керектігін білмесеңіз, осы жерді басыңыз."
    },
    "_SxImp5ewsUToxeAHBkB+pw": {
      "developerToolsTabLoadingText": "Жүктелуде...",
      "developerToolsTabLoadingUnknownError": "Әзірлеуші құралдарының модулін жүктеу кезінде белгісіз қате орын алды."
    },
    "_sovI4qDAUPMnD4jg3Vsyfg": {
      "tabTitle": "Манифестер",
      "noManifestSelected": "Манифест таңдалмады"
    },
    "_g7G0QHJ5bQYlxe+lk+DcxA": {
      "TabTitle": "Өнімділік",
      "ErrorAccessingPerfDataErrorMessage": "Өнімділікке қатысты деректі алу мүмкін емес: нәтижесі жоқ немесе анықталмаған.",
      "ErrorAccessingRedirectDataErrorMessage": "HTTP бағытын өзгертудің өнімділік деректеріне қатынасу кезінде ақаулық орын алды.",
      "ErrorParsingPercievedLatencyErrorMessage": "Деректер кідірісін талдау кезінде қате туралы хабар алынды.",
      "ErrorParsingApiDataErrorMessage": "API деректерін талдау кезінде қате туралы хабар алынды.",
      "UnkownPerformanceDataErrorMessage": "Белгісіз қате пайда болды: {0}",
      "DefaultWebPartName": "Веб-бөлік",
      "ServerResponseLabel": "Сервер жауабы",
      "ApplicationInitializationLabel": "Бағдарламаны іске қосу",
      "ScriptFetchEvalLabel": "Сценарийді таңдау және бағалау",
      "SpLoaderStartLabel": "SPFx инициализациясы",
      "PageRenderLabel": "Бетті көрсету",
      "LeftNavRenderLabel": "Сол жақ навигацияны көрсету",
      "CanvasRenderLabel": "Кенепті көрсету",
      "LayoutRenderLabel": "Орналасуды көрсету",
      "RedirectResponseLabel": "Бағытын өзгерту жауабы",
      "AppLoadLabel": "Бағдарламаны жүктеу",
      "RenderWebPartsLabel": "Веб-бөліктерді көрсету",
      "TotalRenderTimeLabel": "Жалпы",
      "GeneralErrorMessage": "Кешіріңіз, өнімділік деректерін алу кезінде белгісіз бір қате пайда болды.",
      "ErrorMessagePrefix": "Қате туралы хабар: {0}",
      "PerformanceDataHint": "Ескертпе: Веб-бөлікті қосқаннан кейін не жойғаннан кейін, өнімділікке қатысты жаңартылған деректерді көру үшін бетті жаңартыңыз.",
      "ModulesLoadedLegendLabel": "Модульдер жүктелді",
      "InitializationLegendLabel": "Баптандыру",
      "RenderTimeLegendLabel": "Көрсету уақыты",
      "InitializationTimeLabel": "Инициализация уақыты",
      "ModuleLoadingTimeLabel": "Модульді жүктеу уақыты",
      "ModuleLazyLoadingDelayLabel": "Модульдерді жүктеу кідіртілді",
      "DataFetchTimeLabel": "Деректерді таңдау уақыты",
      "DataFetchLegendLabel": "Деректерді шығарып алу",
      "ItemsColumnHeader": "Элементтер",
      "DurationColumnHeader": "Ұзақтық",
      "MillisecondsUnitLabel": "{0} мс",
      "NAPlaceholder": "Жоқ"
    },
    "_gqinlPQb8HZprTeCpwNz2w": {
      "TabTitle": "Қадағау",
      "EmptyTraceData": "Қадағау деректері жүктелмеді.",
      "ExportCSVButtonLabel": "CSV файлы ретінде экспорттау",
      "LevelHeaderLabel": "Деңгей",
      "MessageHeaderLabel": "Хабар",
      "ScopeHeaderLabel": "Аумақ",
      "SourceHeaderLabel": "Ақпарат көзі",
      "TimestampHeaderLabel": "Уақыт белгісі",
      "TimestampFormat": "{0}/{1}/{2} {3}:{4}:{5}.{6}",
      "ErrorAccessingTraceDataErrorMessage": "Қадағау деректеріне қатынасу кезінде ақаулық пайда болды."
    }
  };

  strings.default = strings;
  return strings;
});