using JuliusSweetland.OptiKey.Enums;
using JuliusSweetland.OptiKey.Properties;
using JuliusSweetland.OptiKey.Services;
using log4net;
using Prism.Mvvm;
using System.Collections.Generic;

namespace JuliusSweetland.OptiKey.UI.ViewModels.Management
{
    public class WordsViewModel : BindableBase
    {
        private readonly IDictionaryService dictionaryService;

        #region Private Member Vars

        private static readonly ILog Log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        #endregion

        #region Ctor

        public WordsViewModel(IDictionaryService dictionaryService)
        {
            this.dictionaryService = dictionaryService;

            Load();
        }

        #endregion

        #region Properties

        public List<KeyValuePair<string, Languages>> Languages
        {
            get
            {
                return new List<KeyValuePair<string, Languages>>
                {
                    new KeyValuePair<string, Languages>(Resources.CATALAN_SPAIN, Enums.Languages.CatalanSpain),
                    new KeyValuePair<string, Languages>(Resources.CROATIAN_CROATIA, Enums.Languages.CroatianCroatia),
                    new KeyValuePair<string, Languages>(Resources.CZECH_CZECH_REPUBLIC, Enums.Languages.CzechCzechRepublic),
                    new KeyValuePair<string, Languages>(Resources.DANISH_DENMARK, Enums.Languages.DanishDenmark),
                    new KeyValuePair<string, Languages>(Resources.DUTCH_BELGIUM, Enums.Languages.DutchBelgium),
                    new KeyValuePair<string, Languages>(Resources.DUTCH_NETHERLANDS, Enums.Languages.DutchNetherlands),
                    new KeyValuePair<string, Languages>(Resources.ENGLISH_CANADA, Enums.Languages.EnglishCanada),
                    new KeyValuePair<string, Languages>(Resources.ENGLISH_UK, Enums.Languages.EnglishUK),
                    new KeyValuePair<string, Languages>(Resources.ENGLISH_US, Enums.Languages.EnglishUS),
                    new KeyValuePair<string, Languages>(Resources.FRENCH_FRANCE, Enums.Languages.FrenchFrance),
                    new KeyValuePair<string, Languages>(Resources.GERMAN_GERMANY, Enums.Languages.GermanGermany),
                    new KeyValuePair<string, Languages>(Resources.GREEK_GREECE, Enums.Languages.GreekGreece),
                    new KeyValuePair<string, Languages>(Resources.ITALIAN_ITALY, Enums.Languages.ItalianItaly),
                    new KeyValuePair<string, Languages>(Resources.KOREAN_KOREA, Enums.Languages.KoreanKorea),
                    new KeyValuePair<string, Languages>(Resources.PORTUGUESE_PORTUGAL, Enums.Languages.PortuguesePortugal),
                    new KeyValuePair<string, Languages>(Resources.RUSSIAN_RUSSIA, Enums.Languages.RussianRussia),
                    new KeyValuePair<string, Languages>(Resources.SLOVAK_SLOVAKIA, Enums.Languages.SlovakSlovakia),
                    new KeyValuePair<string, Languages>(Resources.SLOVENIAN_SLOVENIA, Enums.Languages.SlovenianSlovenia),
                    new KeyValuePair<string, Languages>(Resources.SPANISH_SPAIN, Enums.Languages.SpanishSpain),
                    new KeyValuePair<string, Languages>(Resources.TURKISH_TURKEY, Enums.Languages.TurkishTurkey)
                };
            }
        }

        public List<KeyValuePair<string, SuggestionMethods>> SuggestionMethods
        {
            get
                {
                    return new List<KeyValuePair<string, SuggestionMethods>> {
                        new KeyValuePair<string, SuggestionMethods>(Resources.BASIC_SUGGESTION, Enums.SuggestionMethods.Basic),
                        new KeyValuePair<string, SuggestionMethods>(Resources.NGRAM_SUGGESTION, Enums.SuggestionMethods.NGram),
                        new KeyValuePair<string, SuggestionMethods>(Resources.PRESAGE_SUGGESTION, Enums.SuggestionMethods.Presage)
                    };
                }
        }

        private Languages keyboardAndDictionaryLanguage;
        public Languages KeyboardAndDictionaryLanguage
        {
            get { return keyboardAndDictionaryLanguage; }
            set
            {
                SetProperty(ref this.keyboardAndDictionaryLanguage, value);
                OnPropertyChanged(() => UseAlphabeticalKeyboardLayoutIsVisible);
                OnPropertyChanged(() => EnableCommuniKateKeyboardLayoutIsVisible);
                OnPropertyChanged(() => UseCommuniKateKeyboardLayoutByDefault);
                OnPropertyChanged(() => UseSimplifiedKeyboardLayoutIsVisible);
            }
        }

        private Languages uiLanguage;
        public Languages UiLanguage
        {
            get { return uiLanguage; }
            set { SetProperty(ref this.uiLanguage, value); }
        }

        private bool useAlphabeticalKeyboardLayout;
        public bool UseAlphabeticalKeyboardLayout
        {
            get { return useAlphabeticalKeyboardLayout; }
            set { SetProperty(ref useAlphabeticalKeyboardLayout, 
                        value 
                        && useSimplifiedKeyboardLayout == false
                        && useCommuniKateKeyboardLayoutByDefault == false
                        && usingCommuniKateKeyboardLayout == false); }
        }

        public bool UseAlphabeticalKeyboardLayoutIsVisible
        {
            get
            {
                return KeyboardAndDictionaryLanguage == Enums.Languages.EnglishCanada
                    || KeyboardAndDictionaryLanguage == Enums.Languages.EnglishUK
                    || KeyboardAndDictionaryLanguage == Enums.Languages.EnglishUS;
            }
        }

        private bool useSimplifiedKeyboardLayout;
        public bool UseSimplifiedKeyboardLayout
        {
            get { return useSimplifiedKeyboardLayout; }
            set { SetProperty(ref useSimplifiedKeyboardLayout,
                        value
                        && useAlphabeticalKeyboardLayout == false
                        && useCommuniKateKeyboardLayoutByDefault == false
                        && usingCommuniKateKeyboardLayout == false);
            }
        }

        public bool UseSimplifiedKeyboardLayoutIsVisible
        {
            get
            {
                return KeyboardAndDictionaryLanguage == Enums.Languages.EnglishCanada
                    || KeyboardAndDictionaryLanguage == Enums.Languages.EnglishUK
                    || KeyboardAndDictionaryLanguage == Enums.Languages.EnglishUS;
            }
        }

        private bool enableCommuniKateKeyboardLayout;
        public bool EnableCommuniKateKeyboardLayout
        {
            get { return enableCommuniKateKeyboardLayout; }
            set { SetProperty(ref enableCommuniKateKeyboardLayout, value
                        && useAlphabeticalKeyboardLayout == false
                        && useSimplifiedKeyboardLayout == false
                        && (KeyboardAndDictionaryLanguage == Enums.Languages.EnglishCanada
                            || KeyboardAndDictionaryLanguage == Enums.Languages.EnglishUK
                            || KeyboardAndDictionaryLanguage == Enums.Languages.EnglishUS));
            }
        }

        public bool EnableCommuniKateKeyboardLayoutIsVisible
        {
            get
            {
                return KeyboardAndDictionaryLanguage == Enums.Languages.EnglishCanada
                    || KeyboardAndDictionaryLanguage == Enums.Languages.EnglishUK
                    || KeyboardAndDictionaryLanguage == Enums.Languages.EnglishUS;
            }
        }

        private string communiKatePagesetLocation;
        public string CommuniKatePagesetLocation
        {
            get { return communiKatePagesetLocation; }
            set { SetProperty(ref communiKatePagesetLocation, value); }
        }

        private bool communiKateStagedForDeletion;
        public bool CommuniKateStagedForDeletion
        {
            get { return communiKateStagedForDeletion; }
            set { SetProperty(ref communiKateStagedForDeletion, value); }
        }

        private bool useCommuniKateKeyboardLayoutByDefault;
        public bool UseCommuniKateKeyboardLayoutByDefault
        {
            get { return useCommuniKateKeyboardLayoutByDefault; }
            set { SetProperty(ref useCommuniKateKeyboardLayoutByDefault, value
                        && enableCommuniKateKeyboardLayout);
            }
        }

        private bool usingCommuniKateKeyboardLayout;
        public bool UsingCommuniKateKeyboardLayout
        {
            get { return usingCommuniKateKeyboardLayout; }
            set { SetProperty(ref usingCommuniKateKeyboardLayout, useCommuniKateKeyboardLayoutByDefault); }
        }

        private bool forceCapsLock;
        public bool ForceCapsLock
        {
            get { return forceCapsLock; }
            set { SetProperty(ref forceCapsLock, value); }
        }

        private bool autoAddSpace;
        public bool AutoAddSpace
        {
            get { return autoAddSpace; }
            set { SetProperty(ref autoAddSpace, value); }
        }

        private bool autoCapitalise;
        public bool AutoCapitalise
        {
            get { return autoCapitalise; }
            set { SetProperty(ref autoCapitalise, value); }
        }

        private bool suppressAutoCapitaliseIntelligently;
        public bool SuppressAutoCapitaliseIntelligently
        {
            get { return suppressAutoCapitaliseIntelligently; }
            set { SetProperty(ref suppressAutoCapitaliseIntelligently, value); }
        }

        private SuggestionMethods suggestionMethod;
        public SuggestionMethods SuggestionMethod
        {
            get { return suggestionMethod; }
            set { SetProperty(ref suggestionMethod, value); }
        }

        private bool suggestWords;
        public bool SuggestWords
        {
            get {  return suggestWords; }
            set { SetProperty(ref suggestWords, value); }
        }

        private bool multiKeySelectionEnabled;
        public bool MultiKeySelectionEnabled
        {
            get { return multiKeySelectionEnabled; }
            set { SetProperty(ref multiKeySelectionEnabled, value); }
        }

        private int multiKeySelectionMaxDictionaryMatches;
        public int MultiKeySelectionMaxDictionaryMatches
        {
            get { return multiKeySelectionMaxDictionaryMatches; }
            set { SetProperty(ref multiKeySelectionMaxDictionaryMatches, value); }
        }

        public bool ChangesRequireRestart
        {
            get { return ForceCapsLock != Settings.Default.ForceCapsLock
                    || Settings.Default.SuggestionMethod != SuggestionMethod
                    || Settings.Default.CommuniKateStagedForDeletion == true; }
        }

        #endregion

        #region Methods

        private void Load()
        {
            KeyboardAndDictionaryLanguage = Settings.Default.KeyboardAndDictionaryLanguage;
            UiLanguage = Settings.Default.UiLanguage;
            UseAlphabeticalKeyboardLayout = Settings.Default.UseAlphabeticalKeyboardLayout;
            EnableCommuniKateKeyboardLayout = Settings.Default.EnableCommuniKateKeyboardLayout;
            CommuniKatePagesetLocation = Settings.Default.CommuniKatePagesetLocation;
            CommuniKateStagedForDeletion = Settings.Default.CommuniKateStagedForDeletion;
            UseCommuniKateKeyboardLayoutByDefault = Settings.Default.UseCommuniKateKeyboardLayoutByDefault;
            UsingCommuniKateKeyboardLayout = Settings.Default.UseCommuniKateKeyboardLayoutByDefault;
            UseSimplifiedKeyboardLayout = Settings.Default.UseSimplifiedKeyboardLayout;
            ForceCapsLock = Settings.Default.ForceCapsLock;
            AutoAddSpace = Settings.Default.AutoAddSpace;
            AutoCapitalise = Settings.Default.AutoCapitalise;
            SuppressAutoCapitaliseIntelligently = Settings.Default.SuppressAutoCapitaliseIntelligently;
            SuggestionMethod = Settings.Default.SuggestionMethod;
            SuggestWords = Settings.Default.SuggestWords;
            MultiKeySelectionEnabled = Settings.Default.MultiKeySelectionEnabled;
            MultiKeySelectionMaxDictionaryMatches = Settings.Default.MaxDictionaryMatchesOrSuggestions;
        }

        public void ApplyChanges()
        {
            var reloadDictionary = (Settings.Default.KeyboardAndDictionaryLanguage != KeyboardAndDictionaryLanguage)
                                   || (Settings.Default.SuggestionMethod != SuggestionMethod);

            Settings.Default.KeyboardAndDictionaryLanguage = KeyboardAndDictionaryLanguage;
            Settings.Default.UiLanguage = UiLanguage;
            Settings.Default.UseAlphabeticalKeyboardLayout = UseAlphabeticalKeyboardLayout;
            Settings.Default.EnableCommuniKateKeyboardLayout = EnableCommuniKateKeyboardLayout;
            Settings.Default.CommuniKatePagesetLocation = CommuniKatePagesetLocation;
            Settings.Default.CommuniKateStagedForDeletion = CommuniKateStagedForDeletion;
            Settings.Default.UseCommuniKateKeyboardLayoutByDefault = UseCommuniKateKeyboardLayoutByDefault;
            Settings.Default.UsingCommuniKateKeyboardLayout = UseCommuniKateKeyboardLayoutByDefault;
            Settings.Default.UseSimplifiedKeyboardLayout = UseSimplifiedKeyboardLayout;
            Settings.Default.ForceCapsLock = ForceCapsLock;
            Settings.Default.AutoAddSpace = AutoAddSpace;
            Settings.Default.AutoCapitalise = AutoCapitalise;
            Settings.Default.SuppressAutoCapitaliseIntelligently = SuppressAutoCapitaliseIntelligently;
            Settings.Default.SuggestionMethod = SuggestionMethod;
            Settings.Default.SuggestWords = SuggestWords;
            Settings.Default.MultiKeySelectionEnabled = MultiKeySelectionEnabled;
            Settings.Default.MaxDictionaryMatchesOrSuggestions = MultiKeySelectionMaxDictionaryMatches;

            if (reloadDictionary)
            {
                dictionaryService.LoadDictionary();
            }
        }

        #endregion
    }
}
