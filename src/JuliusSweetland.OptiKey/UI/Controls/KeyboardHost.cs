﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Reactive.Disposables;
using System.Reactive.Linq;
using System.Reflection;
using System.Windows;
using System.Windows.Controls;
using JuliusSweetland.OptiKey.Enums;
using JuliusSweetland.OptiKey.Extensions;
using JuliusSweetland.OptiKey.Models;
using JuliusSweetland.OptiKey.Properties;
using JuliusSweetland.OptiKey.UI.Utilities;
using JuliusSweetland.OptiKey.UI.ViewModels.Keyboards.Base;
using log4net;
using CommonViews = JuliusSweetland.OptiKey.UI.Views.Keyboards.Common;
using CatalanViews = JuliusSweetland.OptiKey.UI.Views.Keyboards.Catalan;
using CroatianViews = JuliusSweetland.OptiKey.UI.Views.Keyboards.Croatian;
using CzechViews = JuliusSweetland.OptiKey.UI.Views.Keyboards.Czech;
using DanishViews = JuliusSweetland.OptiKey.UI.Views.Keyboards.Danish;
using DutchViews = JuliusSweetland.OptiKey.UI.Views.Keyboards.Dutch;
using EnglishViews = JuliusSweetland.OptiKey.UI.Views.Keyboards.English;
using FrenchViews = JuliusSweetland.OptiKey.UI.Views.Keyboards.French;
using GermanViews = JuliusSweetland.OptiKey.UI.Views.Keyboards.German;
using GreekViews = JuliusSweetland.OptiKey.UI.Views.Keyboards.Greek;
using ItalianViews = JuliusSweetland.OptiKey.UI.Views.Keyboards.Italian;
using KoreanViews = JuliusSweetland.OptiKey.UI.Views.Keyboards.Korean;
using PortugueseViews = JuliusSweetland.OptiKey.UI.Views.Keyboards.Portuguese;
using RussianViews = JuliusSweetland.OptiKey.UI.Views.Keyboards.Russian;
using SlovakViews = JuliusSweetland.OptiKey.UI.Views.Keyboards.Slovak;
using SlovenianViews = JuliusSweetland.OptiKey.UI.Views.Keyboards.Slovenian;
using SpanishViews = JuliusSweetland.OptiKey.UI.Views.Keyboards.Spanish;
using TurkishViews = JuliusSweetland.OptiKey.UI.Views.Keyboards.Turkish;
using ViewModelKeyboards = JuliusSweetland.OptiKey.UI.ViewModels.Keyboards;
using System.Diagnostics;

namespace JuliusSweetland.OptiKey.UI.Controls
{
    public class KeyboardHost : ContentControl
    {
        #region Private member vars

        private static readonly ILog Log = LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);

        private CompositeDisposable currentKeyboardKeyValueSubscriptions = new CompositeDisposable();

        #endregion

        #region Ctor

        public KeyboardHost()
        {
            Settings.Default.OnPropertyChanges(s => s.KeyboardAndDictionaryLanguage).Subscribe(_ => GenerateContent());
            Settings.Default.OnPropertyChanges(s => s.UiLanguage).Subscribe(_ => GenerateContent());
            Settings.Default.OnPropertyChanges(s => s.MouseKeyboardDockSize).Subscribe(_ => GenerateContent());
            Settings.Default.OnPropertyChanges(s => s.ConversationOnlyMode).Subscribe(_ => GenerateContent());
            Settings.Default.OnPropertyChanges(s => s.ConversationConfirmEnable).Subscribe(_ => GenerateContent());
            Settings.Default.OnPropertyChanges(s => s.ConversationConfirmOnlyMode).Subscribe(_ => GenerateContent());
            Settings.Default.OnPropertyChanges(s => s.UseAlphabeticalKeyboardLayout).Subscribe(_ => GenerateContent());
            Settings.Default.OnPropertyChanges(s => s.EnableCommuniKateKeyboardLayout).Subscribe(_ => GenerateContent());
            Settings.Default.OnPropertyChanges(s => s.UseCommuniKateKeyboardLayoutByDefault).Subscribe(_ => GenerateContent());
            Settings.Default.OnPropertyChanges(s => s.UseSimplifiedKeyboardLayout).Subscribe(_ => GenerateContent());
            Settings.Default.OnPropertyChanges(s => s.CommuniKateKeyboardCurrentContext).Subscribe(_ => GenerateContent());
            Settings.Default.OnPropertyChanges(s => s.SimplifiedKeyboardCurrentContext).Subscribe(_ => GenerateContent());

            Loaded += OnLoaded;

            var contentDp = DependencyPropertyDescriptor.FromProperty(ContentProperty, typeof(KeyboardHost));
            if (contentDp != null)
            {
                contentDp.AddValueChanged(this, ContentChangedHandler);
            }

            this.MouseEnter += this.OnMouseEnter;
        }

        #endregion

        #region Properties
        
        public static readonly DependencyProperty KeyboardProperty =
            DependencyProperty.Register("Keyboard", typeof (IKeyboard), typeof (KeyboardHost),
                new PropertyMetadata(default(IKeyboard),
                    (o, args) =>
                    {
                        var keyboardHost = o as KeyboardHost;
                        if (keyboardHost != null)
                        {
                            keyboardHost.GenerateContent();
                        }
                    }));

        public IKeyboard Keyboard
        {
            get { return (IKeyboard) GetValue(KeyboardProperty); }
            set { SetValue(KeyboardProperty, value); }
        }

        public static readonly DependencyProperty PointToKeyValueMapProperty =
            DependencyProperty.Register("PointToKeyValueMap", typeof(Dictionary<Rect, KeyValue>),
                typeof(KeyboardHost), new PropertyMetadata(default(Dictionary<Rect, KeyValue>)));

        public Dictionary<Rect, KeyValue> PointToKeyValueMap
        {
            get { return (Dictionary<Rect, KeyValue>)GetValue(PointToKeyValueMapProperty); }
            set { SetValue(PointToKeyValueMapProperty, value); }
        }

        public static readonly DependencyProperty ErrorContentProperty =
            DependencyProperty.Register("ErrorContent", typeof (object), typeof (KeyboardHost), new PropertyMetadata(default(object)));

        public object ErrorContent
        {
            get { return GetValue(ErrorContentProperty); }
            set { SetValue(ErrorContentProperty, value); }
        }

        #endregion

        #region OnLoaded - build key map

        private void OnLoaded(object sender, RoutedEventArgs routedEventArgs)
        {
            Log.Debug("KeyboardHost loaded.");

            BuildPointToKeyMap();

            SubscribeToSizeChanges();

            var parentWindow = Window.GetWindow(this);

            if (parentWindow == null)
            {
                var windowException = new ApplicationException(Properties.Resources.PARENT_WINDOW_COULD_NOT_BE_FOUND);

                Log.Error(windowException);

                throw windowException;
            }
            
            SubscribeToParentWindowMoves(parentWindow);
            SubscribeToParentWindowStateChanges(parentWindow);

            Loaded -= OnLoaded; //Ensure this logic only runs once
        }

        #endregion

        #region Generate Content

        private void GenerateContent()
        {
            Log.DebugFormat("GenerateContent called. Keyboard language is '{0}' and Keyboard type is '{1}'", 
                Settings.Default.KeyboardAndDictionaryLanguage, Keyboard != null ? Keyboard.GetType() : null);

            //Clear out point to key map
            PointToKeyValueMap = null;
          
            object newContent = ErrorContent;

            if (Keyboard is ViewModelKeyboards.Alpha1)
            {
                switch (Settings.Default.KeyboardAndDictionaryLanguage)
                {
                    case Languages.CatalanSpain:
                        newContent = new CatalanViews.Alpha1 { DataContext = Keyboard };
                        break;
                    case Languages.CroatianCroatia:
                        newContent = new CroatianViews.Alpha1 { DataContext = Keyboard };
                        break;
                    case Languages.CzechCzechRepublic:
                        newContent = new CzechViews.Alpha1 { DataContext = Keyboard };
                        break;
                    case Languages.DanishDenmark:
                        newContent = new DanishViews.Alpha1 { DataContext = Keyboard };
                        break;
                    case Languages.DutchBelgium:
                        newContent = new DutchViews.BelgiumAlpha { DataContext = Keyboard };
                        break;
                    case Languages.DutchNetherlands:
                        newContent = new DutchViews.NetherlandsAlpha { DataContext = Keyboard };
                        break;
                    case Languages.FrenchFrance:
                        newContent = new FrenchViews.Alpha1 { DataContext = Keyboard };
                        break;
                    case Languages.GermanGermany:
                        newContent = new GermanViews.Alpha1 { DataContext = Keyboard };
                        break;
                    case Languages.GreekGreece:
                        newContent = new GreekViews.Alpha1 { DataContext = Keyboard };
                        break;
                    case Languages.ItalianItaly:
                        newContent = new ItalianViews.Alpha1 { DataContext = Keyboard };
                        break;
                    case Languages.KoreanKorea:
                        newContent = new KoreanViews.Alpha1 { DataContext = Keyboard };
                        break;
                    case Languages.PortuguesePortugal:
                        newContent = new PortugueseViews.Alpha1 { DataContext = Keyboard };
                        break;
                    case Languages.RussianRussia:
                        newContent = new RussianViews.Alpha1 { DataContext = Keyboard };
                        break;
                    case Languages.SlovakSlovakia:
                        newContent = new SlovakViews.Alpha1 { DataContext = Keyboard };
                        break;
                    case Languages.SlovenianSlovenia:
                        newContent = new SlovenianViews.Alpha1 { DataContext = Keyboard };
                        break;
                    case Languages.SpanishSpain:
                        newContent = new SpanishViews.Alpha1 { DataContext = Keyboard };
                        break;
                    case Languages.TurkishTurkey:
                        newContent = new TurkishViews.Alpha1 { DataContext = Keyboard };
                        break;
                    default:
                        newContent = Settings.Default.UsingCommuniKateKeyboardLayout
                            ? (object)new EnglishViews.CommuniKate { DataContext = Keyboard }
                            : Settings.Default.UseSimplifiedKeyboardLayout
                            ? (object)new EnglishViews.SimplifiedAlpha1 { DataContext = Keyboard }
                            : Settings.Default.UseAlphabeticalKeyboardLayout 
                            ? (object)new EnglishViews.AlphabeticalAlpha1 { DataContext = Keyboard }
                            : new EnglishViews.Alpha1 { DataContext = Keyboard };
                        break;
                }
            }
            else if (Keyboard is ViewModelKeyboards.Alpha2)
            {
                switch (Settings.Default.KeyboardAndDictionaryLanguage)
                {
                    case Languages.KoreanKorea:
                        newContent = new KoreanViews.Alpha2 { DataContext = Keyboard };
                        break;
                }
            }
            else if (Keyboard is ViewModelKeyboards.ConversationAlpha1)
            {
                switch (Settings.Default.KeyboardAndDictionaryLanguage)
                {
                    case Languages.CatalanSpain:
                        newContent = new CatalanViews.ConversationAlpha1 { DataContext = Keyboard };
                        break;
                    case Languages.CroatianCroatia:
                        newContent = new CroatianViews.ConversationAlpha1 { DataContext = Keyboard };
                        break;
                    case Languages.CzechCzechRepublic:
                        newContent = new CzechViews.ConversationAlpha1 { DataContext = Keyboard };
                        break;
                    case Languages.DanishDenmark:
                        newContent = new DanishViews.ConversationAlpha1 { DataContext = Keyboard };
                        break;
                    case Languages.DutchBelgium:
                        newContent = new DutchViews.BelgiumConversationAlpha1 { DataContext = Keyboard };
                        break;
                    case Languages.DutchNetherlands:
                        newContent = new DutchViews.NetherlandsConversationAlpha1 { DataContext = Keyboard };
                        break;
                    case Languages.FrenchFrance:
                        newContent = new FrenchViews.ConversationAlpha1 { DataContext = Keyboard };
                        break;
                    case Languages.GermanGermany:
                        newContent = new GermanViews.ConversationAlpha1 { DataContext = Keyboard };
                        break;
                    case Languages.GreekGreece:
                        newContent = new GreekViews.ConversationAlpha1 { DataContext = Keyboard };
                        break;
                    case Languages.ItalianItaly:
                        newContent = new ItalianViews.ConversationAlpha1 { DataContext = Keyboard };
                        break;
                    case Languages.KoreanKorea:
                        newContent = new KoreanViews.ConversationAlpha1 { DataContext = Keyboard };
                        break;
                    case Languages.PortuguesePortugal:
                        newContent = new PortugueseViews.ConversationAlpha1 { DataContext = Keyboard };
                        break;
                    case Languages.RussianRussia:
                        newContent = new RussianViews.ConversationAlpha1 { DataContext = Keyboard };
                        break;
                    case Languages.SlovakSlovakia:
                        newContent = new SlovakViews.ConversationAlpha1 { DataContext = Keyboard };
                        break;
                    case Languages.SlovenianSlovenia:
                        newContent = new SlovenianViews.ConversationAlpha1 { DataContext = Keyboard };
                        break;
                    case Languages.SpanishSpain:
                        newContent = new SpanishViews.ConversationAlpha1 { DataContext = Keyboard };
                        break;
                    case Languages.TurkishTurkey:
                        newContent = new TurkishViews.ConversationAlpha1 { DataContext = Keyboard };
                        break;
                    default:
                        newContent = Settings.Default.UsingCommuniKateKeyboardLayout
                            ? (object)new EnglishViews.CommuniKate { DataContext = Keyboard }
                            : Settings.Default.UseSimplifiedKeyboardLayout
                            ? (object)new EnglishViews.SimplifiedConversationAlpha1 { DataContext = Keyboard }
                            : Settings.Default.UseAlphabeticalKeyboardLayout
                            ? (object)new EnglishViews.AlphabeticalConversationAlpha1 { DataContext = Keyboard }
                            : new EnglishViews.ConversationAlpha1 { DataContext = Keyboard };
                        break;
                }
            }
            else if (Keyboard is ViewModelKeyboards.ConversationAlpha2)
            {
                switch (Settings.Default.KeyboardAndDictionaryLanguage)
                {
                    case Languages.KoreanKorea:
                        newContent = new KoreanViews.ConversationAlpha2 { DataContext = Keyboard };
                        break;
                }
            }
            else if (Keyboard is ViewModelKeyboards.ConversationConfirm)
            {
                newContent = new CommonViews.ConversationConfirm { DataContext = Keyboard };
            }
            else if (Keyboard is ViewModelKeyboards.ConversationNumericAndSymbols)
            {
                newContent = new CommonViews.ConversationNumericAndSymbols { DataContext = Keyboard };
            }
            else if (Keyboard is ViewModelKeyboards.Currencies1)
            {
                newContent = new CommonViews.Currencies1 { DataContext = Keyboard };
            }
            else if (Keyboard is ViewModelKeyboards.Currencies2)
            {
                newContent = new CommonViews.Currencies2 { DataContext = Keyboard };
            }
            else if (Keyboard is ViewModelKeyboards.Diacritics1)
            {
                newContent = new CommonViews.Diacritics1 { DataContext = Keyboard };
            }
            else if (Keyboard is ViewModelKeyboards.Diacritics2)
            {
                newContent = new CommonViews.Diacritics2 { DataContext = Keyboard };
            }
            else if (Keyboard is ViewModelKeyboards.Diacritics3)
            {
                newContent = new CommonViews.Diacritics3 { DataContext = Keyboard };
            }
            else if (Keyboard is ViewModelKeyboards.Language)
            {
                newContent = new CommonViews.Language { DataContext = Keyboard };
            }
            else if (Keyboard is ViewModelKeyboards.Menu)
            {
                newContent = new CommonViews.Menu { DataContext = Keyboard };
            }
            else if (Keyboard is ViewModelKeyboards.Minimised)
            {
                newContent = new CommonViews.Minimised { DataContext = Keyboard };
            }
            else if (Keyboard is ViewModelKeyboards.Mouse)
            {
                newContent = new CommonViews.Mouse { DataContext = Keyboard };
            }
            else if (Keyboard is ViewModelKeyboards.NumericAndSymbols2)
            {
                newContent = new CommonViews.NumericAndSymbols2 { DataContext = Keyboard };
            }
            else if (Keyboard is ViewModelKeyboards.NumericAndSymbols3)
            {
                newContent = new CommonViews.NumericAndSymbols3 { DataContext = Keyboard };
            }
            else if (Keyboard is ViewModelKeyboards.NumericAndSymbols1)
            {
                newContent = new CommonViews.NumericAndSymbols1 { DataContext = Keyboard };
            }
            else if (Keyboard is ViewModelKeyboards.PhysicalKeys)
            {
                newContent = new CommonViews.PhysicalKeys { DataContext = Keyboard };
            }
            else if (Keyboard is ViewModelKeyboards.SizeAndPosition)
            {
                newContent = new CommonViews.SizeAndPosition { DataContext = Keyboard };
            }
            else if (Keyboard is ViewModelKeyboards.WebBrowsing)
            {
                newContent = new CommonViews.WebBrowsing { DataContext = Keyboard };
            }
            else if (Keyboard is ViewModelKeyboards.YesNoQuestion)
            {
                newContent = new CommonViews.YesNoQuestion { DataContext = Keyboard };
            }
            else if (Keyboard is ViewModelKeyboards.DynamicKeyboard)
            {
                var kb = Keyboard as ViewModelKeyboards.DynamicKeyboard;
                newContent = new CommonViews.DynamicKeyboard(kb.Link, kb.ResizeAction) { DataContext = Keyboard };                
            }
            else if (Keyboard is ViewModelKeyboards.DynamicKeyboardSelector)
            {
                var kb = Keyboard as ViewModelKeyboards.DynamicKeyboardSelector;
                newContent = new CommonViews.DynamicKeyboardSelector(kb.PageIndex) { DataContext = Keyboard };
            }
            Content = newContent;
        }

        #endregion
        
        #region Content Change Handler

        private static void ContentChangedHandler(object sender, EventArgs e)
        {
            var keyboardHost = sender as KeyboardHost;
            if (keyboardHost != null)
            {
                keyboardHost.BuildPointToKeyMap();
            }
        }

        private void OnMouseEnter(object sender, System.EventArgs e)
        {
            if (Settings.Default.PointsSource == PointsSources.MousePosition &&
                Settings.Default.PointsMousePositionHideCursor)
            {
                this.Cursor = System.Windows.Input.Cursors.None;
            }
            else
            {
                this.Cursor = System.Windows.Input.Cursors.Arrow;
            }
        }
        
        #endregion

        #region Build Point To Key Map

        private void BuildPointToKeyMap()
        {
            Log.Info("Building PointToKeyMap.");

            if (currentKeyboardKeyValueSubscriptions != null)
            {
                Log.Debug("Disposing of currentKeyboardKeyValueSubscriptions.");
                currentKeyboardKeyValueSubscriptions.Dispose();
            }
            currentKeyboardKeyValueSubscriptions = new CompositeDisposable();

            var contentAsFrameworkElement = Content as FrameworkElement;
            if (contentAsFrameworkElement != null)
            {
                if (contentAsFrameworkElement.IsLoaded)
                {
                    TraverseAllKeysAndBuildPointToKeyValueMap();
                }
                else
                {
                    RoutedEventHandler loaded = null;
                    loaded = (sender, args) =>
                    {
                        TraverseAllKeysAndBuildPointToKeyValueMap();
                        contentAsFrameworkElement.Loaded -= loaded;
                    };
                    contentAsFrameworkElement.Loaded += loaded;
                }
            }
        }

        private void TraverseAllKeysAndBuildPointToKeyValueMap()
        {
            var allKeys = VisualAndLogicalTreeHelper.FindVisualChildren<Key>(this).ToList();

            var pointToKeyValueMap = new Dictionary<Rect, KeyValue>();

            var topLeftPoint = new Point(0, 0);

            foreach (var key in allKeys)
            {
                if (key.IsVisible
                    && PresentationSource.FromVisual(key) != null
                    && key.Value != null
                    && key.Value.HasContent())
                {
                    var rect = new Rect
                    {
                        Location = key.PointToScreen(topLeftPoint),
                        Size = (Size) key.GetTransformToDevice().Transform((Vector) key.RenderSize)
                    };

                    if (rect.Size.Width != 0 && rect.Size.Height != 0)
                    {
                        if (pointToKeyValueMap.ContainsKey(rect))
                        {
                            // In Release, just log error
                            KeyValue existingKeyValue = pointToKeyValueMap[rect];
                            Log.ErrorFormat("Overlapping keys {0} and {1}, cannot add {1} to map", existingKeyValue, key.Value);

                            Debug.Assert(!pointToKeyValueMap.ContainsKey(rect));
                        }
                        else
                        {
                            pointToKeyValueMap.Add(rect, key.Value);
                        }
                    }

                    var keyValueChangedSubscription = key.OnPropertyChanges<KeyValue>(Key.ValueProperty).Subscribe(kv =>
                    {
                        KeyValue mapValue;
                        if (pointToKeyValueMap.TryGetValue(rect, out mapValue))
                        {
                            pointToKeyValueMap[rect] = kv;
                        }
                    });
                    currentKeyboardKeyValueSubscriptions.Add(keyValueChangedSubscription);
                }
            }

            Log.InfoFormat("PointToKeyValueMap rebuilt with {0} keys.", pointToKeyValueMap.Keys.Count);
            PointToKeyValueMap = pointToKeyValueMap;
        }

        #endregion

        #region Subscribe To Size Changes

        private void SubscribeToSizeChanges()
        {
            Observable.FromEventPattern<SizeChangedEventHandler, SizeChangedEventArgs>
                (h => SizeChanged += h,
                h => SizeChanged -= h)
                .Throttle(TimeSpan.FromSeconds(0.1))
                .ObserveOnDispatcher()
                .Subscribe(ep =>
                {
                    Log.InfoFormat("SizeChanged event detected from {0} to {1}.", ep.EventArgs.PreviousSize, ep.EventArgs.NewSize);
                    BuildPointToKeyMap();
                });
        }

        #endregion

        #region Subscribe To Parent Window Moves

        private void SubscribeToParentWindowMoves(Window parentWindow)
        {
            Observable.FromEventPattern<EventHandler, EventArgs>
                (h => parentWindow.LocationChanged += h,
                h => parentWindow.LocationChanged -= h)
                .Throttle(TimeSpan.FromSeconds(0.1))
                .ObserveOnDispatcher()
                .Subscribe(_ =>
                {
                    Log.Info("Window's LocationChanged event detected.");
                    BuildPointToKeyMap();
                });
        }

        #endregion

        #region Subscribe To Parent Window State Changes

        private void SubscribeToParentWindowStateChanges(Window parentWindow)
        {
            Observable.FromEventPattern<EventHandler, EventArgs>
                (h => parentWindow.StateChanged += h,
                h => parentWindow.StateChanged -= h)
                .Throttle(TimeSpan.FromSeconds(0.1))
                .ObserveOnDispatcher()
                .Subscribe(_ =>
                {
                    Log.InfoFormat("Window's StateChange event detected. New state: {0}.", parentWindow.WindowState);
                    BuildPointToKeyMap();
                });
        }

        #endregion
    }
}