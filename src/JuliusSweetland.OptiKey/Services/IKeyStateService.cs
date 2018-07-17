﻿using System;
using System.ComponentModel;
using JuliusSweetland.OptiKey.Enums;
using JuliusSweetland.OptiKey.Models;

namespace JuliusSweetland.OptiKey.Services
{
    public interface IKeyStateService : INotifyPropertyChanged
    {
        bool SimulateKeyStrokes { get; set; }
        NotifyingConcurrentDictionary<KeyValue, double> KeySelectionProgress { get; }
        NotifyingConcurrentDictionary<KeyValue, KeyDownStates> KeyDownStates { get; }
        KeyEnabledStates KeyEnabledStates { get; }

        void ProgressKeyDownState(KeyValue keyValue);
    }
}
