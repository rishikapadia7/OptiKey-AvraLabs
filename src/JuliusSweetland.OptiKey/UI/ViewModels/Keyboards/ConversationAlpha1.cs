﻿using System;
using JuliusSweetland.OptiKey.UI.ViewModels.Keyboards.Base;

namespace JuliusSweetland.OptiKey.UI.ViewModels.Keyboards
{
    public class ConversationAlpha1 : BackActionKeyboard, IConversationKeyboard
    {
        public ConversationAlpha1(Action backAction)
            : base(backAction, simulateKeyStrokes: false, multiKeySelectionSupported: true)
        {
        }
    }
}
