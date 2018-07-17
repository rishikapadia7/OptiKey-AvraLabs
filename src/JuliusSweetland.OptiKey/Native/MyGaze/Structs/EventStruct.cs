﻿using System;

namespace JuliusSweetland.OptiKey.Native.MyGaze.Structs
{
    public struct EventStruct
    {
        public char eventType;
        public char eye;
        public Int64 startTime;
        public Int64 endTime;
        public Int64 duration;
        public double positionX;
        public double positionY;
    };
}
