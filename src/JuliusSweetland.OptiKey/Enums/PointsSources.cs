using JuliusSweetland.OptiKey.Properties;
namespace JuliusSweetland.OptiKey.Enums
{
    public enum PointsSources
    {
        Alienware17,
        GazeTracker,
        MousePosition,
        SteelseriesSentry,
        TheEyeTribe,
        TobiiEyeTracker4C,
        TobiiEyeX,
        TobiiRex,
        TobiiPcEyeGo,
        TobiiPcEyeGoPlus,
        TobiiPcEyeMini,
        TobiiX2_30,
        TobiiX2_60,
        VisualInteractionMyGaze
    }

    public static partial class EnumExtensions
    {
        public static string ToDescription(this PointsSources pointSource)
        {
            switch (pointSource)
            {
                case PointsSources.Alienware17: return Resources.ALIENWARE_17;
                case PointsSources.GazeTracker: return Resources.GAZE_TRACKER;
                case PointsSources.MousePosition: return Resources.MOUSE_POSITION;
                case PointsSources.SteelseriesSentry: return Resources.STEELSERIES_SENTRY;
                case PointsSources.TheEyeTribe: return Resources.THE_EYE_TRIBE;
                case PointsSources.TobiiEyeTracker4C: return Resources.TOBII_EYE_TRACKER_4C;
                case PointsSources.TobiiEyeX: return Resources.TOBII_EYEX;
                case PointsSources.TobiiRex: return Resources.TOBII_REX;
                case PointsSources.TobiiPcEyeGo: return Resources.TOBII_PCEYE_GO;
                case PointsSources.TobiiPcEyeGoPlus: return Resources.TOBII_PCEYE_GO_PLUS;
                case PointsSources.TobiiPcEyeMini: return Resources.TOBII_PCEYE_MINI;
                case PointsSources.TobiiX2_30: return Resources.TOBII_X2_30;
                case PointsSources.TobiiX2_60: return Resources.TOBII_X2_60;
                case PointsSources.VisualInteractionMyGaze: return Resources.VI_MYGAZE;
            }

            return pointSource.ToString();
        }
    }
}
