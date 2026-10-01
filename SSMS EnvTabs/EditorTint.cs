using System;

namespace SSMS_EnvTabs
{
    /// <summary>
    /// Color math for the gentle editor background tint. Kept free of WPF types so it can be unit tested.
    /// </summary>
    internal static class EditorTint
    {
        internal const int DefaultStrength = 8;
        internal const int MinStrength = 1;
        internal const int MaxStrength = 30;

        internal static int ClampStrength(int strengthPercent)
        {
            if (strengthPercent < MinStrength) return MinStrength;
            if (strengthPercent > MaxStrength) return MaxStrength;
            return strengthPercent;
        }

        /// <summary>
        /// Blends <paramref name="strengthPercent"/> percent of the tint over the base color.
        /// Colors are 0xRRGGBB integers.
        /// </summary>
        internal static int Blend(int baseRgb, int tintRgb, int strengthPercent)
        {
            double amount = ClampStrength(strengthPercent) / 100.0;
            return (BlendChannel(baseRgb >> 16, tintRgb >> 16, amount) << 16)
                | (BlendChannel(baseRgb >> 8, tintRgb >> 8, amount) << 8)
                | BlendChannel(baseRgb, tintRgb, amount);
        }

        private static int BlendChannel(int baseChannel, int tintChannel, double amount)
        {
            int b = baseChannel & 0xFF;
            int t = tintChannel & 0xFF;
            int value = (int)Math.Round((b * (1.0 - amount)) + (t * amount));
            return Math.Max(0, Math.Min(255, value));
        }
    }
}
