﻿using System;
using System.Drawing;

namespace Sharepoint.Helper
{

    public class ColorHelper
    {
        ColorHelper() { }

        class Nested
        {
            static Nested()
            {
            }
            internal static readonly ColorHelper instance = new ColorHelper();
        }

        public static ColorHelper Instance
        {
            get
            {
                return Nested.instance;
            }
        }

        public Color FromHtml(string HtmlColor) 
        {
            return ColorTranslator.FromHtml(HtmlColor);
        }
        public String ToHtml(Color ColorARGB)
        {
            return ColorTranslator.ToHtml(ColorARGB);
        }
        public Color FromHsl(int alpha, float hue, float saturation, float lighting)
        {
            if (0 > alpha || 255 < alpha)
            {
                throw new ArgumentOutOfRangeException("alpha");
            }
            if (0f > hue || 360f < hue)
            {
                throw new ArgumentOutOfRangeException("hue");
            }
            if (0f > saturation || 1f < saturation)
            {
                throw new ArgumentOutOfRangeException("saturation");
            }
            if (0f > lighting || 1f < lighting)
            {
                throw new ArgumentOutOfRangeException("lighting");
            }

            if (0 == saturation)
            {
                return Color.FromArgb(alpha, Convert.ToInt32(lighting * 255), Convert.ToInt32(lighting * 255), Convert.ToInt32(lighting * 255));
            }

            float fMax, fMid, fMin;
            int iSextant, iMax, iMid, iMin;

            if (0.5 < lighting)
            {
                fMax = lighting - (lighting * saturation) + saturation;
                fMin = lighting + (lighting * saturation) - saturation;
            }
            else
            {
                fMax = lighting + (lighting * saturation);
                fMin = lighting - (lighting * saturation);
            }

            iSextant = (int)Math.Floor(hue / 60f);
            if (300f <= hue)
            {
                hue -= 360f;
            }
            hue /= 60f;
            hue -= 2f * (float)Math.Floor(((iSextant + 1f) % 6f) / 2f);
            if (0 == iSextant % 2)
            {
                fMid = hue * (fMax - fMin) + fMin;
            }
            else
            {
                fMid = fMin - hue * (fMax - fMin);
            }

            iMax = Convert.ToInt32(fMax * 255);
            iMid = Convert.ToInt32(fMid * 255);
            iMin = Convert.ToInt32(fMin * 255);

            switch (iSextant)
            {
                case 1:
                    return Color.FromArgb(alpha, iMid, iMax, iMin);
                case 2:
                    return Color.FromArgb(alpha, iMin, iMax, iMid);
                case 3:
                    return Color.FromArgb(alpha, iMin, iMid, iMax);
                case 4:
                    return Color.FromArgb(alpha, iMid, iMin, iMax);
                case 5:
                    return Color.FromArgb(alpha, iMax, iMin, iMid);
                default:
                    return Color.FromArgb(alpha, iMax, iMid, iMin);
            }
        }

    }

    public static class ColorExtensions
    {
        public static string Lighten(string ColorBase, float percent)
        {
            ColorHelper _color = ColorHelper.Instance;

            Color color = System.Drawing.ColorTranslator.FromHtml(ColorBase);

            var lighting = color.GetBrightness();
            lighting = lighting + lighting * percent;
            if (lighting > 1.0)
            {
                lighting = 1;
            }
            else if (lighting <= 0)
            {
                lighting = 0.1f;
            }
            var tintedColor = _color.FromHsl(color.A, color.GetHue(), color.GetSaturation(), lighting);
            string strColor = System.Drawing.ColorTranslator.ToHtml(tintedColor);

            return strColor;
        }

        public static string Darken(string ColorBase, float percent)
        {
            ColorHelper _color = ColorHelper.Instance;
            Color color = System.Drawing.ColorTranslator.FromHtml(ColorBase);

            var lighting = color.GetBrightness();
            lighting = lighting - lighting * percent;
            if (lighting > 1.0)
            {
                lighting = 1;
            }
            else if (lighting <= 0)
            {
                lighting = 0;
            }
            var tintedColor = _color.FromHsl(color.A, color.GetHue(), color.GetSaturation(), lighting);
            string strColor = System.Drawing.ColorTranslator.ToHtml(tintedColor);

            return strColor;
        }

    }

}