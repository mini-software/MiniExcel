﻿using System;

namespace MiniExcelLibs.Utils
{
    internal static class ImageHelper
    {
        public enum ImageFormat
        {
            bmp,
            jpg,
            gif,
            tiff,
            png,
            unknown
        }

#if NET45||NETSTANDARD2_0
        public static ImageFormat GetImageFormat(byte[] bytes)
        {
            var bmp = new byte[] { (byte)'B', (byte)'M' };            // BMP
            var gif = new byte[] { (byte)'G', (byte)'I', (byte)'F' }; // GIF
            var png = new byte[] { 137, 80, 78, 71 };                 // PNG
            var tiff = new byte[] { 73, 73, 42 };                     // TIFF
            var tiff2 = new byte[] { 77, 77, 42 };                    // TIFF
            var jpeg = new byte[] { 255, 216, 255, 224 };             // jpeg
            var jpeg2 = new byte[] { 255, 216, 255, 225 };            // jpeg canon

            if (bytes.StartsWith(bmp))
                return ImageFormat.bmp;

            if (bytes.StartsWith(gif))
                return ImageFormat.gif;

            if (bytes.StartsWith(png))
                return ImageFormat.png;

            if (bytes.StartsWith(tiff))
                return ImageFormat.tiff;

            if (bytes.StartsWith(tiff2))
                return ImageFormat.tiff;

            if (bytes.StartsWith(jpeg))
                return ImageFormat.jpg;

            if (bytes.StartsWith(jpeg2))
                return ImageFormat.jpg;

            return ImageFormat.unknown;
        }
        
#elif NET5_0_OR_GREATER
        public static ImageFormat GetImageFormat(ReadOnlySpan<byte> bytes)
        {
            ReadOnlySpan<byte> bmp = stackalloc byte[] { (byte)'B', (byte)'M' };            // BMP
            ReadOnlySpan<byte> gif = stackalloc byte[] { (byte)'G', (byte)'I', (byte)'F' }; // GIF
            ReadOnlySpan<byte> png = stackalloc byte[] { 137, 80, 78, 71 };                 // PNG
            ReadOnlySpan<byte> tiff = stackalloc byte[] { 73, 73, 42 };                     // TIFF
            ReadOnlySpan<byte> tiff2 = stackalloc byte[] { 77, 77, 42 };                    // TIFF
            ReadOnlySpan<byte> jpeg = stackalloc byte[] { 255, 216, 255, 224 };             // jpeg
            ReadOnlySpan<byte> jpeg2 = stackalloc byte[] { 255, 216, 255, 225 };            // jpeg canon

            if (bytes.StartsWith(bmp))
                return ImageFormat.bmp;

            if (bytes.StartsWith(gif))
                return ImageFormat.gif;

            if (bytes.StartsWith(png))
                return ImageFormat.png;

            if (bytes.StartsWith(tiff))
                return ImageFormat.tiff;

            if (bytes.StartsWith(tiff2))
                return ImageFormat.tiff;

            if (bytes.StartsWith(jpeg))
                return ImageFormat.jpg;

            if (bytes.StartsWith(jpeg2))
                return ImageFormat.jpg;

            return ImageFormat.unknown;
        }
#endif
    }
}
