param(
  [string]$LogoPath,
  [string]$OutPath
)

Add-Type -ReferencedAssemblies System.Drawing -TypeDefinition @'
using System;
using System.IO;
using System.Collections.Generic;

// Assemble un GIF anime a partir de plusieurs GIF mono-frame produits par GDI+.
// Chaque frame conserve sa propre palette (table de couleurs LOCALE), et on
// injecte l'extension NETSCAPE (boucle infinie) + un delai par frame.
public static class GifMaker
{
    public static void Save(string path, List<byte[]> frames, int delayCs)
    {
        using (var fs = new FileStream(path, FileMode.Create))
        {
            byte[] first = frames[0];
            fs.Write(first, 0, 6);   // "GIF89a"
            fs.Write(first, 6, 7);   // Logical Screen Descriptor

            int packed = first[10];
            int gct = ((packed & 0x80) != 0) ? 3 * (1 << ((packed & 0x07) + 1)) : 0;
            if (gct > 0) fs.Write(first, 13, gct);

            byte[] loop = {0x21,0xFF,0x0B,0x4E,0x45,0x54,0x53,0x43,0x41,0x50,0x45,
                           0x32,0x2E,0x30,0x03,0x01,0x00,0x00,0x00};
            fs.Write(loop, 0, loop.Length);

            foreach (var fr in frames)
            {
                int p = fr[10];
                int fgct = ((p & 0x80) != 0) ? 3 * (1 << ((p & 0x07) + 1)) : 0;
                int idx = 13 + fgct;

                while (fr[idx] != 0x2C)
                {
                    if (fr[idx] == 0x21) { idx += 2; while (fr[idx] != 0x00) idx += fr[idx] + 1; idx += 1; }
                    else if (fr[idx] == 0x3B) break;
                    else idx++;
                }

                fs.WriteByte(0x21); fs.WriteByte(0xF9); fs.WriteByte(0x04);
                fs.WriteByte(0x00);
                fs.WriteByte((byte)(delayCs & 0xFF)); fs.WriteByte((byte)((delayCs >> 8) & 0xFF));
                fs.WriteByte(0x00); fs.WriteByte(0x00);

                byte[] id = new byte[10];
                Array.Copy(fr, idx, id, 0, 10);
                if (fgct > 0) id[9] = (byte)(0x80 | (id[9] & 0x40) | (p & 0x07));
                fs.Write(id, 0, 10);
                if (fgct > 0) fs.Write(fr, 13, fgct);

                int dataStart = idx + 10;
                int dataEnd = fr.Length - 1;
                while (dataEnd > dataStart && fr[dataEnd] != 0x3B) dataEnd--;
                fs.Write(fr, dataStart, dataEnd - dataStart);
            }

            fs.WriteByte(0x3B);
        }
    }
}
'@

Add-Type -AssemblyName System.Drawing

$W = 520; $H = 300
$frames = 24
$delayCs = 4

$col_bg    = [System.Drawing.Color]::White
$col_track = [System.Drawing.Color]::FromArgb(0xE7,0xE9,0xF0)
$col_seg   = [System.Drawing.Color]::FromArgb(0xE8,0x9B,0x2F)   # orange ASERTI/Metrologo
$col_text  = [System.Drawing.Color]::FromArgb(0x1E,0x2D,0x4F)   # navy ASERTI

$logo = [System.Drawing.Image]::FromFile($LogoPath)
$lw = 380
$lh = [int]($logo.Height * $lw / $logo.Width)
$lx = [int](($W - $lw) / 2)
$ly = 52

$font = New-Object System.Drawing.Font("Segoe UI", 13, [System.Drawing.FontStyle]::Regular)
$txt  = "Installation de Metrologo..."

$tx = 60; $ty = 240; $tw = 400; $th = 10; $r = 5
$segW = 130
$travel = $tw + $segW

function New-RoundRect([int]$x,[int]$y,[int]$w,[int]$h,[int]$rad) {
  $gp = New-Object System.Drawing.Drawing2D.GraphicsPath
  $d = $rad * 2
  $gp.AddArc($x, $y, $d, $d, 180, 90)
  $gp.AddArc($x + $w - $d, $y, $d, $d, 270, 90)
  $gp.AddArc($x + $w - $d, $y + $h - $d, $d, $d, 0, 90)
  $gp.AddArc($x, $y + $h - $d, $d, $d, 90, 90)
  $gp.CloseFigure()
  return $gp
}

$list = New-Object 'System.Collections.Generic.List[byte[]]'
$trackBrush = New-Object System.Drawing.SolidBrush($col_track)
$segBrush   = New-Object System.Drawing.SolidBrush($col_seg)
$textBrush  = New-Object System.Drawing.SolidBrush($col_text)

for ($i = 0; $i -lt $frames; $i++) {
  $bmp = New-Object System.Drawing.Bitmap($W, $H)
  $g = [System.Drawing.Graphics]::FromImage($bmp)
  $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
  $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
  $g.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::ClearTypeGridFit
  $g.Clear($col_bg)

  $g.DrawImage($logo, $lx, $ly, $lw, $lh)

  $sz = $g.MeasureString($txt, $font)
  $g.DrawString($txt, $font, $textBrush, ($W - $sz.Width) / 2, 178)

  $track = New-RoundRect $tx $ty $tw $th $r
  $g.FillPath($trackBrush, $track)

  $pos = $tx - $segW + ($travel * $i / $frames)
  $g.SetClip($track)
  foreach ($off in @(-$travel, 0, $travel)) {
    $g.FillRectangle($segBrush, [float]($pos + $off), [float]$ty, [float]$segW, [float]$th)
  }
  $g.ResetClip()
  $track.Dispose()

  $g.Dispose()
  $ms = New-Object System.IO.MemoryStream
  $bmp.Save($ms, [System.Drawing.Imaging.ImageFormat]::Gif)
  $list.Add($ms.ToArray())
  $ms.Dispose()
  $bmp.Dispose()
}

[GifMaker]::Save($OutPath, $list, $delayCs)
$logo.Dispose()

$chk = [System.Drawing.Image]::FromFile($OutPath)
$fd = New-Object System.Drawing.Imaging.FrameDimension($chk.FrameDimensionsList[0])
$n = $chk.GetFrameCount($fd)
$chk.Dispose()
$fi = Get-Item $OutPath
Write-Output ("GIF genere : {0}  ({1} frames, {2} octets)" -f $OutPath, $n, $fi.Length)
