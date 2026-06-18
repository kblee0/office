for %%C in (
2F2F2F
E6B4CF
ED1C6F
959595
40B79B
27AE60
F39C12
16A085
8E44AD
3498DB
E74C3C
000000
) do (
    C:\home\local\ImageMagick\magick -size 32x32 canvas:none ^
      -fill "#%%C" ^
      -stroke white ^
      -strokewidth 2 ^
      -draw "roundrectangle 2,2 30,30 3,3" ^
      "color\C%%C.png"
)
