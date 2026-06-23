for %%C in (
FFFFFF
EAEAEA
A6A6A6
5A626A
2B3A4A
231815
FCE6E6
E60012
850107
4A90E2
3B7A57
E09924
) do (
    C:\home\local\ImageMagick\magick -size 32x32 canvas:none ^
      -fill "#%%C" ^
      -stroke white ^
      -strokewidth 2 ^
      -draw "roundrectangle 2,2 30,30 3,3" ^
      "color\C%%C.png"
)
