del color\*.png

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
C:\home\local\ImageMagick\magick ^
  -size 32x32 canvas:none ^
  -fill white ^
  -draw "roundrectangle 0,0 31,31 3,3" ^
  -fill "#808080" ^
  -draw "roundrectangle 2,2 29,29 1,1" ^
  -fill "#%%C" ^
  -draw "rectangle 4,4 27,27" ^
  "color\C%%C.png"
)

copy CNONE.png color\