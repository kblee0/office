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

C:\home\local\ImageMagick\magick ^
  -size 32x32 canvas:none ^
  -fill white ^
  -draw "roundrectangle 0,0 31,31 3,3" ^
  -fill "#808080" ^
  -draw "roundrectangle 2,2 29,29 1,1" ^
  -fill "#FFFFFF" ^
  -draw "rectangle 4,4 27,27" ^
  -fill "#ABABAB" -draw "rectangle 4,4 9,9" ^
  -fill "#ABABAB" -draw "rectangle 16,4 21,9" ^
  -fill "#ABABAB" -draw "rectangle 10,10 15,15" ^
  -fill "#ABABAB" -draw "rectangle 22,10 27,15" ^
  -fill "#ABABAB" -draw "rectangle 4,16 9,21" ^
  -fill "#ABABAB" -draw "rectangle 16,16 21,21" ^
  -fill "#ABABAB" -draw "rectangle 10,22 15,27" ^
  -fill "#ABABAB" -draw "rectangle 22,22 27,27" ^
  "color\CNONE.png"
