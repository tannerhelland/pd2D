## pd2D Raster Feature Overview

[Raster graphics](https://en.wikipedia.org/wiki/Raster_graphics) are pictures comprised of a grid of pixels.  Popular raster image formats include JPEG, GIF, PNG, and BMP.

pd2D supports reading and writing raster image files in the following formats:

* Windows Bitmap (bmp)
* GIF (gif)
* JPEG (jpg, jpeg)
* PNG (png)
* TIFF (tif, tiff)

pd2D also supports drawing atop raster images, and affine transformations like resizing, rotating, and skewing.  It also supports painting raster images atop each other (like layers in Photoshop).

All of these operations can be performed in real-time, even with large images.

The included sample project demonstrates some of these features.