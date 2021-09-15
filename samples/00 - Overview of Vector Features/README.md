## pd2D Vector Feature Overview

[Vector graphics](https://en.wikipedia.org/wiki/Vector_graphics) are images comprised of mathematical formulas, like lines, curves, and polygons.  These graphics are commonly drawn at run-time using simple commands (like VB6's .Line and .Circle commands), and they can also be persisted to file formats like WMF or EMF.

pd2D supports a huge variety of vector drawing commands, including lines, arcs, polygons, and more.  These shapes can be antialiased or filled, and they can be transformed in a wide variety of ways (rotation, resize, etc).

Vector drawing commands can be freely intermixed with raster drawing, so you can easily perform tasks like loading a PNG, drawing arrows onto it, then saving it as a JPEG.

All of these operations can be performed in real-time, even with large images and/or complex vector commands.

The included sample project demonstrates some of these features.