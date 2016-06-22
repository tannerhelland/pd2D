# pd2D 0.1 (pre-alpha)

### pd2D is a high-performance 2D graphics library for classic VB (VB 6.0)

pd2D is derived from the 2D graphics library that powers the PhotoDemon open-source photo editor (https://github.com/tannerhelland/PhotoDemon).  It is designed to solve some long-running problems with 2D imaging in classic VB applications.

### pd2D is 100% open-source

The full source code of pd2D is available, and its liberal BSD license allows its use in any project, commercial or otherwise.  You are also free to fork it and/or make your own modifications.  (Of course, I always appreciate it if updates are shared with the main project, so we can pass along new features and bug-fixes to everyone!)

### pd2D has no external dependencies

The library's default backend leans on standard libraries available all the way back to Windows XP.  Windows XP through current Windows 10 builds are fully supported.

### pd2D is designed with multiple backends in mind

A default pd2D build uses a mixture of GDI, GDI+, and custom code, but it can easily be extended to support other backends.  For example, you can wrap pd2D around a 2D library like Cairo with minimal effort.  Suggestions for alternate backends (or even better, pull requests!) are always welcome.

### pd2D does not require you to be a graphics expert

Unlike a bare type library, pd2D manages things like memory allocations and handle disposal for you.  If you know how to use VB6 classes, you know everything you need to use pd2D.

### pd2D is fast

Wherever possible, pd2D leans on hardware-accelerated rendering.  Performance is best under Windows 10, but some measure of hardware acceleration is available all the way back to Windows XP.

## How can I get involved? 
pd2D is maintained by a single individual with a family to support.  The software is provided free-of-charge under a permissive open-source license, and no fees or money will ever be charged for its use.

That said, donations go a long way toward supporting its development.  If you would like to donate and support development, you can donate through pd2D's parent project website:

http://photodemon.org/donate/

If you can't contribute monetarily to pd2D, here are other ways to help:
* Let me know if you find any bugs. Issues can be submitted via pd2D's official bug tracker: https://github.com/tannerhelland/pd2D/issues
* Pull requests (for bug-fixes, new features, new backends, documentation - anything!) are always welcome.
* Do you wish pd2D behaved differently?  Do you want it to offer a certain feature?  Suggestions and feedback are always welcome, and they can be submitted through the same bug tracker mentioned earlier: https://github.com/tannerhelland/pd2D/issues

## How is pd2D licensed?

pd2D is available under a permissive BSD license.  See the LICENSE.MD file for specific details.

## How do I use pd2D?

This repository is divided into two parts: the bare pd2D source (available in the /pd2D folder), and a sample project (available in the /samples folder).

If you just want to see some sample code and a demonstration of features, the pd2D.vbg project in the root folder is a good place to start.