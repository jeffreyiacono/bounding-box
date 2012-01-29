## findBoundingBox ##
`findBoundingBox` macro takes _range of x coordinates_, _range of y coordinates_, and _range of id coordinates_ as parameters.
The macro will create a new worksheet that displays the points of a bounding shape that encompasses all points from the passed (x,y) params.

A picture is worth a thousand words:

![Bounding Box](https://lh5.googleusercontent.com/-86O1LLJqH0s/TyXMSR_Wz1I/AAAAAAAAAfU/H_NZQW_aQkI/s903/bounded.png)

_red series found by using `findBoundingBox` macro - sample app can be downloaded at sample/bounding-box.xlsm_

This is a quick and dirty implementation and has a lot of room for code modularization.

## Basic Usage ##
In a workbook, import __modules/bounding_box.bas__.
You can also import __forms/formCalcBoundingBox.frm__ and __forms/formCalcBoundingBox.frx__ for setting params and running `findBoundingBox` via a UI.
To launch the UI form, run the macro `showForm` (`Alt + L + PM` to launch macro catalogue) and select the desired x, y, and id ranges using the form's controls.
Finally, click "Find!"

A new worksheet will be added that specifies the set of points that will create the bounding shape.

## Todo ##
_Please feel free push any code that implements the following - patches will be happily
accepted_

* Better code modularization
* Better documentation
* Change name to "bounding shape" ... technically not a box :)

Any others come to mind? Email me at [jeff@elegantbuild.com](mailto:jeff@elegantbuild.com).

##MIT License

Copyright (c) 2012 ElegantBuild, LLC, http://elegantbuild.com/

Permission is hereby granted, free of charge, to any person obtaining
a copy of this software and associated documentation files (the
"Software"), to deal in the Software without restriction, including
without limitation the rights to use, copy, modify, merge, publish,
distribute, sublicense, and/or sell copies of the Software, and to
permit persons to whom the Software is furnished to do so, subject to
the following conditions:

The above copyright notice and this permission notice shall be
included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
