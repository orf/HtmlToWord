HtmlToWord
===

## This project has moved to [WordInserter](https://github.com/orf/wordinserter), supporting both Markdown and HTML, with lots more HTML supported. This library is currently unmaintained and out of date!

This module was created for use in an application that uses Word to generate reports based on HTML input from a web based WYSIWYG editor. You can use it like so:

```python
import HtmlToWord, win32com.client

# This opens Microsoft Word, ready for automation
word = win32com.client.gencache.EnsureDispatch("Word.Application")
word.Visible = True # Don't set this to True in production!
# Create a new document and our Parser class
document = word.Documents.Add()
parser = HtmlToWord.Parser()

Html = """
<h3>This is a title</h3>
<p><img src="http://placehold.it/150x150" alt="I go below the image as a caption"></p>
<p><i>This is <b>some</b> text</i> in a <a href="http://google.com">paragraph</a></p>
<ul>
    <li>Boo! I am a <b>list</b></li>
</ul>
"""
# Feed in our HTML, passing the Word instance and a Selection object. 
parser.ParseAndRender(Html, word, document.ActiveWindow.Selection)
```

This will insert a representation of that HTML into the word document, including the image, caption and list.


### Install
Get it [from PyPi here](https://pypi.python.org/pypi/HtmlToWord). It is only tested with Word 2010, but it may still work with 2007.

### Limitations
Its not perfect, but it works. This is created for use with [Redactor](http://imperavi.com/redactor/) as the HTML it generates is very clean and fairly simple to parse. Other editors may vary.

## Supported tags and extensions

HtmlToWord currently supports the following HTML tags:

#### Typography
 * p
 * b / strong
 * br
 * em
 * i
 * u
 * h1/2/3/4
 * pre

#### Tables
 * table
 * tbody
 * thead
 * tr
 * td

#### Lists
 * ul
 * ol
 * li

#### Misc
 * div
 * img
 * a


### Extending
Extending HtmlToWord is pretty easy. Each tag is a class that inherits from BaseElement. It has two methods that are called: *StartRender* and *EndRender*. Take a look in elements/headings.py and elements/text.py for some simple examples.

#### Rendering hooks / Custom styles
The Parser class has three callbacks: preRender, Render and postRender, which are called before, during and after an element is rendered.
You can use these callbacks to modify and elements style post-rendering, for example to change all tables to a set custom style you can do the following (e is the Element instance)

```python
from HtmlToWord.elements.Table import Table
from HtmlToWord.elements.Base import BaseElement
from HtmlToWord.elements.Text import Pre

# List of styles: http://msdn.microsoft.com/en-us/library/office/ff835210(v=office.14).aspx
parser.AddPostRenderCallback(Table, lambda e: setattr(e.Table, "Style", constants.wdSomeTableStyleHere))
parser.AddPostRenderCallback(BaseElement, lambda e: print("This is called for every element"))
# Gives all pre tags a grey-ish background. For use with code.
parser.AddRenderCallback(Pre, lambda e: setattr(e.selection.Shading, "BackgroundPatternColor", -603923969))
```

Callbacks use isinstance to check, which means a callback on a parent class will call for all of the child classes.

## Rationale
#### Why Word? Why not ODF or OpenOffice?
Time. Words Object Model is [very well documented](http://msdn.microsoft.com/en-us/library/ff837519) with lots of samples available on the internet - any .NET code in VB or C# can be translated pretty easily. On top of this you can record Macro's within Word that generates Visual Basic code while you play with a document, meaning its very quick to find out how to do things. ODF looks cool, but again i'm not getting any younger and word's COM interface ticked all the boxes. In the future I might expand this module to generate ODF XML, but for now its a pipe dream.


[![Bitdeli Badge](https://d2weczhvl823v0.cloudfront.net/orf/htmltoword/trend.png)](https://bitdeli.com/free "Bitdeli Badge")

