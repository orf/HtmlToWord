HtmlToWord
===
This module was created for use in an application that uses Word to generate reports based on HTML input from a web frontend. You can use it like so:

```python
import HtmlToWord, win32com.client

word = win32com.client.gencache.EnsureDispatch("Word.Application")
word.Visible = True # Don't set this to True in production!
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
parser.ParseAndRender(Html, word, document.ActiveWindow.Selection)
```
Or if you don't want to use HTML you can create a tree of tags yourself (List elements omitted):
```python
from HtmlToWord.elements import *
parser.Render(word, [
   Heading3([Text("This is a title")]),
   Paragraph([Image(attributes={"src":"http://placehold.it/150x150","alt":"I go below"})]),
   Paragraph([
      Italic([Text("This is "), Bold([Text("some")]), Text(" text")]),
      Text(" in a "),
      HyperLink([Text("paragraph")], {"href":"http://google.com"}),
   ])
], document.ActiveWindow.Selection)
```

## Supported tags and extentions

HtmlToWord currently supports the following HTML tags:
 * p
 * b / strong
 * br
 * div
 * em / i
 * u
 * ul
 * ol
 * li
 * table
 * tbody
 * tr
 * td
 * img
 * a
 * h1/2/3/4

### Extending
Extending HtmlToWord is pretty easy. Each tag is a class that inherits from BaseElement. It has two methods that are called: *StartRender* and *EndRender*. Take a look in elements/headings.py and elements/text.py for some simple examples.



## Rationale
#### Why Word? Why not ODF or OpenOffice?
Time. Words Object Model is [very well documented](http://msdn.microsoft.com/en-us/library/ff837519) with lots of samples available on the internet - any .NET code in VB or C# can be translated pretty easily. On top of this you can record Macro's within Word that generates Visual Basic code while you play with a document, meaning its very quick to find out how to do things. ODF looks cool, but again i'm not getting any younger and word's COM interface ticked all the boxes. In the future I might expand this module to generate ODF XML, but for now its a pipe dream.