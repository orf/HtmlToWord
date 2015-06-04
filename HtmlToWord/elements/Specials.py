# -*- coding: utf-8 -*-
"""
module for special and custom elements
"""
import warnings

from HtmlToWord.elements.Base import ChildlessElement


class Footnote(ChildlessElement):
    """
    Footnotes support:
    in order to correctly insert a footnote, it expect a custom html tag 'footnote' with an attribute
    'data-content' that store the actual footnote text, e.g.:
    <footnote data-content="footnote text content"></footnote>
    """

    def StartRender(self):
        self.initial_position = None
        footnote_content = self.GetAttrs()["data-content"]
        if footnote_content:
            self.initial_position = self.getStartPosition()
            footnote = self.document.Footnotes.Add(self.selection.Range)
            footnote.Range.Text = footnote_content

    def EndRender(self):
        # restore initial cursor position + 1
        if self.initial_position:
            rng = self.document.Range(self.initial_position+1, self.initial_position+1)
            rng.Select()


