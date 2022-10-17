CSharpishStringFormatter
This code implements C sharpish string interpolation.<br><br>
The code was originally writter by Mathieu Guindon at [A CSharpish String.Format formatting helper](https://codereview.stackexchange.com/questions/30817/a-csharpish-string-format-formatting-helper)

As rescued by Greedquest at [Greedquest/VBA-Gems](https://github.com/Greedquest/VBA-Gems)

The included code has been inspected by Rubberduck and the majority of the code inspections resolved favourably.  The inspection for implicit use of .Item was not implemented.  twinBasic also found a couple more paramarray declaration that did not have 'as Variant'