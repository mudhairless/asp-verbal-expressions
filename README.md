VBscript Verbal Expressions for Classic ASP
===========================================

## Regular Expressions made easy

Verbal Expressions is a VBscript (Classic ASP) module that helps to construct difficult regular expressions.

This VBscript module is based off of the original Javascript [Verbal expressions library](https://github.com/jehna/VerbalExpressions) by [jehna](https://github.com/jehna/).

## Examples

### Testing if we have a valid URL

```VBscript

    <!-- #include virtual="/path/to/verbalx.asp" -->
    <%

        url = "httpw://tacos.com"

        set vr = new VerbalExpression
        with vr
            .searchOneLine()
            .startOfLine()
            .then_("http")
            .maybe("s")
            .then_("://")
            .maybe("www.")
            .anythingBut(" ")
            .endOfLine()
        end with

        if (vr.test(url)) then
            response.write url & " is a valid url."
        else
            response.write url & " is NOT a valid url."
        end if

        response.write "<br/>"
        response.write vr.m_pattern

        set vr = nothing
    %>

```

## Using in your project

Include this library into your project with:

```VBscript
    <!-- #include virtual="/path/to/verbalx.asp" -->
```

## API

### Terms
* .anything()
* .anythingBut(string value)
* .something()
* .somethingBut(string value)
* .endOfLine()
* .find(string value)
* .maybe(string value)
* .startOfLine()
* .then_(string value)

### Special characters and groups
* .any(string value)
* .anyOf(string value)
* .br()
* .lineBreak()
* .range(args() as string)
* .tab()
* .word()

### Modifiers
* .withAnyCase()
* .searchOneLine()
* .searchGlobal()

### Functions
* .replace(string source, string value)
* .test(string value)
* .match(string value)

### Other
* .multiple(string value)
* .alt()

## Other implementations
You can view all implementations on [VerbalExpressions.github.io](http://VerbalExpressions.github.io)
