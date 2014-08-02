<%
'Verbal Expression based Regular Expressions

class VerbalExpression

    private sub Class_Initialize
        set m_regexp = new Regexp
    end sub

    private sub Class_Deinitialize
        set m_regexp = nothing
    end sub

    public sub withAnyCase
        m_regexp.IgnoreCase = true
    end sub

    public sub searchGlobal
        m_regexp.Global = true
    end sub

    public sub searchOneLine
        m_regexp.Global = false
    end sub

    public function add( fnVal )
        m_source = m_source & fnVal
        call compile()
    end function

    public sub startOfLine
        m_prefixes = "^"
        call compile()
    end sub

    public sub endOfLine
        m_suffixes = "$"
        call compile()
    end sub

    public function then_( fnVal )
        add("(?:" & fnVal & ")")
    end function

    public function find( fnVal )
        then_(fnVal)
    end function

    public function maybe( fnVal )
        add("(?:" & fnVal & ")?")
    end function

    public sub anything
        add("(?:.*)")
    end sub

    public function anythingBut( fnVal )
        add("(?:[^" & fnVal & "]*)")
    end function

    public sub something
        add("(?:.+)")
    end sub

    public function somethingBut( fnVal )
        add("(?:[^" & fnVal & "]+)")
    end function

    public sub lineBreak
        add("(?:(?:\n)|(?:\r\n))")
    end sub

    public sub br
        lineBreak
    end sub

    public sub tab
        add("\t")
    end sub

    public sub word
        add("\w+")
    end sub

    public function anyOf( fnVal )
        add("[" & fnVal & "]")
    end function

    public function any( fnVal )
        anyOf(fnVal)
    end function

    public function range( fnArrVals )
        dim fnValue, fnArrItem
        fnvalue = "["

        for fnArrItem = lbound(fnArrVals) to ubound(fnArrVals) step 2
            fnvalue = fnvalue & fnArrVals(fnArrItem) & "-" & fnArrVals(fnArrItem+1)
        next

        fnvalue = fnvalu & "]"
        add(fnvalue)
    end function

    public function multiple( fnVal )
        if (left(fnVal,1) <> "*" and left(fnVal,1) <> "+") then
            add("+")
        end if
        add(fnVal)
    end function

    public function alt( fnVal )
        if instr(m_prefixes,"(") < 1 then m_prefixes = m_prefixes & "("
        if instr(m_suffixes,")") < 1 then m_suffixes = ")" & m_suffixes
        add(")|(")
        then_(fnval)
    end function

    public function test( fnVal )

        m_regexp.Pattern = m_pattern
        test = m_regexp.Test( fnVal )

    end function

    public function match( fnVal )
        m_regexp.Pattern = m_pattern
        match = m_regexp.Execute(fnVal)
    end function

    public function replace( fnReplace, fnVal )
        m_regexp.Pattern = m_pattern
        replace = m_regexp.replace(fnReplace,fnVal)
    end function

    private sub compile
        m_pattern = m_prefixes & m_source & m_suffixes
    end sub

    private m_prefixes
    private m_source
    private m_suffixes
    public m_pattern
    private m_regexp

end class

%>
