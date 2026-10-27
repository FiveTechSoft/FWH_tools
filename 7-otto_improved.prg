
*+
*+      "Virtual Keyboard"
#include "FiveWin.ch"
#include "colors.ch"

//----------------------------------------------------------------//
   static m_TextString := ""
   static m_shift := .f.
   static o11,o1,o2,o3,o4,o5,o6,o7,o8,o9,o0,o22,o33,oq,ow,oe,or,ot,oy,ou,oi,oo,op,o44,o55,o66,oa,os,od,off,og,oh,oj,ok,ol,o77,o88,oz,ox,oc,ov,ob,on,om,o99,o111,o222 
   static m_caps :=.f.
   static omessage
   static oDlg

function Keyboard()

    local ofont
    local cTitle := "Virtual Keyboard"

  define font ofont;
    name "ms sans serif";
    size 0,-16


   DEFINE DIALOG oDlg FROM 0,0 TO 500,1100 TITLE cTitle pixel font ofont color CLR_BLACK,CLR_BLUE

   ***** UNComment Just One ********
   @1,25 SAY omessage var m_TextString OF oDlg SIZE 250, 16 update border
   //@1,25 GET omessage var m_TextString OF oDlg SIZE 250, 16 
   //@ 1, 25 EDIT omessage VAR m_TextString OF oDlg SIZE 250, 16
   
           @ 205,7 ;
      BTNBMP  ;
      PROMPT "OK" ;
          OF oDlg ;
        FILE ("keyboard_58x28.bmp"), ("keyboard_58x28.bmp") ;
      ACTION (MsgInfo(m_TextString), oDlg:End()) ;
        SIZE 58, 28 ;
      ADJUST ;
      CENTER NOBORDER


      @ 50,455 btnbmp prompt "7" OF oDlg SIZE 28, 28 ;
      FILE ("keyboard_28x28.bmp"), ("keyboard_28x28.bmp") adjust center NOBORDER ;
     ACTION (m_TextString+="7", omessage:REFRESH()) 

     @ 50, 485 btnbmp prompt "8" OF oDlg SIZE 28, 28 ;
      FILE ("keyboard_28x28.bmp"), ("keyboard_28x28.bmp") adjust center NOBORDER ;
      ACTION (m_TextString+="8",omessage:refresh()) 

     @ 50, 515 btnbmp prompt "9" OF oDlg SIZE 28, 28 ;
      FILE ("keyboard_28x28.bmp"), ("keyboard_28x28.bmp") adjust center NOBORDER ;
   ACTION (m_TextString+="9",omessage:refresh()) 
     
     @ 80, 455 btnbmp prompt "4" OF oDlg SIZE 28, 28 ;
      FILE ("keyboard_28x28.bmp"), ("keyboard_28x28.bmp") adjust center NOBORDER ;
   ACTION (m_TextString+="4",omessage:refresh()) 
     
     @ 80, 485 btnbmp prompt "5" OF oDlg SIZE 28, 28 ;
      FILE ("keyboard_28x28.bmp"), ("keyboard_28x28.bmp") adjust center NOBORDER ;
  ACTION (m_TextString+="5",omessage:refresh()) 
     
     @ 80, 515 btnbmp prompt "6" OF oDlg SIZE 28, 28 ;
      FILE ("keyboard_28x28.bmp"), ("keyboard_28x28.bmp") adjust center NOBORDER ;
   ACTION (m_TextString+="6",omessage:refresh()) 
     
     @ 110, 455 btnbmp prompt "1" OF oDlg SIZE 28, 28 ;
      FILE ("keyboard_28x28.bmp"), ("keyboard_28x28.bmp") adjust center NOBORDER ;
   ACTION (m_TextString+="1",omessage:refresh()) 
     
     @ 110, 485 btnbmp prompt "2" OF oDlg SIZE 28, 28 ;
      FILE ("keyboard_28x28.bmp"), ("keyboard_28x28.bmp") adjust center NOBORDER ;
  ACTION (m_TextString+="2",omessage:refresh()) 
     
     @ 110, 515 btnbmp prompt "3" OF oDlg SIZE 28, 28 ;
      FILE ("keyboard_28x28.bmp"), ("keyboard_28x28.bmp") adjust center NOBORDER ;
  ACTION (m_TextString+="3",omessage:refresh()) 
     
     @ 50, 5 btnbmp o11 prompt "`" OF oDlg SIZE 28, 28 ;
          FILE ("keyboard_28x28.bmp"), ("keyboard_28x28.bmp") adjust center NOBORDER ;
    ACTION (iif(m_shift, m_TextString+="~",m_TextString+="`"), unshift(), omessage:refresh()) 
     
     @ 50, 35 btnbmp o1 prompt "1" OF oDlg SIZE 28, 28 ;
          FILE ("keyboard_28x28.bmp"), ("keyboard_28x28.bmp") adjust center NOBORDER ;
   ACTION (iif(m_shift, m_TextString+="!",m_TextString+="1"), unshift(), omessage:refresh()) 
     
     @ 50, 65 btnbmp o2 prompt "2" OF oDlg SIZE 28, 28 ;
          FILE ("keyboard_28x28.bmp"), ("keyboard_28x28.bmp") adjust center NOBORDER ;
   ACTION (iif(m_shift, m_TextString+="@",m_TextString+="2"), unshift(), omessage:refresh()) 
     
     @ 50, 95 btnbmp o3 prompt "3" OF oDlg SIZE 28, 28 ;
          FILE ("keyboard_28x28.bmp"), ("keyboard_28x28.bmp") adjust center NOBORDER ;
   ACTION (iif(m_shift, m_TextString+="#",m_TextString+="3"), unshift(), omessage:refresh()) 
     
     @ 50, 125 btnbmp o4 prompt "4" OF oDlg SIZE 28, 28 ;
          FILE ("keyboard_28x28.bmp"), ("keyboard_28x28.bmp") adjust center NOBORDER ;
  ACTION (iif(m_shift, m_TextString+="$",m_TextString+="4"), unshift(), omessage:refresh()) 
     
     @ 50, 155 btnbmp o5 prompt "5" OF oDlg SIZE 28, 28 ;
          FILE ("keyboard_28x28.bmp"), ("keyboard_28x28.bmp") adjust center NOBORDER ;
  ACTION (iif(m_shift, m_TextString+="%",m_TextString+="5"), unshift(), omessage:refresh()) 
     
     @ 50, 185 btnbmp o6 prompt "6" OF oDlg SIZE 28, 28 ;
          FILE ("keyboard_28x28.bmp"), ("keyboard_28x28.bmp") adjust center NOBORDER ;
   ACTION (iif(m_shift, m_TextString+="^",m_TextString+="6"), unshift(), omessage:refresh()) 
     
     @ 50, 215 btnbmp o7 prompt "7" OF oDlg SIZE 28, 28 ;
          FILE ("keyboard_28x28.bmp"), ("keyboard_28x28.bmp") adjust center NOBORDER ;
   ACTION (iif(m_shift, m_TextString+="&",m_TextString+="7"), unshift(), omessage:refresh()) 
     
     @ 50, 245 btnbmp o8 prompt "8" OF oDlg SIZE 28, 28 ;
          FILE ("keyboard_28x28.bmp"), ("keyboard_28x28.bmp") adjust center NOBORDER ;
  ACTION (iif(m_shift, m_TextString+="*",m_TextString+="8"), unshift(), omessage:refresh()) 
     
     @ 50, 275 btnbmp o9 prompt "9" OF oDlg SIZE 28, 28 ;
          FILE ("keyboard_28x28.bmp"), ("keyboard_28x28.bmp") adjust center NOBORDER ;
  ACTION (iif(m_shift, m_TextString+="(",m_TextString+="9"), unshift(), omessage:refresh()) 
     
     @ 50, 305 btnbmp o0 prompt "0" OF oDlg SIZE 28, 28 ;
          FILE ("keyboard_28x28.bmp"), ("keyboard_28x28.bmp") adjust center NOBORDER ;
   ACTION (iif(m_shift, m_TextString+=")",m_TextString+="0"), unshift(), omessage:refresh()) 
     
     @ 50, 335 btnbmp o22 prompt "-" OF oDlg SIZE 28, 28 ;
          FILE ("keyboard_28x28.bmp"), ("keyboard_28x28.bmp") adjust center NOBORDER ;
  ACTION (iif(m_shift, m_TextString+="_",m_TextString+="-"), unshift(), omessage:refresh()) 
     
     @ 50, 365 btnbmp o33 prompt "=" OF oDlg SIZE 28, 28 ;
          FILE ("keyboard_28x28.bmp"), ("keyboard_28x28.bmp") adjust center NOBORDER ;
   ACTION (iif(m_shift, m_TextString+="+",m_TextString+="="), unshift(), omessage:refresh()) 
     
     @ 50, 395 btnbmp prompt "BKS" OF oDlg SIZE 28, 28 ;
          FILE ("keyboard_28x28.bmp"), ("keyboard_28x28.bmp") adjust center NOBORDER ;
  ACTION (m_TextString:=substr(m_TextString,1,len(m_TextString)-1),omessage:refresh()) 
     
     @ 80, 5 btnbmp prompt "TAB" OF oDlg SIZE 28, 28 ;
          FILE ("keyboard_28x28.bmp"), ("keyboard_28x28.bmp") adjust center NOBORDER ;
  ACTION (m_TextString+="    ",omessage:refresh()) 
     
     @ 80, 35 btnbmp oQ prompt "q" OF oDlg SIZE 28, 28 ;
          FILE ("keyboard_28x28.bmp"), ("keyboard_28x28.bmp") adjust center NOBORDER ;
  ACTION (iif(m_shift, m_TextString+="Q",m_TextString+="q"), unshift(), omessage:refresh()) 
     
     @ 80, 65 btnbmp ow prompt "w" OF oDlg SIZE 28, 28 ;
          FILE ("keyboard_28x28.bmp"), ("keyboard_28x28.bmp") adjust center NOBORDER ;
  ACTION (iif(m_shift, m_TextString+="W",m_TextString+="w"), unshift(), omessage:refresh()) 
     
     @ 80, 95 btnbmp oe prompt "e" OF oDlg SIZE 28, 28 ;
          FILE ("keyboard_28x28.bmp"), ("keyboard_28x28.bmp") adjust center NOBORDER ;
  ACTION (iif(m_shift, m_TextString+="E",m_TextString+="e"), unshift(), omessage:refresh()) 
     
     @ 80, 125 btnbmp or prompt "r" OF oDlg SIZE 28, 28 ;
          FILE ("keyboard_28x28.bmp"), ("keyboard_28x28.bmp") adjust center NOBORDER ;
  ACTION (iif(m_shift, m_TextString+="R",m_TextString+="r"), unshift(), omessage:refresh()) 
     
     @ 80, 155 btnbmp ot prompt "t" OF oDlg SIZE 28, 28 ;
          FILE ("keyboard_28x28.bmp"), ("keyboard_28x28.bmp") adjust center NOBORDER ;
  ACTION (iif(m_shift, m_TextString+="T",m_TextString+="t"), unshift(), omessage:refresh()) 
     
     @ 80, 185 btnbmp oy prompt "y" OF oDlg SIZE 28, 28 ;
          FILE ("keyboard_28x28.bmp"), ("keyboard_28x28.bmp") adjust center NOBORDER ;
  ACTION (iif(m_shift, m_TextString+="Y",m_TextString+="y"), unshift(), omessage:refresh()) 
     
     @ 80, 215 btnbmp ou prompt "u" OF oDlg SIZE 28, 28 ;
          FILE ("keyboard_28x28.bmp"), ("keyboard_28x28.bmp") adjust center NOBORDER ;
   ACTION (iif(m_shift, m_TextString+="U",m_TextString+="u"), unshift(), omessage:refresh()) 
     
     @ 80, 245 btnbmp oi prompt "i" OF oDlg SIZE 28, 28 ;
          FILE ("keyboard_28x28.bmp"), ("keyboard_28x28.bmp") adjust center NOBORDER ;
   ACTION (iif(m_shift, m_TextString+="I",m_TextString+="i"), unshift(), omessage:refresh()) 
     
     @ 80, 275 btnbmp oo prompt "o" OF oDlg SIZE 28, 28 ;
          FILE ("keyboard_28x28.bmp"), ("keyboard_28x28.bmp") adjust center NOBORDER ;
   ACTION (iif(m_shift, m_TextString+="O",m_TextString+="o"), unshift(), omessage:refresh()) 
     
     @ 80, 305 btnbmp op prompt "p" OF oDlg SIZE 28, 28 ;
          FILE ("keyboard_28x28.bmp"), ("keyboard_28x28.bmp") adjust center NOBORDER ;
  ACTION (iif(m_shift, m_TextString+="P",m_TextString+="p"), unshift(), omessage:refresh()) 
     
     @ 80, 335 btnbmp o44 prompt "[" OF oDlg SIZE 28, 28 ;
          FILE ("keyboard_28x28.bmp"), ("keyboard_28x28.bmp") adjust center NOBORDER ;
  ACTION (iif(m_shift, m_TextString+="{",m_TextString+="["), unshift(), omessage:refresh()) 
     
     @ 80, 365 btnbmp o55 prompt "]" OF oDlg SIZE 28, 28 ;
          FILE ("keyboard_28x28.bmp"), ("keyboard_28x28.bmp") adjust center NOBORDER ;
  ACTION (iif(m_shift, m_TextString+="}",m_TextString+="]"), unshift(), omessage:refresh()) 
     
     @ 80, 395 btnbmp o66 prompt "\" OF oDlg SIZE 28, 28 ;
          FILE ("keyboard_28x28.bmp"), ("keyboard_28x28.bmp") adjust center NOBORDER ;
   ACTION (iif(m_shift, m_TextString+="|",m_TextString+="\"), unshift(), omessage:refresh()) 
     
     @ 110, 5 btnbmp prompt "Cap" OF oDlg SIZE 28, 28 ;
          FILE ("keyboard_28x28.bmp"), ("keyboard_28x28.bmp") adjust center NOBORDER ;
   ACTION (docaps()) 
     
     @ 110, 35 btnbmp oa prompt "a" OF oDlg SIZE 28, 28 ;
          FILE ("keyboard_28x28.bmp"), ("keyboard_28x28.bmp") adjust center NOBORDER ;
  ACTION (iif(m_shift, m_TextString+="A",m_TextString+="a"), unshift(), omessage:refresh()) 
     
     @ 110, 65 btnbmp os prompt "s" OF oDlg SIZE 28, 28 ;
          FILE ("keyboard_28x28.bmp"), ("keyboard_28x28.bmp") adjust center NOBORDER ;
  ACTION (iif(m_shift, m_TextString+="S",m_TextString+="s"), unshift(), omessage:refresh()) 
     
     @ 110, 95 btnbmp od prompt "d" OF oDlg SIZE 28, 28 ;
          FILE ("keyboard_28x28.bmp"), ("keyboard_28x28.bmp") adjust center NOBORDER ;
  ACTION (iif(m_shift, m_TextString+="D",m_TextString+="d"), unshift(), omessage:refresh()) 
     
     @ 110, 125 btnbmp off prompt "f" OF oDlg SIZE 28, 28 ;
          FILE ("keyboard_28x28.bmp"), ("keyboard_28x28.bmp") adjust center NOBORDER ;
  ACTION (iif(m_shift, m_TextString+="F",m_TextString+="f"), unshift(), omessage:refresh()) 
     
     @ 110, 155 btnbmp og prompt "g" OF oDlg SIZE 28, 28 ;
          FILE ("keyboard_28x28.bmp"), ("keyboard_28x28.bmp") adjust center NOBORDER ;
  ACTION (iif(m_shift, m_TextString+="G",m_TextString+="g"), unshift(), omessage:refresh()) 
     
     @ 110, 185 btnbmp oh prompt "h" OF oDlg SIZE 28, 28 ;
          FILE ("keyboard_28x28.bmp"), ("keyboard_28x28.bmp") adjust center NOBORDER ;
   ACTION (iif(m_shift, m_TextString+="H",m_TextString+="h"), unshift(), omessage:refresh()) 
     
     @ 110, 215 btnbmp oj prompt "j" OF oDlg SIZE 28, 28 ;
          FILE ("keyboard_28x28.bmp"), ("keyboard_28x28.bmp") adjust center NOBORDER ;
   ACTION (iif(m_shift, m_TextString+="J",m_TextString+="j"), unshift(), omessage:refresh()) 
     
     @ 110, 245 btnbmp ok prompt "k" OF oDlg SIZE 28, 28 ;
          FILE ("keyboard_28x28.bmp"), ("keyboard_28x28.bmp") adjust center NOBORDER ;
  ACTION (iif(m_shift, m_TextString+="K",m_TextString+="k"), unshift(), omessage:refresh()) 
     
     @ 110, 275 btnbmp ol prompt "l" OF oDlg SIZE 28, 28 ;
          FILE ("keyboard_28x28.bmp"), ("keyboard_28x28.bmp") adjust center NOBORDER ;
  ACTION (iif(m_shift, m_TextString+="L",m_TextString+="l"), unshift(), omessage:refresh()) 
     
     @ 110, 305 btnbmp o77 prompt ";" OF oDlg SIZE 28, 28 ;
          FILE ("keyboard_28x28.bmp"), ("keyboard_28x28.bmp") adjust center NOBORDER ;
 ACTION (iif(m_shift, m_TextString+=":",m_TextString+=";"), unshift(), omessage:refresh()) 
     
     @ 110, 335 btnbmp o88 prompt "'" OF oDlg SIZE 28, 28 ;
          FILE ("keyboard_28x28.bmp"), ("keyboard_28x28.bmp") adjust center NOBORDER ;
 ACTION (iif(m_shift, m_TextString+='"',m_TextString+="'"), unshift(), omessage:refresh()) 
     
     @ 110, 365 btnbmp prompt "Enter" OF oDlg SIZE 58, 28 ;
          FILE ("keyboard_58x28.bmp"), ("keyboard_58x28.bmp") adjust center NOBORDER ;
  ACTION oDlg:End() 
     
     @ 140, 5 btnbmp prompt "Shift" OF oDlg SIZE 58, 28 ;
          FILE ("keyboard_58x28.bmp"), ("keyboard_58x28.bmp") adjust center NOBORDER ;
  ACTION (doshift()) 
     
     @ 140, 65 btnbmp oz prompt "z" OF oDlg SIZE 28, 28 ;
          FILE ("keyboard_28x28.bmp"), ("keyboard_28x28.bmp") adjust center NOBORDER ;
   ACTION (iif(m_shift, m_TextString+="Z",m_TextString+="z"), unshift(), omessage:refresh()) 
     
     @ 140, 95 btnbmp ox prompt "x" OF oDlg SIZE 28, 28 ;
          FILE ("keyboard_28x28.bmp"), ("keyboard_28x28.bmp") adjust center NOBORDER ;
    ACTION (iif(m_shift, m_TextString+="X",m_TextString+="x"), unshift(), omessage:refresh()) 
     
     @ 140, 125 btnbmp oc prompt "c" OF oDlg SIZE 28, 28 ;
          FILE ("keyboard_28x28.bmp"), ("keyboard_28x28.bmp") adjust center NOBORDER ;
  ACTION (iif(m_shift, m_TextString+="C",m_TextString+="c"), unshift(), omessage:refresh()) 
     
     @ 140, 155 btnbmp ov prompt "v" OF oDlg SIZE 28, 28 ;
          FILE ("keyboard_28x28.bmp"), ("keyboard_28x28.bmp") adjust center NOBORDER ;
  ACTION (iif(m_shift, m_TextString+="V",m_TextString+="v"), unshift(), omessage:refresh()) 
     
     @ 140, 185 btnbmp ob prompt "b" OF oDlg SIZE 28, 28 ;
          FILE ("keyboard_28x28.bmp"), ("keyboard_28x28.bmp") adjust center NOBORDER ;
   ACTION (iif(m_shift, m_TextString+="B",m_TextString+="b"), unshift(), omessage:refresh()) 
     
     @ 140, 215 btnbmp on prompt "n" OF oDlg SIZE 28, 28 ;
          FILE ("keyboard_28x28.bmp"), ("keyboard_28x28.bmp") adjust center NOBORDER ;
  ACTION (iif(m_shift, m_TextString+="N",m_TextString+="n"), unshift(), omessage:refresh()) 
     
     @ 140, 245 btnbmp om prompt "m" OF oDlg SIZE 28, 28 ;
          FILE ("keyboard_28x28.bmp"), ("keyboard_28x28.bmp") adjust center NOBORDER ;
  ACTION (iif(m_shift, m_TextString+="M",m_TextString+="m"), unshift(), omessage:refresh()) 
     
     @ 140, 275 btnbmp o99 prompt "," OF oDlg SIZE 28, 28 ;
          FILE ("keyboard_28x28.bmp"), ("keyboard_28x28.bmp") adjust center NOBORDER ;
  ACTION (iif(m_shift, m_TextString+="<",m_TextString+=","), unshift(), omessage:refresh()) 
     
     @ 140, 305 btnbmp o111 prompt "." OF oDlg SIZE 28, 28 ;
          FILE ("keyboard_28x28.bmp"), ("keyboard_28x28.bmp") adjust center NOBORDER ;
 ACTION (iif(m_shift, m_TextString+=">",m_TextString+="."), unshift(), omessage:refresh()) 
     
     @ 140, 335 btnbmp o222 prompt "/" OF oDlg SIZE 28, 28 ;
          FILE ("keyboard_28x28.bmp"), ("keyboard_28x28.bmp") adjust center NOBORDER ;
  ACTION (iif(m_shift, m_TextString+="?",m_TextString+="/"), unshift(), omessage:refresh()) 
     
     @ 140, 365 btnbmp prompt "Shift" OF oDlg SIZE 58, 28 ;
          FILE ("keyboard_58x28.bmp"), ("keyboard_58x28.bmp") adjust center NOBORDER ;
   ACTION (doshift()) 
     
     @ 140, 455 btnbmp prompt "0" OF oDlg SIZE 58, 28 ;
          FILE ("keyboard_58x28.bmp"), ("keyboard_58x28.bmp") adjust center NOBORDER ;
  ACTION (m_TextString+="0",omessage:refresh()) 
     
     @ 140, 515 btnbmp prompt "." OF oDlg SIZE 28, 28 ;
          FILE ("keyboard_28x28.bmp"), ("keyboard_28x28.bmp") adjust center NOBORDER ;
  ACTION (m_TextString+=".",omessage:refresh()) 
     
     @ 170, 120 btnbmp prompt "Space" OF oDlg SIZE 165, 28 ;
          FILE ("keyboard_165x28.bmp"), ("keyboard_165x28.bmp") adjust center NOBORDER ;
  ACTION (m_TextString+=" ",omessage:refresh()) 
     

  ***** Comment out for GET and EDIT **********  
  oDlg:bKeyDown := { | nKey | KEYSEEK(nKey)}
     
   
   
   ACTIVATE DIALOG oDlg CENTERED

return(m_TextString)


*+??????????????????????????????????
*+
*+    Function KeySeek()
*+
*+??????????????????????????????????
*+
Function KeySeek( nKey )


MsgInfo(STR(nKey,3)+"   |"+Chr(nKey)+"|")

omessage:SetFocus()
if nKey = 8         // Backup on BackSpace Key
    m_TextString:=substr(m_TextString,1,len(m_TextString)-1)
    omessage:Refresh()
//elseif ( nKey > 64 .and. nKey < 91 ) .or. ( nKey > 96 .and. nKey < 123 ) .or. ;
//             ( nKey = 32 )              // Check for valid keys ONLY
//    m_TextString := m_TextString + chr( nKey )
//    omessage:Refresh()
elseif ( nKey == 27 )
    m_TextString := ""
    omessage:Refresh()
elseif ( nKey == 13 )
    oDlg:end()
else
   m_TextString += Chr(nKey)
   omessage:Refresh()
endif

return .T.




function docaps()
if m_caps=.f.
    m_caps:=.t.
    doshift()
else
    m_caps:=.f.
    unshift()
end if
return .t.

function doshift()
m_shift:=.t.
o11:settext("~")
o1:settext("!")
o2:settext("@")
o3:settext("#")
o4:settext("$")
o5:settext("%")
o6:settext("^")
o7:settext("&&")
o8:settext("*")
o9:settext("(")
o0:settext(")")
o22:settext("_")
o33:settext("+")
oq:settext("Q")
ow:settext("W")
oe:settext("E")
or:settext("R")
ot:settext("T")
oy:settext("Y")
ou:settext("U")
oi:settext("I")
oo:settext("O")
op:settext("P")
o44:settext("{")
o55:settext("}")
o66:settext("|")
oa:settext("A")
os:settext("S")
od:settext("D")
off:settext("F")
og:settext("G")
oh:settext("H")
oj:settext("J")
ok:settext("K")
ol:settext("L")
o77:settext(":")
o88:settext('"')
oz:settext("Z")
ox:settext("X")
oc:settext("C")
ov:settext("V")
ob:settext("B")
on:settext("N")
om:settext("M")
o99:settext("<")
o111:settext(">")
o222:settext("?")
o11:refresh()
o1:refresh()
o2:refresh()
o3:refresh()
o4:refresh()
o5:refresh()
o6:refresh()
o7:refresh()
o8:refresh()
o9:refresh()
o0:refresh()
o22:refresh()
o33:refresh()
oq:refresh()
ow:refresh()
oe:refresh()
or:refresh()
ot:refresh()
oy:refresh()
ou:refresh()
oi:refresh()
oo:refresh()
op:refresh()
o44:refresh()
o55:refresh()
o66:refresh()
oa:refresh()
os:refresh()
od:refresh()
off:refresh()
og:refresh()
oh:refresh()
oj:refresh()
ok:refresh()
ol:refresh()
o77:refresh()
o88:refresh()
oz:refresh()
ox:refresh()
oc:refresh()
ov:refresh()
ob:refresh()
on:refresh()
om:refresh()
o99:refresh()
o111:refresh()
o222:refresh()
return .t.

function unshift()
if m_caps=.t.
    m_shift:=.t.
    return .t.
else 
    m_shift:=.f.
end if
o11:settext("`")
o1:settext("1")
o2:settext("2")
o3:settext("3")
o4:settext("4")
o5:settext("5")
o6:settext("6")
o7:settext("7")
o8:settext("8")
o9:settext("9")
o0:settext("0")
o22:settext("-")
o33:settext("=")
oq:settext("q")
ow:settext("w")
oe:settext("e")
or:settext("r")
ot:settext("t")
oy:settext("y")
ou:settext("u")
oi:settext("i")
oo:settext("o")
op:settext("p")
o44:settext("[")
o55:settext("]")
o66:settext("\")
oa:settext("a")
os:settext("s")
od:settext("d")
off:settext("f")
og:settext("g")
oh:settext("h")
oj:settext("j")
ok:settext("k")
ol:settext("l")
o77:settext(";")
o88:settext("'")
oz:settext("z")
ox:settext("x")
oc:settext("c")
ov:settext("v")
ob:settext("b")
on:settext("n")
om:settext("m")
o99:settext(",")
o111:settext(".")
o222:settext("/")
o11:refresh()
o1:refresh()
o2:refresh()
o3:refresh()
o4:refresh()
o5:refresh()
o6:refresh()
o7:refresh()
o8:refresh()
o9:refresh()
o0:refresh()
o22:refresh()
o33:refresh()
oq:refresh()
ow:refresh()
oe:refresh()
or:refresh()
ot:refresh()
oy:refresh()
ou:refresh()
oi:refresh()
oo:refresh()
op:refresh()
o44:refresh()
o55:refresh()
o66:refresh()
oa:refresh()
os:refresh()
od:refresh()
off:refresh()
og:refresh()
oh:refresh()
oj:refresh()
ok:refresh()
ol:refresh()
o77:refresh()
o88:refresh()
oz:refresh()
ox:refresh()
oc:refresh()
ov:refresh()
ob:refresh()
on:refresh()
om:refresh()
o99:refresh()
o111:refresh()
o222:refresh()
return .t.

