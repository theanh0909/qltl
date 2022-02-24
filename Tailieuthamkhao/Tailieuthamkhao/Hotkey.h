/*
	http://www.physics.udel.edu/~watson/scen103/ascii.html

	Alt: WM_SYSKEYDOWN
	pMsg->message == WM_SYSKEYDOWN  && pMsg->wParam == VK_F4

	Ctrl + @	0x00
	Ctrl + A	0x01
	Ctrl + B	0x02
	Ctrl + C	0x03
	Ctrl + D	0x04
	Ctrl + E	0x05
	Ctrl + F	0x06
	Ctrl + G	0x07
	Ctrl + H	0x08
	Ctrl + I	0x09
	Ctrl + J	0x0A

	Ctrl + K	0x0B
	Ctrl + L	0x0C
	Ctrl + M	0x0D
	Ctrl + N	0x0E
	Ctrl + O	0x0F
	Ctrl + P	0x10
	Ctrl + Q	0x11
	Ctrl + R	0x12
	Ctrl + S	0x13
	Ctrl + T	0x14

	Ctrl + U	0x15
	Ctrl + V	0x16
	Ctrl + W	0x17
	Ctrl + X	0x18
	Ctrl + Y	0x19
	Ctrl + Z	0x1A

	Ctrl + [	0x1B
	Ctrl + \	0x1C
	Ctrl + ]	0x1D
	Ctrl + ^	0x1E
	Ctrl + _	0x1F

	TCHAR chr = static_cast<TCHAR>(pMsg->wParam);
	Decimal Octal Hex   Binairy   Character
	   035  043  0x23  00100011    #  (number sign)
	   036  044  0x24  00100100    $  (dollar sign)
	   037  045  0x25  00100101    %  (percent)
	   038  046  0x26  00100110    &  (ampersand)
	   039  047  0x27  00100111    '  (single quote)
	   040  050  0x28  00101000    (  (left opening parenthesis)
	   041  051  0x29  00101001    )  (right closing parenthesis)
	   042  052  0x2A  00101010    *  (asterisk)
	   043  053  0x2B  00101011    +  (plus)
	   044  054  0x2C  00101100    ,  (comma)
	   045  055  0x2D  00101101    -  (minus or dash)
	   046  056  0x2E  00101110    .  (dot)
	   047  057  0x2F  00101111    /  (forward slash)
	   048  060  0x30  00110000    0
	   049  061  0x31  00110001    1
	   050  062  0x32  00110010    2
	   051  063  0x33  00110011    3
	   052  064  0x34  00110100    4
	   053  065  0x35  00110101    5
	   054  066  0x36  00110110    6
	   055  067  0x37  00110111    7
	   056  070  0x38  00111000    8
	   057  071  0x39  00111001    9
	   058  072  0x3A  00111010    :  (colon)
	   059  073  0x3B  00111011    ;  (semi-colon)
	   060  074  0x3C  00111100    <  (less than sign)
	   061  075  0x3D  00111101    =  (equal sign)
	   062  076  0x3E  00111110    >  (greater than sign)
	   063  077  0x3F  00111111    ?  (question mark)
	   064  100  0x40  01000000    @  (AT symbol)
	   065  101  0x41  01000001    A
	   066  102  0x42  01000010    B
	   067  103  0x43  01000011    C
	   068  104  0x44  01000100    D
	   069  105  0x45  01000101    E
	   070  106  0x46  01000110    F
	   071  107  0x47  01000111    G
	   072  110  0x48  01001000    H
	   073  111  0x49  01001001    I
	   074  112  0x4A  01001010    J
	   075  113  0x4B  01001011    K
	   076  114  0x4C  01001100    L
	   077  115  0x4D  01001101    M
	   078  116  0x4E  01001110    N
	   079  117  0x4F  01001111    O
	   080  120  0x50  01010000    P
	   081  121  0x51  01010001    Q
	   082  122  0x52  01010010    R
	   083  123  0x53  01010011    S
	   084  124  0x54  01010100    T
	   085  125  0x55  01010101    U
	   086  126  0x56  01010110    V
	   087  127  0x57  01010111    W
	   088  130  0x58  01011000    X
	   089  131  0x59  01011001    Y
	   090  132  0x5A  01011010    Z
	   091  133  0x5B  01011011    [  (left opening bracket)
	   092  134  0x5C  01011100    \  (back slash)
	   093  135  0x5D  01011101    ]  (right closing bracket)
	   094  136  0x5E  01011110    ^  (caret cirumflex)
	   095  137  0x5F  01011111    _  (underscore)
	   096  140  0x60  01100000    `
	   097  141  0x61  01100001    a
	   098  142  0x62  01100010    b
	   099  143  0x63  01100011    c
	   100  144  0x64  01100100    d
	   101  145  0x65  01100101    e
	   102  146  0x66  01100110    f
	   103  147  0x67  01100111    g
	   104  150  0x68  01101000    h
	   105  151  0x69  01101001    i
	   106  152  0x6A  01101010    j
	   107  153  0x6B  01101011    k
	   108  154  0x6C  01101100    l
	   109  155  0x6D  01101101    m
	   110  156  0x6E  01101110    n
	   111  157  0x6F  01101111    o
	   112  160  0x70  01110000    p
	   113  161  0x71  01110001    q
	   114  162  0x72  01110010    r
	   115  163  0x73  01110011    s
	   116  164  0x74  01110100    t
	   117  165  0x75  01110101    u
	   118  166  0x76  01110110    v
	   119  167  0x77  01110111    w
	   120  170  0x78  01111000    x
	   121  171  0x79  01111001    y
	   122  172  0x7A  01111010    z
	   123  173  0x7B  01111011    {  (left opening brace)
	   124  174  0x7C  01111100    |  (vertical bar)
	   125  175  0x7D  01111101    }  (right closing brace)
	   126  176  0x7E  01111110    ~  (tilde)
	   127  177  0x7F  01111111  DEL  (delete)

*/
