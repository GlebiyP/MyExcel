grammar LabCalc;

/*
* Parser Rules
*/

compileUnit : expression EOF;
expression:
  LPAREN expression RPAREN #ParenthesizedExpr
  |expression EXPONENT expression #ExponentialExpr
  |expression operatorToken=(MULTIPLY | DIVIDE) expression #MultiplicativeExpr
  |expression operatorToken=(ADD | SUBTRACT) expression #AdditiveExpr
  |expression operatorToken=(MLESS | MMORE | LESSOREQUAL | MOREOREQUAL | ISEQUAL | NOTEQUAL) expression #CompareExpr
  |MMAX LPAREN paramlist=arglist RPAREN #MmaxExpr
  |MMIN LPAREN paramlist=arglist RPAREN #MminExpr
  |operatorToken=SUBTRACT expression #MinusExpr
  |NUMBER #NumberExpr
  |NOT LPAREN expression RPAREN #NotExpr
  |IDENTIFIER #IdentifierExpr;


/*
* Lexer Rules
*/

arglist: expression (',' expression)+;
paramlist: expression(',' expression)+;
NUMBER : INT ('.' INT)?;
IDENTIFIER : [A-Z]+[1-9][0-9]*;

INT:('0'..'9')+;

EXPONENT : '^';
MULTIPLY : '*';
DIVIDE : '/';
SUBTRACT : '-';
ADD : '+';
LPAREN : '(';
RPAREN : ')';
MLESS : '<';
MMORE : '>';
LESSOREQUAL : '<=';
MOREOREQUAL : '>=';
ISEQUAL : '==';
NOTEQUAL : '<>';
NOT : 'not';
MMAX : 'mmax';
MMIN : 'mmin';

WS : [ \t\r\n] -> channel(HIDDEN);