//-----------------------------------------------------------------------------
//   mathexpr.c
//
//   Evaluates symbolic mathematical expression consisting
//   of numbers, variable names, math functions & arithmetic
//   operators.
//
//   Date:     6/16/03
//             11/22/03
//             2/20/04
//             9/03/04
//   Author:   L. Rossman
//-----------------------------------------------------------------------------

#include <malloc.h>
#include <string.h>
#include <math.h>
#include "mathexpr.h"

//--------------------
// Math Function Names
//-------------------- 
char *MathFunc[] = {"COS", "SIN", "TAN", "COT", "ABS", "SGN",
                    "SQRT", "LOG", "EXP", "ASIN", "ACOS", "ATAN",
                    "ACOT", "SINH", "COSH", "TANH", "COTH", "LOG10", NULL};

//-----------------------------------------------------------------------------
// Shared variables
//-----------------------------------------------------------------------------
static int    Err;
static int    Bc;
static int    PrevLex, CurLex;
static int    Len, Pos;
static char   *S;
static char   Token[255];
static int    Ivar;
static double Fvalue;

//-----------------------------------------------------------------------------
// Local functions
//-----------------------------------------------------------------------------
int     isDigit(char);
int     isLetter(char);
void    getToken(void);
int     getMathFunc(void);
int     getVariable(void);
int     getOperand(void);
int     getLex(void);
double  getNumber(void);
double  c(ExprTree *);
ExprTree*  newNode(void);
ExprTree*  getSingleOp(int *);
ExprTree*  getOp(int *);
ExprTree*  getTree(void);
                                            // Callback functions
static int    (*getVariableIndex) (char *); // return index of named variable
static double (*getVariableValue) (int);    // return value of indexed variable

//-----------------------------------------------------------------------------
//  External functions (must be provided in calling program)
//-----------------------------------------------------------------------------
int strcomp(char *, char *);                // case-insensitive string compare


int isDigit(char c)
{
    if (c >= '1' && c <= '9') return 1;
    if (c == '0') return 1;
    return 0;
}

int isLetter(char c)
{
    if (c >= 'a' && c <= 'z') return 1;
    if (c >= 'A' && c <= 'Z') return 1;
    if (c == '_') return 1;
    return 0;
}

void getToken()
{
    char c[] = " ";
    strcpy(Token, "");
    while ( Pos <= Len &&
        ( isLetter(S[Pos]) || isDigit(S[Pos]) ) )
    {
        c[0] = S[Pos];
        strcat(Token, c);
        Pos++;
    }
    Pos--;
}

int getMathFunc()
{
    int i = 0;
    while (MathFunc[i] != NULL)
    {
        if (strcomp(MathFunc[i], Token)) return i+10;
        i++;
    }
    return(0);
}

int getVariable()
{
    if ( !getVariableIndex ) return 0;
    Ivar = getVariableIndex(Token);
    if (Ivar >= 0) return 8;
    return 0;
}

double getNumber()
{
    char c[] = " ";
    char sNumber[255];
    int  errflag = 0;

    /* --- get whole number portion of number */
    strcpy(sNumber, "");
    while (Pos < Len && isDigit(S[Pos]))
    {
        c[0] = S[Pos];
        strcat(sNumber, c);
        Pos++;
    }

    /* --- get fractional portion of number */
    if (Pos < Len)
    {
        if (S[Pos] == '.')
        {
            strcat(sNumber, ".");
            Pos++;
            while (Pos < Len && isDigit(S[Pos]))
            {
                c[0] = S[Pos];
                strcat(sNumber, c);  
                Pos++;
            }
        }

        /* --- get exponent */
        if (Pos < Len && (S[Pos] == 'e' || S[Pos] == 'E'))
        {
            strcat(sNumber, "E");  
            Pos++;
            if (Pos >= Len) errflag = 1;
            else
            {
                if (S[Pos] == '-' || S[Pos] == '+')
                {
                    c[0] = S[Pos];
                    strcat(sNumber, c);  
                    Pos++;
                }
                if (Pos >= Len || !isDigit(S[Pos])) errflag = 1;
                else while ( Pos < Len && isDigit(S[Pos]))
                {
                    c[0] = S[Pos];
                    strcat(sNumber, c);  
                    Pos++;
                }
            }
        }
    }
    Pos--;
    if (errflag) return 0;
    else return atof(sNumber);
}

int getOperand()
{
    int code;
    switch(S[Pos])
    {
      case '(': code = 1;  break;
      case ')': code = 2;  break;
      case '+': code = 3;  break;
      case '-': code = 4;
                if (Pos < Len-1 &&
                    isDigit(S[Pos+1]) &&
                        (CurLex == 0 || CurLex == 1))
                {
                    Pos++;
                    Fvalue = -getNumber();
                    code = 7;
                }
                break;
      case '*': code = 5;  break;
      case '/': code = 6;  break;
      case '^': code = 31; break;
      default:  code = 0;
    }
    return code;
}

int getLex()
{
    int n;

    /* --- skip spaces */
    while ( Pos < Len && S[Pos] == ' ' ) Pos++;
    if ( Pos >= Len ) return 0;

    /* --- check for operand */
    n = getOperand();

    /* --- check for function/variable/number */
    if ( n == 0 )
    {
        if ( isLetter(S[Pos]) )
        {
            getToken();
            n = getMathFunc();
            if ( n == 0 ) n = getVariable();
        }
        else if ( isDigit(S[Pos]) )
        {
            n = 7;
            Fvalue = getNumber();
        }
    }
    Pos++;
    PrevLex = CurLex;
    CurLex = n;
    return n;
}

ExprTree* newNode()
{
    ExprTree* node;
    node = (ExprTree *) malloc(sizeof(ExprTree));
    if (!node) Err = 2;
    else
    {
        node->opcode = 0;
        node->ivar   = -1;
        node->fvalue = 0.;
        node->left   = NULL;
        node->right  = NULL;
    }
    return node;
}

ExprTree* getSingleOp(int *lex)
{
    int bracket;
    int opcode;
    ExprTree *left;
    ExprTree *right;
    ExprTree *node;

    /* --- open parenthesis, so continue to grow the tree */
    if ( *lex == 1 )
    {
        Bc++;
        left = getTree();
    }

    else
    {
        /* --- Error if not a singleton operand */
        if ( *lex < 7 || *lex == 9 || *lex > 30)
        {
            Err = 1;
            return NULL;
        }

        opcode = *lex;

        /* --- simple number or variable name */
        if ( *lex == 7 || *lex == 8 )
        {
            left = newNode();
            left->opcode = opcode;
            if ( *lex == 7 ) left->fvalue = Fvalue;
            if ( *lex == 8 ) left->ivar = Ivar;
        }

        /* --- function which must have a '(' after it */
        else
        {
            *lex = getLex();
            if ( *lex != 1 )
            {
               Err = 1;
               return NULL;
            }
            Bc++;
            left = newNode();
            left->left = getTree();
            left->opcode = opcode;
        }
    }   
    *lex = getLex();

    /* --- exponentiation */
    while ( *lex == 31 )
    {
        *lex = getLex();
        bracket = 0;
        if ( *lex == 1 )
        {
            bracket = 1;
            *lex = getLex();
        }
        if ( *lex != 7 )
        {
            Err = 1;
            return NULL;
        }
        right = newNode();
        right->opcode = *lex;
        right->fvalue = Fvalue;
        node = newNode();
        node->left = left;
        node->right = right;
        node->opcode = 31;
        left = node;
        if (bracket)
        {
            *lex = getLex();
            if ( *lex != 2 )
            {
                Err = 1;
                return NULL;
            }
        }
        *lex = getLex();
    }
    return left;
}

ExprTree* getOp(int *lex)
{
    int opcode;
    ExprTree* left;
    ExprTree* right;
    ExprTree* node;
    int neg = 0;

    *lex = getLex();
    if (PrevLex == 0 || PrevLex == 1)
    {
        if ( *lex == 4 )
        {
            neg = 1;
            *lex = getLex();
        }
        else if ( *lex == 3) *lex = getLex();
    }
    left = getSingleOp(lex);
    while ( *lex == 5 || *lex == 6 )
    {
        opcode = *lex;
        *lex = getLex();
        right = getSingleOp(lex);
        node = newNode();
        if (Err) return NULL;
        node->left = left;
        node->right = right;
        node->opcode = opcode;
        left = node;
    }
    if ( neg )
    {
        node = newNode();
        if (Err) return NULL;
        node->left = left;
        node->right = NULL;
        node->opcode = 9;
        left = node;
    }
    return left;
}

ExprTree* getTree()
{
    int       lex;
    int       opcode;
    ExprTree* left;
    ExprTree* right;
    ExprTree* node;

    left = getOp(&lex);
    for (;;)
    {
        if ( lex == 0 || lex == 2 )
        {
            if ( lex == 2 ) Bc--;
            break;
        }

        if (lex != 3 && lex != 4 )
        {
            Err = 1;
            break;
        }

        opcode = lex;
        right = getOp(&lex);
        node = newNode();
        if (Err) break;
        node->left = left;
        node->right = right;
        node->opcode = opcode;
        left = node;
    } 
    return left;
}

double c(ExprTree* t)
{
    double r;
    switch (t->opcode)
    {
      case 3:  r = c(t->left) + c(t->right); break;
      case 4:  r = c(t->left) - c(t->right); break;
      case 5:  r = c(t->left) * c(t->right); break;
      case 6:  r = c(t->left) / c(t->right); break;
      case 7:  r = t->fvalue;                break;
      case 8:  if (getVariableValue != NULL)
                   r = getVariableValue(t->ivar);
               else r = 0.0;                 break;
      case 9:  r = -c(t->left);              break;
      case 10: r = cos( c(t->left) );        break;
      case 11: r = sin( c(t->left) );        break;
      case 12: r = tan( c(t->left) );        break;
      case 13: r = 1.0/tan( c(t->left) );    break;
      case 14: r = fabs( c(t->left) );       break;
      case 15: r = c(t->left);
               if (r < 0.0) r = -1.0;
               else if (r > 0.0) r = 1.0;
               else r = 0.0;
               break;
      case 16: r = sqrt( c(t->left) );       break;
      case 17: r = log( c(t->left) );        break;
      case 27: r = log10( c(t->left) );      break;
      case 18: r = exp( c(t->left) );        break;
      case 19: r = asin( c(t->left) );       break;
      case 20: r = acos( c(t->left) );       break;
      case 21: r = atan( c(t->left) );       break;
      case 22: r = 1.5708-atan(c(t->left));  break;
      case 23: r = c(t->left);
               r = (exp(r)-exp(-r))/2.0;
               break;
      case 24: r = c(t->left);
               r = (exp(r)+exp(-r))/2.0;
               break;
      case 25: r = c(t->left);
               r = (exp(r)-exp(-r))/(exp(r)+exp(-r));
               break;
      case 26: r = c(t->left);
               r = (exp(r)+exp(-r))/(exp(r)-exp(-r));
               break;
      case 31: r = exp( c(t->right)*log( c(t->left) ) );
               break;
      default: r = 0.0;
    }
    return r;
}

double mathexpr_eval(ExprTree* tree, double (*getVal) (int))
{
    getVariableValue = getVal;
    return c(tree);
}

void mathexpr_delete(ExprTree* t)
{
    if (t)
    {
        if (t->left)  mathexpr_delete(t->left);
        if (t->right) mathexpr_delete(t->right);
        free(t);
    }
}

ExprTree* mathexpr_create(char *formula, int (*getVar) (char *))
{
    ExprTree* t;
    getVariableIndex = getVar;
    Err = 0;
    PrevLex = 0;
    CurLex = 0;
    S = formula;
    Len = strlen(S);
    Pos = 0;
    Bc = 0;
    t = getTree();
    if (Bc != 0 || Err > 0)
    {
        mathexpr_delete(t);
        return NULL;
    }
    return t;
}
