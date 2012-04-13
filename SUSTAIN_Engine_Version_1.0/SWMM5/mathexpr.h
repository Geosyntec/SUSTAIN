//-----------------------------------------------------------------------------
//  mathexpr.h
//
// Header file for Math Expression module mathexpr.c.
//
// Last Updated: 6/16/03
//
//-----------------------------------------------------------------------------

/* Operand codes:
   1 = (
   2 = )
   3 = +
   4 = - (subtraction)
   5 = *
   6 = /
   7 = number
   8 = variable
   9 = - (negative)
  10 = cos
  11 = sin
  12 = tan
  13 = cot
  14 = abs
  15 = sgn
  16 = sqrt
  17 = log
  18 = exp
  19 = asin
  20 = acos
  21 = atan
  22 = acot
  23 = sinh
  24 = cosh
  25 = tanh
  26 = coth
  27 = log10
  31 = ^
*/

//-----------------------------------------
// Node in a tokenized math expression tree
//-----------------------------------------
struct ExprNode
{
    int    opcode;                // operator code
    int    ivar;                  // variable index
    double fvalue;                // numerical value
    struct ExprNode* left;        // left sub-tree of tokenized formula
    struct ExprNode* right;       // right sub-tree of tokenized formula
};

typedef struct ExprNode ExprTree;

// Create a tokenized expression tree from a string s
ExprTree* mathexpr_create(char* s, int (*getVar) (char *));

// Evaluate an expression tree
double mathexpr_eval(ExprTree* t, double (*getVal) (int));

// Delete a tokenized expression tree
void  mathexpr_delete(ExprTree* t);
