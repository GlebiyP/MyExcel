using Antlr4.Runtime.Misc;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MyExcel
{
    class RecursiveException : Exception
    {
    }
    class DivideZeroException : Exception
    {
    }   
    class LabCalcVisitor : LabCalcBaseVisitor<double>
    {
        public Dictionary<string, Cell> tableIdentifier;
        public string CurrentCell;

        public LabCalcVisitor(Dictionary<string, Cell> dictionary, string cell)
        {
            tableIdentifier = dictionary;
            CurrentCell = cell;
        }

        public override double VisitCompileUnit(LabCalcParser.CompileUnitContext context)
        {
            return Visit(context.expression());
        }

        public override double VisitMinusExpr(LabCalcParser.MinusExprContext context)
        {
            var num = WalkLeft(context);
            return 0 - num;
        }

        public override double VisitNumberExpr(LabCalcParser.NumberExprContext context)
        {
            var result = double.Parse(context.GetText());
            Debug.WriteLine(result);
            return result;
        }

        public override double VisitIdentifierExpr(LabCalcParser.IdentifierExprContext context)
        {
            var result = context.GetText();
            if(!Recursive(CurrentCell, result))
            {
                if (!tableIdentifier[CurrentCell].CellReference.Contains(result))
                    tableIdentifier[CurrentCell].CellReference.Add(result);
                string StrValue = tableIdentifier[result].Val;
                if (StrValue == "Error")
                    throw new Exception();
                if (StrValue == "ErrorCycle")
                    throw new RecursiveException();
                if (StrValue == "ErrorDivZero")
                    throw new DivideByZeroException();
                double value = Convert.ToDouble(StrValue);
                Debug.WriteLine(value);
                return value;
            }
            else
            {
                throw new RecursiveException();
            }
        }

        public bool Recursive(string CurrentCell, string ReferedCell)
        {
            if (ReferedCell == CurrentCell)
                return true;
            else if (tableIdentifier[ReferedCell].CellReference.Count != 0)
            {
                foreach (string i in tableIdentifier[ReferedCell].CellReference)
                {
                    if (Recursive(CurrentCell, i)) return true;
                }
                return false;
            }
            else return false;
        }

        public override double VisitParenthesizedExpr(LabCalcParser.ParenthesizedExprContext context)
        {
            return Visit(context.expression());
        }

        public override double VisitExponentialExpr(LabCalcParser.ExponentialExprContext context)
        {
            var left = WalkLeft(context);
            var right = WalkRight(context);
            if (left == 0 && right < 0)
                throw new Exception();
            Debug.WriteLine("{0} ^ {1}", left, right);
            double res = System.Math.Pow(left, right);
            if (res == double.PositiveInfinity || res == double.NegativeInfinity)
                throw new Exception();
            return res;
        }

        public override double VisitAdditiveExpr(LabCalcParser.AdditiveExprContext context)
        {
            var left = WalkLeft(context);
            var right = WalkRight(context);
            if (context.operatorToken.Type == LabCalcLexer.ADD)
            {
                Debug.WriteLine("{0} + {1}", left, right);
                return left + right;
            }
            else //LabCalculatorLexer.SUBTRACT
            {
                Debug.WriteLine("{0} - {1}", left, right);
                return left - right;
            }
        }

        public override double VisitMultiplicativeExpr(LabCalcParser.MultiplicativeExprContext context)
        {
            var left = WalkLeft(context);
            var right = WalkRight(context);
            if (context.operatorToken.Type == LabCalcLexer.MULTIPLY)
            {
                Debug.WriteLine("{0} * {1}", left, right);
                double res = left * right;
                if (res == double.PositiveInfinity || res == double.NegativeInfinity)
                    throw new Exception();
                return res;
            }
            else //LabCalculatorLexer.DIVIDE
            {
                Debug.WriteLine("{0} / {1}", left, right);
                if (right + 1 == 0) throw new DivideZeroException();
                double res = left / right;
                if (res == double.PositiveInfinity || res == double.NegativeInfinity)
                    throw new Exception();
                return res;
            }
        }

        public override double VisitCompareExpr(LabCalcParser.CompareExprContext context)
        {
            var left = WalkLeft(context);
            var right = WalkRight(context);
            if (context.operatorToken.Type == LabCalcLexer.MLESS)
            {
                Debug.WriteLine("{0} < {1}", left, right);
                return Convert.ToDouble(left < right);
            }
            else if(context.operatorToken.Type==LabCalcLexer.MMORE)
            {
                Debug.WriteLine("{0} > {1}", left, right);
                return Convert.ToDouble(left > right);
            }
            else if(context.operatorToken.Type==LabCalcLexer.LESSOREQUAL)
            {
                Debug.WriteLine("{0} <= {1}", left, right);
                return Convert.ToDouble(left <= right);
            }
            else if(context.operatorToken.Type==LabCalcLexer.MOREOREQUAL)
            {
                Debug.WriteLine("{0} >= {1}", left, right);
                return Convert.ToDouble(left >= right);
            }
            else if(context.operatorToken.Type==LabCalcLexer.ISEQUAL)
            {
                Debug.WriteLine("{0} == {1}", left, right);
                return Convert.ToDouble(left == right);
            }
            else
            {
                //NOTEQUAL
                Debug.WriteLine("{0} <> {1}", left, right);
                return Convert.ToDouble(left != right);
            }
        }

        public override double VisitNotExpr([NotNull] LabCalcParser.NotExprContext context)
        {
            var left = WalkLeft(context);
            if (left == 0)
            {
                return 1;
            }
            return 0;
        }

        public override double VisitMminExpr([NotNull] LabCalcParser.MminExprContext context)
        {
            double minValue = Double.PositiveInfinity;

            foreach (var child in context.paramlist.children.OfType<LabCalcParser.ExpressionContext>()) 
            {
                double childValue = this.Visit(child);
                if (childValue < minValue)
                {
                    minValue = childValue;
                }
            }
            return minValue;
        }

        public override double VisitMmaxExpr([NotNull] LabCalcParser.MmaxExprContext context)
        {
            double maxValue = Double.NegativeInfinity;

            foreach (var child in context.paramlist.children.OfType<LabCalcParser.ExpressionContext>())
            {
                double childValue = this.Visit(child);
                if (childValue > maxValue)
                {
                    maxValue = childValue;
                }
            }
            return maxValue;
        }

        private double WalkLeft(LabCalcParser.ExpressionContext context)
        {
            return Visit(context.GetRuleContext<LabCalcParser.ExpressionContext>(0));
        }

        private double WalkRight(LabCalcParser.ExpressionContext context)
        {
            return Visit(context.GetRuleContext<LabCalcParser.ExpressionContext>(1));
        }
    }
}
