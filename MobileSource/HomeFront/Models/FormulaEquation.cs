using System;

namespace HomeFront.Models;


/// <summary>Faithful C# port of VB6 <c>clsEquation</c> (clsEquation.cls).
/// Recursive-descent parser over the same grammar and operator
/// precedence VB6 used, with the same built-in function library and the
/// same DB-lookup hooks. Extracted from FTakeoff (was a private nested
/// class) so the Formula Editor's Test button (HHM-211) can evaluate with
/// the SAME engine: the host supplies <paramref name="itemTotalLookup"/> for
/// the TotalOrderQty/TotalTakeoffQty/TotalCost aggregate functions — pass
/// null to stub them to 0, exactly what VB6's FFormulaTest.GetItemTotal
/// does (FFormulaTest.frm:173-175). Built per-evaluation; not reused, so it
/// holds no mutable solver state between rows.</summary>
public sealed class FormulaEquation
{
private readonly string _src;
private readonly Func<string, string, string, string, string, double>? _itemTotal;
private int _pos;

public FormulaEquation(string equation, Func<string, string, string, string, string, double>? itemTotalLookup = null)
{
    // VB6 normalizes line breaks and wraps the whole thing — we just
    // normalize whitespace; precedence is handled by the parser.
    _src = (equation ?? "").Replace('\r', ' ').Replace('\n', ' ');
    _itemTotal = itemTotalLookup;
}

    // Error strings mirror VB6 clsEquation verbatim (HHM-614): unknown tokens
    // raise "Undefined Function: …" (clsEquation.cls:441/473), structural
    // failures "Invalid Equation" (cls:478), parens "Unbalanced parenthesis"
    // (cls:1060). Hosts surface them under the "Error Evaluating Formula"
    // caption exactly like MFormulas.bas CalcFormula's handler.
    public double Solve()
    {
        _pos = 0;
        double result = ParseExpression();
        SkipWhitespace();
        if (_pos < _src.Length)
            throw new FormatException("Invalid Equation");
        return result;
    }

    // Grammar (lowest → highest precedence), mirroring clsEquation's
    // PRI_* ladder: OR/AND → comparisons → +,- → *,/,\,% (mod) → unary
    // minus → ^ → atoms (numbers, parens, functions, constants).
    private double ParseExpression() => ParseOr();

    private double ParseOr()
    {
        double left = ParseAnd();
        while (MatchKeyword("OR"))
        {
            double right = ParseAnd();
            left = (left != 0 || right != 0) ? 1 : 0;
        }
        return left;
    }

    private double ParseAnd()
    {
        double left = ParseComparison();
        while (MatchKeyword("AND"))
        {
            double right = ParseComparison();
            left = (left != 0 && right != 0) ? 1 : 0;
        }
        return left;
    }

    private double ParseComparison()
    {
        double left = ParseAddSub();
        while (true)
        {
            SkipWhitespace();
            if (TryConsume("<=")) { left = left <= ParseAddSub() ? 1 : 0; }
            else if (TryConsume(">=")) { left = left >= ParseAddSub() ? 1 : 0; }
            else if (TryConsume("<>")) { left = left != ParseAddSub() ? 1 : 0; }
            else if (TryConsume("<")) { left = left < ParseAddSub() ? 1 : 0; }
            else if (TryConsume(">")) { left = left > ParseAddSub() ? 1 : 0; }
            else if (TryConsume("=")) { left = left == ParseAddSub() ? 1 : 0; }
            else break;
        }
        return left;
    }

    private double ParseAddSub()
    {
        double left = ParseMulDiv();
        while (true)
        {
            SkipWhitespace();
            if (TryConsume("+")) left += ParseMulDiv();
            else if (TryConsume("-")) left -= ParseMulDiv();
            else break;
        }
        return left;
    }

    private double ParseMulDiv()
    {
        double left = ParseUnary();
        while (true)
        {
            SkipWhitespace();
            // " mod " is a word operator in VB6 (clsEquation.cls:833).
            if (MatchKeyword("mod")) { left = VbMod(left, ParseUnary()); continue; }
            if (_pos >= _src.Length) break;
            char c = _src[_pos];
            if (c == '*') { _pos++; left *= ParseUnary(); }
            else if (c == '/') { _pos++; left /= ParseUnary(); }
            else if (c == '\\') { _pos++; left = IntDiv(left, ParseUnary()); }
            else if (c == '%') { _pos++; left = VbMod(left, ParseUnary()); }
            else if (IsImplicitMultiplyStart(c)) { left *= ParseUnary(); } // 2(3), 2 x
            else break;
        }
        return left;
    }

    private double ParseUnary()
    {
        SkipWhitespace();
        if (_pos < _src.Length && _src[_pos] == '-') { _pos++; return -ParseUnary(); }
        if (_pos < _src.Length && _src[_pos] == '+') { _pos++; return ParseUnary(); }
        return ParsePower();
    }

    private double ParsePower()
    {
        double baseVal = ParseAtom();
        SkipWhitespace();
        if (_pos < _src.Length && _src[_pos] == '^')
        {
            _pos++;
            double exp = ParseUnary(); // right-associative
            return Math.Pow(baseVal, exp);
        }
        return baseVal;
    }

    // True when the next atom should be multiplied into the running
    // product without an explicit operator (digit-then-paren / identifier).
    private bool IsImplicitMultiplyStart(char c) =>
        c == '(' || char.IsLetter(c) || c == '_';

    private double ParseAtom()
    {
        SkipWhitespace();
        if (_pos >= _src.Length)
            throw new FormatException("Invalid Equation");

        char c = _src[_pos];

        if (c == '(')
        {
            _pos++;
            double v = ParseExpression();
            SkipWhitespace();
            if (_pos >= _src.Length || _src[_pos] != ')')
                throw new FormatException("Unbalanced parenthesis");
            _pos++;
            return v;
        }

        if (char.IsDigit(c) || c == '.')
            return ParseNumber();

        if (c == '"')
            throw new FormatException("String result is not a number.");

        if (char.IsLetter(c) || c == '_')
            return ParseIdentifierOrFunction();

        throw new FormatException("Invalid Equation");
    }

    private double ParseNumber()
    {
        int start = _pos;
        while (_pos < _src.Length && (char.IsDigit(_src[_pos]) || _src[_pos] == '.'))
            _pos++;
        var text = _src.Substring(start, _pos - start);
        return double.Parse(text, System.Globalization.CultureInfo.InvariantCulture);
    }

    private double ParseIdentifierOrFunction()
    {
        int start = _pos;
        while (_pos < _src.Length && (char.IsLetterOrDigit(_src[_pos]) || _src[_pos] == '_'))
            _pos++;
        string name = _src.Substring(start, _pos - start);
        SkipWhitespace();

        // Function call?
        if (_pos < _src.Length && _src[_pos] == '(')
        {
            _pos++;
            var args = ParseArgList();
            SkipWhitespace();
            if (_pos >= _src.Length || _src[_pos] != ')')
                throw new FormatException("Unbalanced parenthesis");
            _pos++;
            return CallFunction(name, args);
        }

        // Bare constant (clsEquation.cls:446-448). Unknown bare
        // identifiers are an error, exactly like VB6 (CDbl(Vars(name))
        // raised when the variable wasn't supplied — cls:473
        // "Undefined Function: " & v).
        switch (name.ToLowerInvariant())
        {
            case "pi": return Math.PI;
            case "e": return 2.718281828;
            case "rnd": return 0; // deterministic for takeoff; VB6 Rnd is non-deterministic
            default:
                throw new FormatException($"Undefined Function: {name}");
        }
    }

    private List<EqArg> ParseArgList()
    {
        var args = new List<EqArg>();
        SkipWhitespace();
        if (_pos < _src.Length && _src[_pos] == ')')
            return args; // zero-arg function, e.g. pi()
        while (true)
        {
            args.Add(ParseArg());
            SkipWhitespace();
            if (_pos < _src.Length && _src[_pos] == ',') { _pos++; continue; }
            break;
        }
        return args;
    }

    // An argument may be a numeric sub-expression OR a bare string
    // literal ("Community", a phase code, …) used by the DB-lookup
    // functions. We detect a leading quote and capture the raw string;
    // otherwise we parse a numeric expression.
    private EqArg ParseArg()
    {
        SkipWhitespace();
        if (_pos < _src.Length && _src[_pos] == '"')
        {
            _pos++;
            int start = _pos;
            while (_pos < _src.Length && _src[_pos] != '"') _pos++;
            string str = _src.Substring(start, _pos - start);
            if (_pos < _src.Length) _pos++; // closing quote
            return EqArg.FromString(str);
        }
        return EqArg.FromNumber(ParseExpression());
    }

    private double CallFunction(string name, List<EqArg> a)
    {
        double P(int i) => i < a.Count ? a[i].Num : 0;
        string S(int i) => i < a.Count ? a[i].Str : "";

        switch (name.ToLowerInvariant())
        {
            case "abs": return Math.Abs(P(0));
            case "int": return Math.Floor(P(0));         // VB6 Int = floor toward -inf
            case "fix": return Math.Truncate(P(0));       // VB6 Fix toward zero
            case "sgn": return Math.Sign(P(0));
            case "sqrt": return Math.Sqrt(P(0));
            case "square": return P(0) * P(0);
            case "exp": return Math.Exp(P(0));
            case "ln":
            case "log": return Math.Log(P(0));
            case "log10": return Math.Log10(P(0));
            case "log2": return Math.Log(P(0)) / Math.Log(2);
            case "logn": return Math.Log(P(0)) / Math.Log(P(1));
            case "power": return Math.Pow(P(0), P(1));
            case "mod": return VbMod(P(0), P(1));
            case "pi": return Math.PI;
            case "floor": return VbFloor(P(0));
            case "ceiling": return VbCeiling(P(0));
            case "min": return VbMin(a);
            case "max": return VbMax(a);
            case "sum": return a.Sum(x => x.Num);
            case "avg": return a.Count == 0 ? 0 : a.Average(x => x.Num);
            case "round": return VbRound(P(0), P(1));
            case "trunc": return VbTrunc(P(0), P(1));
            case "roundto": return VbRoundTo(P(0), P(1), P(2));
            case "if": return P(0) != 0 ? P(1) : P(2);

            // ── DB-aggregate lookups (clsEquation.cls:413-415) → host ──
            case "totalorderqty": return _itemTotal?.Invoke("OrderQty", S(0), S(1), S(2), S(3)) ?? 0;
            case "totaltakeoffqty": return _itemTotal?.Invoke("TakeoffQty", S(0), S(1), S(2), S(3)) ?? 0;
            case "totalcost": return _itemTotal?.Invoke("Cost", S(0), S(1), S(2), S(3)) ?? 0;

            default:
                // VB6 clsEquation.cls:441 — "Undefined Function: " & v & ")",
                // where v carries the token's opening paren → "name()".
                throw new FormatException($"Undefined Function: {name}()");
        }
    }

    // ── VB6 numeric helpers (clsEquation.cls) ──
    private static double VbMod(double a, double b)
    {
        // VB6 Mod operates on rounded longs.
        long la = (long)Math.Round(a);
        long lb = (long)Math.Round(b);
        return lb == 0 ? 0 : la % lb;
    }

    private static double IntDiv(double a, double b)
    {
        long lb = (long)Math.Round(b);
        return lb == 0 ? 0 : (long)Math.Round(a) / lb;
    }

    private static double VbFloor(double n)
    {
        if (n >= 0) return Math.Truncate(n);          // Fix toward zero (clsEquation.cls:1270)
        return n == Math.Truncate(n) ? n : Math.Truncate(n) - 1;
    }

    private static double VbCeiling(double n)
    {
        if (n >= 0) return n == Math.Truncate(n) ? n : Math.Truncate(n) + 1;
        return Math.Truncate(n);                        // clsEquation.cls:1256-1265
    }

    // clsEquation.cls:1113 trunc: keep N decimal places, truncating.
    private static double VbTrunc(double number, double decimalPlaces)
    {
        double dp = -1 * Math.Truncate(decimalPlaces);
        double scale = Math.Pow(10, dp);
        return Math.Truncate(number / scale) * scale;
    }

    // clsEquation.cls:1118 round wrapper (decimalPlaces shift + banker's).
    private static double VbRound(double number, double decimalPlaces)
    {
        int dp = (int)Math.Truncate(decimalPlaces);
        if (dp < 0) dp = 0;
        return Math.Round(number, Math.Min(dp, 15), MidpointRounding.ToEven);
    }

    // clsEquation.cls:1129 RoundTo(number, method, significance):
    // method 1 = up, 0 = closest, -1 = down to the nearest 'significance'.
    private static double VbRoundTo(double number, double method, double significance)
    {
        if (significance == 0 || number == significance) return number;
        double ratio = number / significance;
        if (ratio == Math.Truncate(ratio)) return number;

        if (method > 0) return Math.Ceiling(ratio) * significance;          // up
        if (method < 0) return Math.Floor(ratio) * significance;            // down
        // closest
        double down = Math.Floor(ratio) * significance;
        double up = Math.Ceiling(ratio) * significance;
        return Math.Abs(number - down) < Math.Abs(number - up) ? down : up;
    }

    private static double VbMin(List<EqArg> a)
    {
        double v = a.Count > 0 ? a[0].Num : 0;
        for (int i = 1; i < a.Count; i++) v = a[i].Num < v ? a[i].Num : v;
        return v;
    }

    private static double VbMax(List<EqArg> a)
    {
        double v = a.Count > 0 ? a[0].Num : 0;
        for (int i = 1; i < a.Count; i++) v = a[i].Num > v ? a[i].Num : v;
        return v;
    }

    // ── scanner helpers ──
    private void SkipWhitespace()
    {
        while (_pos < _src.Length && char.IsWhiteSpace(_src[_pos])) _pos++;
    }

    private bool TryConsume(string token)
    {
        if (string.CompareOrdinal(_src, _pos, token, 0, token.Length) == 0
            && _pos + token.Length <= _src.Length)
        {
            _pos += token.Length;
            return true;
        }
        return false;
    }

    // Matches a whole-word keyword (AND/OR/mod) case-insensitively,
    // requiring a non-identifier boundary on both sides so a variable
    // named "ANDREW" or "MODEL" isn't mistaken for an operator.
    private bool MatchKeyword(string kw)
    {
        SkipWhitespace();
        if (_pos + kw.Length > _src.Length) return false;
        if (string.Compare(_src, _pos, kw, 0, kw.Length, StringComparison.OrdinalIgnoreCase) != 0)
            return false;
        // left boundary: previous non-space char already consumed; right
        // boundary must not continue an identifier.
        int after = _pos + kw.Length;
        if (after < _src.Length && (char.IsLetterOrDigit(_src[after]) || _src[after] == '_'))
            return false;
        _pos = after;
        return true;
    }

    private readonly struct EqArg
    {
        public double Num { get; }
        public string Str { get; }
        private EqArg(double n, string s) { Num = n; Str = s; }
        public static EqArg FromNumber(double n) =>
            new(n, n.ToString(System.Globalization.CultureInfo.InvariantCulture));
        public static EqArg FromString(string s)
        {
            double.TryParse(s, System.Globalization.NumberStyles.Any,
                System.Globalization.CultureInfo.InvariantCulture, out var n);
            return new EqArg(n, s);
        }
    }
}
