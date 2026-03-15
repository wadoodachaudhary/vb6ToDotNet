namespace VB6Migration;

public static class Calculator
{
    public static double Add(double a, double b)
    {
        return a + b;
    }

    public static double Subtract(double a, double b)
    {
        return a - b;
    }

    public static double Multiply(double a, double b)
    {
        return a * b;
    }

    public static double Divide(double a, double b)
    {
        if (b == 0)
        {
            throw new DivideByZeroException("Division by zero");
        }
        return a / b;
    }

    public static long Factorial(long n)
    {
        if (n < 0)
        {
            throw new ArgumentException("Negative number", nameof(n));
        }

        long result = 1;
        for (long i = 2; i <= n; i++)
        {
            result = result * i;
        }

        return result;
    }

    public static bool IsPrime(long n)
    {
        if (n <= 1)
        {
            return false;
        }

        if (n <= 3)
        {
            return true;
        }

        if (n % 2 == 0 || n % 3 == 0)
        {
            return false;
        }

        long i = 5;
        while (i * i <= n)
        {
            if (n % i == 0 || n % (i + 2) == 0)
            {
                return false;
            }
            i += 6;
        }

        return true;
    }
}