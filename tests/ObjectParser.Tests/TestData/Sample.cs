using System;
using System.Collections.Generic;

namespace Billing.Core
{
    /* Billing types used by the samples. */
    public interface IShape
    {
        double Area();
    }

    public enum Color { Red, Green, Blue }

    // An invoice with line items.
    public class Invoice : IDisposable
    {
        private readonly List<decimal> _lines = new List<decimal>();

        public Invoice(string number)
        {
            Number = number;
        }

        public string Number { get; }

        #region Totals
        public void Add(decimal amount, int qty)
        {
            if (qty > 0)
            {
                _lines.Add(amount * qty);
            }
        }

        public decimal Total()
        {
            decimal sum = 0;
            foreach (var line in _lines) { sum += line; }
            return sum;
        }
        #endregion

        public void Dispose() { _lines.Clear(); }
    }

    public class InvoiceLine
    {
        public decimal Amount { get; set; }
    }

    public struct Point
    {
        public int X;
        public int Y;
    }
}
