using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TaxRefund
{
    public class ReportData
    {
        public decimal TotalGoodsValue { get; set; }
        public decimal AccumulatedGoodsValue { get; set; }
        public decimal TotalVATRefund { get; set; }
        public decimal AccumulatedVATRefund { get; set; }
        public decimal TotalServiceFee { get; set; }
        public decimal AccumulatedServiceFee { get; set; }
        public decimal PassengerTurns { get; set; }
        public decimal AccumulatedPassengerTurns { get; set; }

        // Previous year data
        public decimal PreviousYearGoodsValue { get; set; }
        public decimal PreviousYearVATRefund { get; set; }
        public decimal PreviousYearPassengerTurns { get; set; }

        // Comparison percentages
        public decimal GoodsValueComparison => CalculatePercentage(TotalGoodsValue, PreviousYearGoodsValue);
        public decimal VATRefundComparison => CalculatePercentage(TotalVATRefund, PreviousYearVATRefund);
        public decimal PassengerTurnsComparison => CalculatePercentage(PassengerTurns, PreviousYearPassengerTurns);

        private decimal CalculatePercentage(decimal current, decimal previous)
        {
            if (previous == 0) return 0;
            return (current / previous) * 100;
        }
    }
}
