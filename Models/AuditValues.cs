using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelProcessor.Models
{
    public class AuditValues
    {
        public string date { get; set; }
        public string bid { get; set; }
        public string locale { get; set; }
        public string contractEntity { get; set; }
        public string collectEntity { get; set; }
        public string contractCurrency { get; set; }
        public string collectCurrency { get; set; }
        public decimal commissionRevenue { get; set; }
        public decimal transactionFee { get; set; }
        public decimal premium { get; set; }
        public decimal discount { get; set; }
        public decimal coupon { get; set; }
        public decimal redeemedPoints { get; set; }
        public decimal uniqueCode { get; set; }
        public decimal installmentFee { get; set; }
        public decimal deliveryFee { get; set; }
        public decimal invoiceAmount { get; set; }
        public decimal refundFee { get; set; }
        public decimal rescheduleFee { get; set; }
        public decimal rebookCost { get; set; }
        public AuditValues()
        {
            date = "01/01/1901";
            bid = "00000000";
            locale = "Locale";
            contractEntity = "ID02";
            collectEntity = "ID02";
            contractCurrency = "IDR";
            collectCurrency = "IDR";
            commissionRevenue = 0;
            transactionFee = 0;
            premium = 0;
            discount = 0;
            coupon = 0;
            redeemedPoints = 0;
            uniqueCode = 0;
            installmentFee = 0;
            deliveryFee = 0;
            invoiceAmount = 0;
            refundFee = 0;
            rescheduleFee = 0;
            rebookCost = 0;
        }
    }
}
