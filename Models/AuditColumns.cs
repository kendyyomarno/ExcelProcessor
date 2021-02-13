using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelProcessor.Models
{
    public class AuditColumns
    {
        public string date { get; set; }
        public string bid { get; set; }
        public string locale { get; set; }
        public string contractEntity { get; set; }
        public string collectEntity { get; set; }
        public string contractCurrency { get; set; }
        public string collectCurrency { get; set; }
        public string commissionRevenue { get; set; }
        public string transactionFee { get; set; }
        public string premium { get; set; }
        public string discount { get; set; }
        public string coupon { get; set; }
        public string redeemedPoints { get; set; }
        public string uniqueCode { get; set; }
        public string installmentFee { get; set; }
        public string deliveryFee { get; set; }
        public string invoiceAmount { get; set; }
        public string refundFee { get; set; }
        public string rescheduleFee { get; set; }
        public string rebookCost { get; set; }
        public AuditColumns()
        {
            date = "Non Refundable Date,Booking Issue Date";
            bid = "Booking ID";
            locale = "Locale";
            contractEntity = "Contract Entity,Inventory Owner";
            collectEntity = "Collecting Entity";
            contractCurrency = "Contract Currency";
            collectCurrency = "Collecting Currency";
            commissionRevenue = "Gross Commission";
            transactionFee = "Service Fee";
            premium = "Discount/Premium";
            discount = "Discount/Premium";
            coupon = "Coupon";
            redeemedPoints = "Point Redemption";
            uniqueCode = "Unique Code";
            installmentFee = "Installment Fee";
            deliveryFee = "Delivery Fee";
            invoiceAmount = "Customer Invoice";
            refundFee = "Refund Fee";
            rescheduleFee = "Reschedule Fee";
            rebookCost = "Rebook Cost";
        }
    }
}
